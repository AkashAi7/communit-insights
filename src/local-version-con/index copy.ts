// src/index.ts

import * as dotenv from 'dotenv';
dotenv.config(); // Load environment variables from .env file

import express from 'express';
import bodyParser from 'body-parser'; // For parsing request bodies



import { App } from "@microsoft/teams-ai";
import { analyzeFeedbackTool } from "./mcpPlugin"; // ✅ import new tool

const app = new App({
  ai: {
    tools: [analyzeFeedbackTool] // ✅ register the tool here
  },
});


// Teams AI App core imports
import { App } from '@microsoft/teams.apps';
//import App  from '@microsoft/teams.ai'; // <--- FIXED: Import App as default export
import { AI } from '@microsoft/teams-ai'; // <--- FIXED: Import AI directly
import { DevtoolsPlugin } from '@microsoft/teams.dev';

// MCP Plugin (Still instantiate it for its tool definition, but not for hosting its route via App)
import { McpPlugin } from '@microsoft/teams.mcp';
import { z } from 'zod'; // For schema validation

// AI Model for analysis
import { ChatPrompt } from '@microsoft/teams.ai';
import { OpenAIChatModel } from '@microsoft/teams.openai';

// --- Configuration Constants ---
const TEAMS_BOT_PORT = +(process.env.PORT || 3976); // Main bot port
const MCP_SERVER_PORT = 3975; // Dedicated port for MCP Ingestion Server
const DEVTOOLS_PORT = 3977; // Devtools will always run on this separate port

// --- Global Data Store for Proactive Messaging (in-memory for this example) ---
const userToConversationId = new Map<string, string>();

// --- Initialize AI Model for the Analysis Tool ---
const openaiModel = new OpenAIChatModel({
  apiKey: process.env.AZURE_OPENAI_API_KEY!,
  endpoint: process.env.AZURE_OPENAI_ENDPOINT!,
  apiVersion: process.env.AZURE_OPENAI_API_VERSION!,
  model: process.env.AZURE_OPENAI_DEPLOYMENT_NAME!,
});


// --- NEW: Initialize AI instance and pass the model to it ---
const ai = new AI(openaiModel); // <--- FIXED: Use direct AI import

// --- Define Zod Schemas for MCP Tool Input/Output ---
const FeedbackItemSchema = z.object({
    id: z.number().describe('Unique ID of the feedback item'),
    text: z.string().describe('The content of the feedback (title + body)'),
    source: z.string().describe('Source of the feedback (e.g., Stack Overflow, GitHub Issues)'),
    url: z.string().url().describe('URL to the original feedback item').optional()
});

const AnalyzeFeedbackInputSchema = z.object({
    feedback: z.array(FeedbackItemSchema).describe('An array of developer feedback items to analyze.')
}).describe('Input for the analyzeFeedback MCP tool.');

const AnalysisObjectSchema = z.object({
    painPoints: z.array(z.string()).describe('Array of identified pain points.'),
    summary: z.string().describe('A summary of the feedback.'),
    priority: z.enum(['low', 'medium', 'high']).describe('Priority of the feedback.')
});

const AnalysisErrorSchema = z.object({
    error: z.string().describe('Error message'),
    rawOutput: z.any().optional().describe('Raw output from AI if available')
});

const AnalysisResultSchema = z.object({
    originalId: z.number().describe('Original ID of the feedback item'),
    originalSource: z.string().describe('Original source of the feedback'),
    originalUrl: z.string().url().describe('URL to the original feedback item').optional(),
    analysis: z.union([AnalysisObjectSchema, AnalysisErrorSchema])
});

const AnalyzeFeedbackOutputSchema = z.object({
    analyzedResults: z.array(AnalysisResultSchema).describe('Array of analyzed feedback results.')
});


// --- AI-Driven Feedback/Pain Point Extraction Logic (reusable handler) ---
const analyzeFeedbackToolHandler = async (
    args: { [x: string]: any },
    _extra: any // You can type this as needed
) => {
    const feedback = args.feedback;
    console.log(`AI analysis received ${feedback.length} feedback items.`);
    const analyzedResults: z.infer<typeof AnalysisResultSchema>[] = [];

    for (const item of feedback) {
        console.log(`Processing feedback item ID: ${item.id} from ${item.source}`);
        try {
            const prompt = new ChatPrompt({
                instructions: `Analyze the following developer feedback to identify key pain points, recurring issues, and actionable insights.
                              Output a JSON object with 'painPoints' (array of strings), 'summary' (string), and 'priority' (low, medium, high).
                              Ensure the output is always a valid JSON string.`,
                model: openaiModel,
            });

            const analysisResult = await prompt.send(item.text);

            if (analysisResult.content) {
                try {
                    let jsonString = analysisResult.content;
                    if (jsonString.startsWith('```json')) {
                        jsonString = jsonString.substring(7);
                    }
                    if (jsonString.endsWith('```')) {
                        jsonString = jsonString.substring(0, jsonString.length - 3);
                    }
                    jsonString = jsonString.trim();

                    const parsedAnalysis = JSON.parse(jsonString);
                    analyzedResults.push({
                        originalId: item.id,
                        originalSource: item.source,
                        originalUrl: item.url,
                        analysis: parsedAnalysis
                    });
                    console.log(`Successfully analyzed item ${item.id}. Priority: ${parsedAnalysis.priority}`);
                } catch (parseError) {
                    console.error(`Error parsing AI output for item ${item.id}:`, parseError);
                    analyzedResults.push({
                        originalId: item.id,
                        originalSource: item.source,
                        originalUrl: item.url,
                        analysis: { error: "AI output not valid JSON", rawOutput: analysisResult.content }
                    });
                }
            } else {
                console.warn(`AI returned no content for item ${item.id}.`);
                analyzedResults.push({
                    originalId: item.id,
                    originalSource: item.source,
                    originalUrl: item.url,
                    analysis: { error: "No AI content" }
                });
            }
        } catch (aiError: any) {
            console.error(`AI analysis failed for item ${item.id}:`, aiError.message || aiError);
            analyzedResults.push({
                originalId: item.id,
                originalSource: item.source,
                originalUrl: item.url,
                analysis: { error: aiError.message || "AI analysis failed" }
            });
        }
    }
    console.log(`Completed analysis for ${analyzedResults.length} items.`);
    return { analyzedResults };
};

// --- Instantiate McpPlugin for tool definition (used by bot's `invoke`) ---
const mcpServerPlugin = new McpPlugin({
  name: 'developerFeedbackAgent',
  description: 'An MCP server that analyzes developer feedback from various sources.',
  inspector: `http://localhost:${DEVTOOLS_PORT}/devtools`,
});

// Manually register the tool with the plugin's internal registry
mcpServerPlugin.tool(
    'analyzeFeedback',
    'Analyzes an array of developer feedback items and returns AI-driven insights.',
    AnalyzeFeedbackInputSchema.shape,
    AnalyzeFeedbackOutputSchema.shape,
    analyzeFeedbackToolHandler
);

// --- Create a dedicated Express app for the MCP Ingestion Endpoint ---
const mcpExpressApp = express();
mcpExpressApp.use(bodyParser.json({ limit: '50mb' })); // Explicitly set limit for this app
mcpExpressApp.use(bodyParser.urlencoded({ limit: '50mb', extended: true }));

// --- Store latest analysis results and user state ---
let latestAnalysis: any[] = [];
const userToInsightIndex = new Map<string, number>();
const userToSearchResults = new Map<string, { matches: any[], idx: number }>();

mcpExpressApp.post('/api/mcp/ingest', async (req, res) => {
    try {
        console.log(`Received POST to /api/mcp/ingest on port ${MCP_SERVER_PORT} at ${new Date().toISOString()}`);
        const validationResult = AnalyzeFeedbackInputSchema.safeParse(req.body);

        if (!validationResult.success) {
            console.error('MCP Input Validation Error:', validationResult.error);
            return res.status(400).json({ error: 'Invalid input schema', details: validationResult.error.issues });
        }

        const input = validationResult.data;
        const output = await analyzeFeedbackToolHandler({ feedback: input.feedback });
        latestAnalysis = output.analyzedResults || [];
        userToInsightIndex.clear();
        userToSearchResults.clear();
        const outputValidationResult = AnalyzeFeedbackOutputSchema.safeParse(output);
        if (!outputValidationResult.success) {
            console.error('MCP Output Validation Error:', outputValidationResult.error);
            return res.status(500).json({ error: 'AI analysis returned invalid output', details: outputValidationResult.error.issues, rawOutput: output });
        }

        return res.status(200).json(outputValidationResult.data);

    } catch (error: any) {
        console.error('Error handling /api/mcp/ingest request on dedicated server:', error);
        return res.status(500).json({ error: 'Internal server error processing MCP request', details: error.message });
    }
});


// --- Community Insider Bot in Teams (using Teams AI Library v2) ---
// const teamsApp =  App({
//   // IMPORTANT: Pass the 'ai' instance to the App constructor
//   ai: ai, // <--- NEW: Added 'ai' property
//   plugins: [
//     new DevtoolsPlugin(), // Devtools will still run on 3977
//     mcpServerPlugin, // This registers the tool with the App.
//   ]
// });

const teamsApp = new App({
  ai,
  plugins: [
    new DevtoolsPlugin(),
    mcpServerPlugin, 

  ],
});




// --- Helper function to create an Adaptive Card for displaying analysis ---
function createFeedbackAnalysisCard(analysis: z.infer<typeof AnalysisResultSchema>['analysis']) {
    const card: any = {
        type: "AdaptiveCard",
        $schema: "http://adaptivecards.io/schemas/adaptiveCard.json",
        version: "1.3",
        body: [
            {
                type: "TextBlock",
                text: "Feedback Analysis",
                wrap: true,
                size: "Large",
                weight: "Bolder"
            },
            {
                type: "TextBlock",
                text: `Priority: **${analysis.priority.toUpperCase()}**`,
                wrap: true,
                color: analysis.priority === 'high' ? 'Attention' : (analysis.priority === 'medium' ? 'Warning' : 'Good')
            },
            {
                type: "TextBlock",
                text: "Summary:",
                wrap: true,
                weight: "Bolder",
                spacing: "Medium"
            },
            {
                type: "TextBlock",
                text: analysis.summary,
                wrap: true
            },
            {
                type: "TextBlock",
                text: "Key Pain Points:",
                wrap: true,
                weight: "Bolder",
                spacing: "Medium"
            }
        ]
    };

    if (analysis.painPoints && analysis.painPoints.length > 0) {
        card.body.push({
            type: "Container",
            items: analysis.painPoints.map(point => ({
                type: "TextBlock",
                text: `- ${point}`,
                wrap: true
            }))
        });
    } else {
        card.body.push({
            type: "TextBlock",
            text: "No specific pain points identified.",
            wrap: true,
            isSubtle: true
        });
    }

    return card;
}


// --- Conversational Message handler for the Teams Bot. ---
// Access properties directly from the 'context' object
teamsApp.on('message', async (context) => {
  await context.send({ type: 'typing' });

  const text = context.activity.text?.trim();
  const lowerText = text?.toLowerCase() || "";
  const userKey = context.activity.conversation.id;

  // --- NEW: Mission statement/help command ---
  if (lowerText === '/about' || lowerText === '/help') {
    await context.send(
      "I am your Teams Platform Community Insights Bot.\n\n" +
      "I analyze and surface actionable insights from developer feedback across forums, " +
      "helping the Ops team proactively identify and prioritize developer-expressed pain points for the Teams Platform in various dev community forums.\n\n" +
      "Commands:\n" +
      "- `/show_insights` — View feedback insights one by one\n" +
      "- `/next_insight` — Next insight\n" +
      "- `/latest_insight` — Most recent insight\n" +
      "- `/search_insights <keyword>` — Search insights by topic or pain point\n" +
      "- `/next_search_result` — Next search result\n" +
      "- `/ask_about_current <your question>` — Ask about the currently displayed card"
    );
    return;
  }

  // --- Show the most recent card ---
  if (lowerText === '/latest_insight') {
    if (!latestAnalysis.length) {
      await context.send('No analysis results are available yet.');
      return;
    }
    const idx = latestAnalysis.length - 1;
    userToInsightIndex.set(userKey, idx);
    const result = latestAnalysis[idx];
    const card = createFeedbackAnalysisCard(result.analysis);
    await context.send({
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: card
      }]
    });
    await context.send('This is the most recent insight. You can ask follow-up questions using `/ask_about_current <your question>`.');
    return;
  }

  // --- Search insights by keyword (optimized for topic/issue search) ---
  if (lowerText.startsWith('/search_insights')) {
    const keyword = text?.slice('/search_insights'.length).trim().toLowerCase();
    if (!keyword) {
      await context.send('Please provide a keyword to search. Example: `/search_insights bot`');
      return;
    }
    if (!latestAnalysis.length) {
      await context.send('No analysis results are available yet.');
      return;
    }
    // Improved: Search in summary, painPoints, and also in the original feedback text/source
    const matches = latestAnalysis.filter(r => {
      const a = r.analysis;
      if ('error' in a) return false;
      const feedbackText = (r.originalSource || '') + ' ' + (r.originalUrl || '');
      return (
        (a.summary && a.summary.toLowerCase().includes(keyword)) ||
        (a.painPoints && a.painPoints.some((p: string) => p.toLowerCase().includes(keyword))) ||
        (feedbackText && feedbackText.toLowerCase().includes(keyword))
      );
    });
    if (!matches.length) {
      await context.send(`No insights found matching "${keyword}".`);
      return;
    }
    userToSearchResults.set(userKey, { matches, idx: 0 });
    const card = createFeedbackAnalysisCard(matches[0].analysis);
    await context.send({
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: card
      }]
    });
    // --- Optimized: Send concise summary and priority for quick context ---
    const a = matches[0].analysis;
    if (!('error' in a)) {
      let msg = `**Summary:** ${a.summary}\n**Priority:** ${a.priority.toUpperCase()}\n**Pain Points:**\n`;
      if (a.painPoints && a.painPoints.length > 0) {
        msg += a.painPoints.map((p: string) => `- ${p}`).join('\n');
      } else {
        msg += 'None identified.';
      }
      await context.send(msg);
    }
    if (matches.length > 1) {
      await context.send(`Found ${matches.length} results. Type /next_search_result to see the next match.`);
    }
    // Track this as the current card for follow-up
    userToInsightIndex.set(userKey, latestAnalysis.indexOf(matches[0]));
    return;
  }

  // --- Next search result (optimized) ---
  if (lowerText === '/next_search_result') {
    const search = userToSearchResults.get(userKey);
    if (!search || !search.matches.length) {
      await context.send('No active search. Use `/search_insights <keyword>` first.');
      return;
    }
    search.idx++;
    if (search.idx >= search.matches.length) {
      await context.send('No more search results. Use `/search_insights <keyword>` to search again.');
      userToSearchResults.delete(userKey);
      return;
    }
    const card = createFeedbackAnalysisCard(search.matches[search.idx].analysis);
    await context.send({
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: card
      }]
    });
    // --- Optimized: Send concise summary and priority for quick context ---
    const a = search.matches[search.idx].analysis;
    if (!('error' in a)) {
      let msg = `**Summary:** ${a.summary}\n**Priority:** ${a.priority.toUpperCase()}\n**Pain Points:**\n`;
      if (a.painPoints && a.painPoints.length > 0) {
        msg += a.painPoints.map((p: string) => `- ${p}`).join('\n');
      } else {
        msg += 'None identified.';
      }
      await context.send(msg);
    }
    if (search.idx < search.matches.length - 1) {
      await context.send(`Type /next_search_result to see the next match.`);
    } else {
      await context.send('That was the last search result.');
    }
    // Track this as the current card for follow-up
    userToInsightIndex.set(userKey, latestAnalysis.indexOf(search.matches[search.idx]));
    return;
  }

  // --- Ask about the current card ---
  if (lowerText.startsWith('/ask_about_current')) {
    const idx = userToInsightIndex.get(userKey);
    if (idx === undefined || !latestAnalysis[idx]) {
      await context.send('No current insight selected. Use `/show_insights`, `/latest_insight`, or `/search_insights <keyword>` first.');
      return;
    }
    const question = text?.slice('/ask_about_current'.length).trim();
    if (!question) {
      await context.send('Please provide a question. Example: `/ask_about_current What is the main pain point?`');
      return;
    }
    const analysis = latestAnalysis[idx].analysis;
    // Use AI to answer based on the card's content
    try {
      const prompt = new ChatPrompt({
        instructions: `You are an assistant. Given the following feedback analysis, answer the user's question as helpfully as possible.\n\nFeedback Analysis:\n${JSON.stringify(analysis, null, 2)}`,
        model: openaiModel,
      });
      const aiResponse = await prompt.send(question);
      await context.send(aiResponse.content?.trim() || "Sorry, I couldn't answer your question.");
    } catch (err: any) {
      await context.send("Sorry, I couldn't process your question.");
    }
    return;
  }

  // --- Show insights one by one (paging) ---
  if (lowerText === '/show_insights') {
    if (!latestAnalysis.length) {
      await context.send('No analysis results are available yet.');
      return;
    }
    userToInsightIndex.set(userKey, 0);
    const result = latestAnalysis[0];
    const card = createFeedbackAnalysisCard(result.analysis);
    await context.send({
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: card
      }]
    });
    if (latestAnalysis.length > 1) {
      await context.send(`Type /next_insight to see the next insight (${latestAnalysis.length - 1} more).`);
    }
    return;
  }

  if (lowerText === '/next_insight') {
    if (!latestAnalysis.length) {
      await context.send('No analysis results are available yet.');
      return;
    }
    let idx = userToInsightIndex.get(userKey) ?? 0;
    idx++;
    if (idx >= latestAnalysis.length) {
      await context.send('No more insights. Type /show_insights to start over.');
      userToInsightIndex.set(userKey, 0);
      return;
    }
    userToInsightIndex.set(userKey, idx);
    const result = latestAnalysis[idx];
    const card = createFeedbackAnalysisCard(result.analysis);
    await context.send({
      type: 'message',
      attachments: [{
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: card
      }]
    });
    if (idx < latestAnalysis.length - 1) {
      await context.send(`Type /next_insight to see the next insight (${latestAnalysis.length - idx - 1} more).`);
    } else {
      await context.send('That was the last insight. Type /show_insights to start over.');
    }
    return;
  }

  // --- Default: echo or help ---
  await context.send(`You said: "${context.activity.text}".\n\nCommands:\n- /show_insights\n- /next_insight\n- /latest_insight\n- /search_insights <keyword>\n- /next_search_result\n- /ask_about_current <your question>`);
  // Store conversation ID for proactive messages later (if needed)
  if (context.activity.from.aadObjectId && context.activity.conversation.id && !userToConversationId.has(context.activity.from.aadObjectId)) {
    userToConversationId.set(context.activity.from.aadObjectId, context.activity.conversation.id);
    teamsApp.log.info( // Keep using teamsApp.log as it's directly in scope
      `Just added user ${context.activity.from.aadObjectId} to conversation ${context.activity.conversation.id}`
    );
  }
});

// --- Start both applications ---
(async () => {
    try {
        // Start the Teams AI App (Bot)
        await teamsApp.start(TEAMS_BOT_PORT);
        console.log(`Teams AI App (Bot & Devtools) listening on http://localhost:${TEAMS_BOT_PORT}`);
        console.log(`Teams Bot Endpoint: http://localhost:${TEAMS_BOT_PORT}/api/messages`);
        console.log(`Devtools available at http://localhost:${DEVTOOLS_PORT}/devtools`);

        // Start the dedicated MCP Ingestion Server
        mcpExpressApp.listen(MCP_SERVER_PORT, () => {
            console.log(`MCP Ingestion Server listening on http://localhost:${MCP_SERVER_PORT}`);
            console.log(`MCP Ingestion Endpoint: http://localhost:${MCP_SERVER_PORT}/api/mcp/ingest`);
        });

        console.log('Ensure your .env file has AZURE_OPENAI_API_KEY, ENDPOINT, DEPLOYMENT_NAME, and API_VERSION.');

    } catch (error) {
        console.error('Failed to start one or both applications:', error);
        process.exit(1); // Exit if app fails to start
    }
})();