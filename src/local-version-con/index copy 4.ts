// src/index.ts

import * as dotenv from 'dotenv';
dotenv.config(); // Load environment variables from .env file

import express from 'express';
import bodyParser from 'body-parser'; // For parsing request bodies




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

// --- NEW: Store latest feedback and analysis results ---
let latestFeedback: any[] = [];
let latestAnalysis: any[] = []; // <-- cache analysis results

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

mcpExpressApp.post('/api/mcp/ingest', async (req, res) => {
    try {
        console.log(`Received POST to /api/mcp/ingest on port ${MCP_SERVER_PORT} at ${new Date().toISOString()}`);
        const validationResult = AnalyzeFeedbackInputSchema.safeParse(req.body);

        if (!validationResult.success) {
            console.error('MCP Input Validation Error:', validationResult.error);
            return res.status(400).json({ error: 'Invalid input schema', details: validationResult.error.issues });
        }

        const input = validationResult.data;
        // --- Store latest feedback for /show_insights ---
        latestFeedback = input.feedback;

        const output = await analyzeFeedbackToolHandler({ feedback: input.feedback });

        // --- Cache the analysis results for Teams bot ---
        latestAnalysis = output.analyzedResults || [];

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
    // --- Handle error analysis objects gracefully ---
    if ('error' in analysis) {
        return {
            type: "AdaptiveCard",
            $schema: "http://adaptivecards.io/schemas/adaptiveCard.json",
            version: "1.3",
            body: [
                {
                    type: "TextBlock",
                    text: "Feedback Analysis Error",
                    wrap: true,
                    size: "Large",
                    weight: "Bolder",
                    color: "Attention"
                },
                {
                    type: "TextBlock",
                    text: `Error: ${analysis.error}`,
                    wrap: true,
                    color: "Attention"
                },
                ...(analysis.rawOutput ? [{
                    type: "TextBlock",
                    text: `Raw Output: ${typeof analysis.rawOutput === 'string' ? analysis.rawOutput : JSON.stringify(analysis.rawOutput)}`,
                    wrap: true,
                    isSubtle: true
                }] : [])
            ]
        };
    }

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

  const text = context.activity.text?.trim().toLowerCase();

  if (text === 'hello') {
      await context.send('Hello! I am your Community Insider Bot. Type `/show_insights` to see recent feedback analysis.');
  } else if (text === '/show_insights') {
      // --- Check if there is any feedback to analyze ---
      if (!latestFeedback || latestFeedback.length === 0) {
          await context.send('No developer feedback has been ingested yet. Please submit feedback via the ingestion endpoint before requesting insights.');
          return;
      }

      if (!latestAnalysis || latestAnalysis.length === 0) {
          await context.send('No analysis results are available yet. Please try again after feedback is processed.');
          return;
      }

      await context.send('Getting the latest feedback analysis for you...');
      // --- Use cached analysis results ---
      let sentAny = false;
      for (const result of latestAnalysis) {
          const analysis = result.analysis;
          const card = createFeedbackAnalysisCard(analysis);

          if (!card) {
              await context.send("Could not render analysis card for a feedback item.");
              continue;
          }

          await context.send({
              type: 'message',
              attachments: [{
                  contentType: 'application/vnd.microsoft.card.adaptive',
                  content: card
              }]
          });
          sentAny = true;
      }
      if (!sentAny) {
          await context.send('No valid analysis results to display.');
      }
      console.log('Sent Adaptive Card(s) with cached analysis results to Teams.');
  } else {
      // --- NEW: Use AI model for general conversation ---
      try {
        const prompt = new ChatPrompt({
          instructions: `You are a helpful, friendly assistant for developer communities. Respond conversationally to the user's message.`,
          model: openaiModel,
        });
        const aiResponse = await prompt.send(context.activity.text || "");
        if (aiResponse.content) {
          await context.send(aiResponse.content.trim());
        } else {
          await context.send("I'm here to help! Please ask me anything about developer feedback or insights.");
        }
      } catch (err: any) {
        console.error('Error generating conversational AI response:', err);
        await context.send("Sorry, I couldn't process your message right now.");
      }
  }

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