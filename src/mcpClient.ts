import fetch from 'node-fetch';
import * as dotenv from 'dotenv';
dotenv.config();

const STACK_OVERFLOW_API_KEY = process.env.STACK_OVERFLOW_API_KEY!;
const GITHUB_TOKEN = process.env.GITHUB_TOKEN!;
// CHANGE THIS: The URL for your custom Express ingestion endpoint
const INGEST_ENDPOINT_URL = 'http://localhost:3975/api/mcp/ingest';

// Helper function to send data to your custom ingestion endpoint
async function sendFeedbackToIngestEndpoint(feedbackData: any[]) {
  const res = await fetch(INGEST_ENDPOINT_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    // The server expects a JSON body with a 'feedback' array
    body: JSON.stringify({ feedback: feedbackData }),
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Ingestion endpoint error ${res.status}: ${text}`);
  }

  return res.json();
}

async function fetchStackOverflowQuestions() {
  const res = await fetch(
    `https://api.stackexchange.com/2.3/questions?order=desc&sort=activity&tagged=microsoft-teams&site=stackoverflow&filter=withbody&key=${STACK_OVERFLOW_API_KEY}`
  );
  const data = await res.json();
  if (!Array.isArray(data.items)) {
    throw new Error('Stack Overflow response invalid');
  }
  return data.items.map((q: any) => ({
    id: q.question_id,
    text: `${q.title} ${q.body}`,
    source: 'Stack Overflow',
    url: `https://stackoverflow.com/questions/${q.question_id}`
  }));
}

async function fetchGitHubIssues() {
  const res = await fetch(`https://api.github.com/repos/MicrosoftDocs/msteams-docs/issues`, {
    headers: { Authorization: `token ${GITHUB_TOKEN}` },
  });
  const data = await res.json();
  if (!Array.isArray(data)) {
    throw new Error('GitHub response invalid');
  }
  return data.map((issue: any) => ({
    id: issue.id,
    text: `${issue.title} ${issue.body}`,
    source: 'GitHub Issues',
    url: issue.html_url
  }));
}

async function runMcpClient() {
  try {
    const stackOverflow = await fetchStackOverflowQuestions();
    const gitHub = await fetchGitHubIssues();
    const allFeedback = [...stackOverflow, ...gitHub];

    console.log(`📤 Fetching ${allFeedback.length} feedback items from sources...`);
    console.log(`🚀 Sending feedback to local ingestion endpoint...`);

    // CHANGE THIS: Call the new function to send to the ingestion endpoint
    const response = await sendFeedbackToIngestEndpoint(allFeedback);

    console.log('✅ Ingestion endpoint response:', JSON.stringify(response, null, 2));
    if (response.analyzedResults && response.analyzedResults.length > 0) {
      console.log('Analyzed Insights for first 3 items:', JSON.stringify(response.analyzedResults.slice(0, 3), null, 2));
      if (response.analyzedResults.length > 3) {
        console.log(`... and ${response.analyzedResults.length - 3} more items.`);
      }
    }

  } catch (err) {
    console.error('❌ MCP client error:', err);
  }
}

runMcpClient();