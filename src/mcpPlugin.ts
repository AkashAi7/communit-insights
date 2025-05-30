// import { Tool, TurnContext } from "@microsoft/teams-ai";
// import fetch from "node-fetch";

// // Tool to analyze feedback using your local ingestion endpoint
// export const analyzeFeedbackTool: Tool = {
//   name: "analyze-feedback",
//   description: "Analyzes latest developer feedback from GitHub and Stack Overflow",
//   inputSchema: {
//     type: "object",
//     properties: {
//       input: { type: "string" },
//     },
//     required: ["input"],
//   },
//   async function(context: TurnContext, state: any, input: { input: string }) {
//     try {
//       const res = await fetch("http://localhost:3975/api/mcp/ingest", {
//         method: "POST",
//         headers: { "Content-Type": "application/json" },
//         body: JSON.stringify({ feedback: [{ text: input.input }] }),
//       });

//       if (!res.ok) {
//         const text = await res.text();
//         throw new Error(`Ingestion endpoint error ${res.status}: ${text}`);
//       }

//       const result = await res.json();
//       return {
//         output: JSON.stringify(result.analyzedResults?.[0] || "No result"),
//       };
//     } catch (err: any) {
//       return { output: `Error: ${err.message}` };
//     }
//   },
// };
