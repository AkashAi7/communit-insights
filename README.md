 # Community Insider Bot for Teams

## Project Overview

The Community Insider Bot is a Microsoft Teams application designed to help developers and product teams stay informed about community feedback and pain points from various developer channels. This bot integrates with the Teams AI Library v2 to facilitate:

- **Feedback Ingestion via MCP Protocol:** Configures an MCP client to pull developer feedback from sources like Stack Overflow and GitHub Issues and sends it to a dedicated ingestion service.
- **AI-Driven Feedback/Pain Point Extraction:** Processes ingested feedback using an Azure OpenAI model to identify key pain points, summarize insights, and assign priority levels.
- **Conversational Teams Bot:** Provides an interactive interface within Microsoft Teams, allowing users to browse, search, and inquire about the extracted insights using Adaptive Cards for a rich user experience.

This solution demonstrates the power of the Teams AI Library v2 for building intelligent, data-driven Teams applications.

## Features

- **Multi-Source Feedback Ingestion:** Connects to Stack Overflow (tagged `microsoft-teams` questions) and GitHub (`MicrosoftTeams/msteams-docs` issues) to collect developer feedback.
- **MCP Protocol Integration:** Implements a server-side MCP (Microsoft Collaboration Protocol) service to receive and process feedback data from the ingestion client.
- **AI-Powered Analysis:** Utilizes Azure OpenAI to:
    - Extract key pain points from feedback.
    - Generate concise summaries of issues.
    - Assign a priority (low, medium, high) to each feedback item.
- **Conversational Bot Interface:**
    - **Browse Insights:** Step through analyzed feedback items one by one (`/show_insights`, `/next_insight`, `/latest_insight`).
    - **Search Insights:** Find specific insights by keyword (`/search_insights <keyword>`, `/next_search_result`).
    - **Ask Follow-up Questions:** Interact with the AI to get more details about the currently displayed insight (`/ask_about_current <your question>`).
- **Adaptive Cards:** Presents feedback analysis in a clear, actionable format within Teams chats.
- **Scalable Architecture:** Separates the MCP ingestion server from the main Teams bot application for better maintainability and potential scaling.

## Deliverables

- **Source Code:** All necessary code for the client, ingestion service, and Teams bot.
- **Readme with Setup Instructions:** This document, providing all information needed to set up and run the project.

## Project Structure

```
.
├── .vscode/                  # VS Code configuration and debug instructions
├── appPackage/               # Teams app package related files (app manifest)
├── devTools/                 # Development tools configuration
├── env/                      # Environment-specific configuration files (managed by Teams Toolkit)
│   ├── .env.local            # Local environment variables (managed by Toolkit)
│   ├── .env.local.user       # User-specific local environment variables (managed by Toolkit)
│   └── .env.testtool         # Test tool environment variables (managed by Toolkit)
├── infra/                    # Infrastructure-as-code (if any, e.g., for deployment)
├── node_modules/             # Installed Node.js packages
├── src/
│   ├── client.ts             # MCP client to fetch data from Stack Overflow and GitHub
│   └── index.ts              # Teams Bot application and MCP ingestion service
├── .env                      # **Your primary environment variables for API keys and global settings**
├── package-lock.json         # Records exact dependency versions
├── package.json              # Project dependencies and scripts
├── teamsapp.local.yml        # Teams Toolkit local debug configuration
├── teamsapp.testtool.yml     # Teams Toolkit test tool configuration
├── teamsapp.yml              # Teams Toolkit app manifest configuration
├── tsconfig.json             # TypeScript compiler configuration
└── tsup.config.js            # tsup configuration for building (if used)
```

## Setup Instructions

Follow these steps to get your Community Insider Bot up and running.

### Prerequisites

- **Node.js:** (v18 or higher recommended)
- **pnpm:** (or npm/yarn) - `npm install -g pnpm`
- **Azure Subscription:** Required to access Azure OpenAI Service. You can create a free Azure account.
- **Stack Apps Account:** To obtain a Stack Overflow API Key (free).
- **GitHub Personal Access Token (PAT):** For accessing GitHub Issues. Generate one with `repo` scope if accessing private repos or `public_repo` for public ones.
- **Microsoft Teams Toolkit for VS Code:** Essential for local development and deployment of Teams apps. Install it from the VS Code Marketplace.

### 1. Environment Configuration

1.  **Clone the repository:**
    ```bash
    git clone <your-repository-url>
    cd community-insider-bot
    ```
2.  **Install dependencies:**
    ```bash
    pnpm install
    ```
3.  **Create and Populate `.env` file in the root directory:**
    Create a new file named `.env` directly in the root of your project (i.e., in the `quote-agent` directory). This file is where you will store your sensitive API keys.
    ```bash
    touch .env
    ```
    Open this `.env` file and add the following variables:
    ```dotenv
    # Azure OpenAI Configuration
    AZURE_OPENAI_API_KEY=<YOUR_AZURE_OPENAI_API_KEY>
    AZURE_OPENAI_ENDPOINT=<YOUR_AZURE_OPENAI_ENDPOINT> # e.g., https://YOUR_RESOURCE_NAME.openai.azure.com/
    AZURE_OPENAI_API_VERSION=2024-02-15-preview # Or your specific API version
    AZURE_OPENAI_DEPLOYMENT_NAME=<YOUR_AZURE_OPENAI_DEPLOYMENT_NAME> # The name of your deployed model (e.g., gpt-4, gpt-35-turbo)

    # API Keys for Feedback Ingestion
    STACK_OVERFLOW_API_KEY=<YOUR_STACK_OVERFLOW_API_KEY>
    GITHUB_TOKEN=<YOUR_GITHUB_PERSONAL_ACCESS_TOKEN>

    # Bot and Server Ports (Defaults are fine, adjust if conflicts)
    PORT=3976 # Teams Bot Port
    MCP_SERVER_PORT=3975 # Dedicated MCP Ingestion Server Port
    DEVTOOLS_PORT=3977 # Devtools Port (internal)
    ```
    - **Azure OpenAI:** Ensure you have a deployed model in Azure OpenAI Studio and note down its key, endpoint, API version, and deployment name.
    - **Stack Overflow API Key:** Register an application on the Stack Apps site to get a key.
    - **GitHub Token:** Generate a Personal Access Token in your GitHub settings (Settings > Developer settings > Personal access tokens).

### 2. Configure with Microsoft 365 Agents Toolkit

To configure your agent for Teams, you'll use the Microsoft 365 Agents Toolkit CLI.

1.  **Install Microsoft 365 Agents Toolkit IDE extension:**
    Visit the Microsoft 365 Agents Toolkit installation guide to install it on your preferred IDE (e.g., VS Code Marketplace).
2.  **Add Teams configuration files via `teams` CLI:**
    Open your terminal inside your `quote-agent` project folder and run the following command:
    ```bash
    npx @microsoft/teams.cli config add ttk.basic
    ```
    **Tip:** If you have `teams` CLI installed globally, use `teams` instead of `npx`.

    **Tip:** The `ttk.basic` configuration provides a basic setup for Microsoft 365 Agents Toolkit. It includes the necessary files and configuration to get started with Teams development. Explore more advanced configurations as needed with `teams config --help`.

    This CLI command adds configuration files required by Microsoft 365 Agents Toolkit, including:
    - Environment setup in the `env` folder and populates the root `.env` file with necessary Teams Toolkit variables.
    - Teams app manifest in the `appPackage` folder (if not already present).
    - Debug instructions in `.vscode/launch.json` and `.vscode/tasks.json`.
    - Agents Toolkit automation files to your project (e.g., `teamsapp.local.yml`).

### 3.  INex.ts ( server and teams bot logic) 

First, start the main application which includes the MCP Ingestion Server. This needs to be running before you try to ingest data.

```bash
npm run dev
```

You should see output similar to this, indicating both the Teams App server (which will be exposed by Toolkit) and MCP Ingestion Server are listening:

```
Teams AI App (Bot & Devtools) listening on http://localhost:3976
Teams Bot Endpoint: http://localhost:3976/api/messages
Devtools available at http://localhost:3977/devtools
MCP Ingestion Server listening on http://localhost:3975
MCP Ingestion Endpoint: http://localhost:3975/api/mcp/ingest
```

Ensure your `.env` file has `AZURE_OPENAI_API_KEY`, `ENDPOINT`, `DEPLOYMENT_NAME`, and `API_VERSION`.

### 4. Run the MCP Ingestion Client

In a separate terminal window, run the MCP Client to fetch and send feedback to your locally running ingestion server:

```bash
npx ts-node -r dotenv/config ./src/mcpClient.ts
```

You'll see messages indicating the client is fetching data and sending it to your local ingestion endpoint. Once complete, the `index.ts` server will process and analyze the feedback using your Azure OpenAI model.

```
📤 Fetching X feedback items from sources...
🚀 Sending feedback to local ingestion endpoint...
✅ Ingestion endpoint response: ... (JSON output of analyzed results)
Analyzed Insights for first 3 items: ...
```

### 5. Debugging in Teams (Using Microsoft 365 Agents Toolkit)

Now that your agent is running locally, let's deploy it to Microsoft Teams for testing.

1.  Open your agent's project in VS Code.
2.  Open the Microsoft 365 Agents Toolkit extension panel (usually on the left sidebar, the Teams logo icon).
3.  Log in to your Microsoft 365 and Azure accounts in the Agents Toolkit extension.
4.  Select "Local" under Environment Settings of the Agents Toolkit extension.
5.  Click on **Debug (Chrome)** or **Debug (Edge)** to start debugging via the 'play' button within the Teams Toolkit panel.

When debugging starts, the Microsoft 365 Agents Toolkit will:

- Build your application.
- Start a devtunnel which will assign a temporary public URL to your local server.
- Provision the Teams app for your tenant so that it can be installed and be authenticated on Teams.
- Set up the local variables necessary for your agent to run in Teams in `env/.env.local` and `env/env.local.user`. This includes propagating the app manifest with your newly provisioned resources.
- Start the local server.
- Package your app manifest into a Teams application zip package and the manifest json with variables inserted in `appPackage/build`.
- Launch Teams in an incognito window in your browser.
- Upload the package to Teams and signal it to sideload the app (installing this app just for your user).

## Usage

Once your bot is installed in Teams and the MCP client has ingested data, you can interact with it in a chat:

- `hi` / `hello`: Get a friendly welcome message and an overview of capabilities.
- `help` / `what can you do` / `/commands`: See a list of available commands.
- `/show_insights`: Start Browse the analyzed feedback insights one by one.
- `/next_insight`: View the next analyzed insight.
- `/latest_insight`: Display the most recently ingested and analyzed insight.
- `/search_insights <keyword>`: Search for insights containing a specific keyword in their summary or pain points.
    - *Example:* `/search_insights authentication`
- `/next_search_result`: If you've performed a search, see the next result.
- `/ask_about_current <your question>`: Ask a follow-up question about the currently displayed insight. The AI will try to answer based on the analysis.
    - *Example:* `/ask_about_current What are the implications of this pain point?`

## Technologies Used

- **Microsoft Teams AI Library v2:** Core framework for building intelligent Teams bots.
- **TypeScript:** Primary programming language.
- **Node.js:** JavaScript runtime.
- **Express.js:** Web framework used for the MCP ingestion server.
- **Azure OpenAI Service:** For AI model capabilities (analysis, summarization, pain point extraction).
- **Zod:** For schema validation of MCP tool inputs and outputs.
- **Adaptive Cards:** For rich UI in Teams messages.
- **node-fetch:** For making HTTP requests to external APIs.
- **dotenv:** For managing environment variables.
- **Microsoft Teams Toolkit:** For streamlined local development, debugging, and deployment to Teams.

## AI Model and Prompt Engineering

The `index.ts` file contains the `analyzeFeedbackToolHandler` which uses `ChatPrompt` with specific instructions to guide the Azure OpenAI model.

- **Prompt Template:** The prompt instructs the AI to:
    - Analyze developer feedback.
    - Identify key pain points, recurring issues, and actionable insights.
    - Output a JSON object with `painPoints` (array of strings), `summary` (string), and `priority` (low, medium, high).
- **Fine-tuning Strategies (Conceptual):** While explicit fine-tuning a model is a separate process, the current implementation maximizes accuracy through:
    - **Clear Instructions:** Providing precise instructions in the `instructions` property of `ChatPrompt`.
    - **Structured Output:** Demanding a JSON output schema helps the model produce parsable results.
    - **Robust Error Handling:** The code includes error handling for AI output parsing, making the system more resilient to unexpected model responses.
- **For further maximization of extraction accuracy, consider:**
    - **Few-shot Learning:** Include examples of good feedback analysis (input and desired output) within the prompt for better guidance.
    - **Custom Models (if needed):** For highly specific domains, fine-tuning your own model with a custom dataset could yield even better results.

