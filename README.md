# 🚀 All-in-One Google MCP (Google Apps Script)

A unified, high-performance **Model Context Protocol (MCP)** backend powered by **Google Apps Script (GAS)**. This tool allows LLMs (like Antigravity, Gemini, ChatGPT, and Claude) to interact directly with your Google Workspace services with zero local overhead.

---

## ✨ Features

- **📧 Gmail**: List threads, search emails, get details, send replies, archive, and delete.
- **📁 Drive**: List files, search, create/update/delete files and folders, and share.
- **📅 Calendar**: Manage events, search by date range, create/update/delete.
- **✅ Tasks**: Unified task management (List, Create, Update, Delete).
- **📊 Sheets**: Read, write, append, and clear spreadsheet data.
- **📝 Docs**: Create, read, and append to Google Documents.

---

## 🛠️ Installation & Setup

### Option 1: Google Apps Script Web IDE (Fastest)

1.  **Create a New Script**: Go to [script.google.com](https://script.google.com) and create a new project.
2.  **Add Code**: Copy the content of `Code.gs` from this repository and paste it into the editor.
3.  **Configure Scopes**: Click on the Gear icon (Project Settings) and check **"Show 'appsscript.json' manifest file in editor"**.
4.  **Update Manifest**: Copy the content of `appsscript.json` from this repository into the script's manifest file.
5.  **Enable Advanced Services**:
    - Click the **+** next to "Services".
    - Find **Tasks API** and click **Add**.
6.  **Set API Key**:
    - In Project Settings, scroll to **Script Properties**.
    - Add a new property: `MCP_API_KEY` with a secure random string as the value.
7.  **Deploy**:
    - Click **Deploy > New Deployment**.
    - Select **Web App**.
    - **Execute as**: Me.
    - **Who has access**: Anyone (The `MCP_API_KEY` protects your script).
    - Copy the **Web App URL**.

### Option 2: Local Development (VS Code + Clasp)

1.  **Initialize**:
    ```bash
    npm install -g @google/clasp
    clasp login
    clasp clone <YOUR_SCRIPT_ID>
    ```
2.  **Push Changes**:
    ```bash
    clasp push
    ```
3.  **Deploy**: Use the Web IDE or `clasp deploy` to manage your deployments.

---

## ⚙️ Configuration

### Integration with LLM Clients

To use this with an MCP client (like Antigravity or a custom proxy), configure it to talk to your Web App URL.

**Base URL**: `https://script.google.com/macros/s/.../exec`

#### Query Parameters:
- `?tools=1`: Discovery (List all available tools).
- `?schema=1`: Fetch OpenAI-compatible tool definitions.

#### POST Request:
The server expects a standard JSON-RPC payload:
```json
{
  "method": "tools/call",
  "params": {
    "name": "gmail_listThreads",
    "arguments": { "limit": 5 }
  },
  "apiKey": "YOUR_MCP_API_KEY"
}
```

---

## 🛡️ Security

This script uses a simple API Key authentication mechanism. **Never share your Web App URL or API Key publicly.** If you believe your key is compromised, update the `MCP_API_KEY` in Script Properties and re-deploy if necessary.

## 📄 License

MIT © [SH20RAJ](https://github.com/SH20RAJ)
