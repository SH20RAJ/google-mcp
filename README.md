# 🚀 All-in-One Google MCP

[![Stars](https://img.shields.io/github/stars/SH20RAJ/google-mcp?style=for-the-badge)](https://github.com/SH20RAJ/google-mcp/stargazers)
[![Forks](https://img.shields.io/github/forks/SH20RAJ/google-mcp?style=for-the-badge)](https://github.com/SH20RAJ/google-mcp/network/members)
[![Issues](https://img.shields.io/github/issues/SH20RAJ/google-mcp?style=for-the-badge)](https://github.com/SH20RAJ/google-mcp/issues)
[![MIT License](https://img.shields.io/github/license/SH20RAJ/google-mcp?style=for-the-badge)](https://github.com/SH20RAJ/google-mcp/blob/main/LICENSE)
![Visitors](https://api.visitorbadge.io/api/visitors?path=SH20RAJ%2Fgoogle-mcp&label=VISITORS&countColor=%23263238&style=for-the-badge)

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

## 🛠️ Connection & Setup

### 🔗 Using `mcp-remote` (Recommended)

The easiest way to connect your LLM client to this backend is using `mcp-remote`. Add the following configuration to your MCP settings file (e.g., `mcp_config.json` for Antigravity or Desktop App):

```json
{
  "mcpServers": {
    "google-mcp": {
      "command": "npx",
      "args": [
        "-y",
        "mcp-remote@latest",
        "https://script.google.com/macros/s/AKfycbzbiXiGCEg1bjWkp58A6DSfU78I-n2bNohZW80tGFk3-hZjbsZh6Vj8En8SNMJd8Exa/exec?key=MCP_API_KEY"
      ],
      "disabledTools": []
    }
  }
}
```

> [!TIP]
> Replace the URL with your own deployment URL if you are hosting your own instance.

---

## 💻 Platform Setup

### 🌌 Antigravity / Gemini Desktop
1. Open your `mcp_config.json`.
2. Add the `google-mcp` configuration block shown above.
3. Restart the application or refresh the MCP server list.

### 💻 VS Code (MCP Client)
If you are using an MCP-compatible extension in VS Code:
1. Locate the extension's configuration for MCP servers.
2. Add the `npx mcp-remote` command and your Web App URL.

### 🌐 Other Clients & Platforms
Since this is a standard MCP server reached via `mcp-remote`, it works with any client that supports the Model Context Protocol, including:
- **Claude Desktop**
- **Custom MCP Proxies**
- **Self-hosted LLM interfaces**

---

## 📝 Usage & Example Prompts

Once connected, you can ask your AI to perform tasks across Google Workspace. Here are some examples:

- **Gmail**: "Read my latest email and summarize it."
- **Calendar**: "Do I have any meetings tomorrow?" or "Schedule a 'Focus Time' from 2 PM to 4 PM today."
- **Sheets**: "Read the first 10 rows of the 'Budget' spreadsheet."
- **Docs**: "Create a new document called 'Project Notes' and add a summary of our discussion."
- **Tasks**: "List my pending tasks for today."

---

## 📊 Quotas & Limits

Google Apps Script has daily quotas that are generous for individual use but worth noting:

| Service | Consumer (@gmail.com) | Workspace |
| :--- | :--- | :--- |
| **URL Fetch** | 20,000 / day | 100,000 / day |
| **Email Send** | 100 recipients / day | 1,500 recipients / day |
| **Triggers Runtime** | 90 min / day | 6 hours / day |
| **Simultaneous Executions** | 30 | 30 |

> [!NOTE]
> Each script execution has a hard limit of **6 minutes**. Large batch operations should be broken down into smaller chunks.

---

## ⚙️ Backend Installation (Self-Hosting)

1.  **Create a New Script**: Go to [script.google.com](https://script.google.com) and create a new project.
2.  **Add Code**: Copy the content of `Code.gs` from this repository and paste it into the editor.
3.  **Update Manifest**: Switch to the `appsscript.json` tab (enable it in Settings if hidden) and paste the manifest content.
4.  **Set API Key**: In **Project Settings**, add a Script Property `MCP_API_KEY` with your secret key.
5.  **Deploy**: Click **Deploy > New Deployment > Web App**. Set Access to "Anyone".

---

## 🛡️ Security

This project uses a simple API Key authentication mechanism passed via the `key` query parameter or `apiKey` JSON field. **Keep your Deployment URL secret.**

## 🤝 Contributing

Contributions are welcome! If you'd like to improve this project, please follow these steps:

1.  **Fork** the repository.
2.  Create a new **Feature Branch** (`git checkout -b feature/AmazingFeature`).
3.  **Commit** your changes (`git commit -m 'Add some AmazingFeature'`).
4.  **Push** to the branch (`git push origin feature/AmazingFeature`).
5.  Open a **Pull Request**.

Please ensure your code follows the existing style and includes appropriate documentation.

## 📄 License

MIT © [SH20RAJ](https://github.com/SH20RAJ)
