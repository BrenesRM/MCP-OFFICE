# 📂 Office MCP Server

An **MCP (Model Context Protocol) server** for handling **Microsoft Office files**:

- **Word (.docx)** → create, read, update, delete  
- **Excel (.xlsx)** → create, read, update, delete  
- **PowerPoint (.pptx)** → create, read, update, delete  

Built with **Python**, **uv**, and MCP’s `fastmcp`.

---

## ⚡ Host Setup

First, make sure you have **uv** installed:  
👉 [uv installation guide](https://github.com/astral-sh/uv)

---

## 📦 Installation

```bash
# Initialize project
uv init
uv venv

# Activate virtual environment
.venv\Scripts\activate   # on Windows
# source .venv/bin/activate   # on Linux/Mac

# Install MCP and required Office libraries
uv add mcp[cli]
uv add openpyxl python-pptx python-docx

mcp dev server.py   # Run in development mode
mcp install server.py   # Install into your MCP client

🗂 Documents Folder

All Office files are created in the local ./documents folder by default.

🛠 Available Tools

Word → create_docx, read_docx, update_docx, delete_docx

Excel → create_xlsx, read_xlsx, update_xlsx, delete_xlsx

PowerPoint → create_pptx, read_pptx, update_pptx, delete_pptx
