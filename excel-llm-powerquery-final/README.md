
# Excel LLM PowerQuery Generator (MCP Server)

## 🔧 Setup

```bash
pip install -r requirements.txt
```

## 🔑 Configure

Update the `.env` file:

```env
OPENAI_API_KEY=your-key
EXCEL_FILE_PATH=examples/sample_excel_powerquery.xlsx
DEFAULT_PROMPT=Create a column that returns "High" if [Sales] > 1000, otherwise "Low".
```

## 🏃 Run It

```bash
python mcp_excel_server.py
```

## 🛠 Local Development with FastMCP CLI

Install FastMCP CLI:

```bash
npm install -g @flexdotai/fastmcp-cli
```

Run in local dev mode:

```bash
mcp dev --config mcp.json
```

Then type prompts like:

```
update
Create a column that returns "High" if [Sales] > 1000, otherwise "Low"
```

The tool will update the Excel Power Query using GPT-generated M code.


## 📄 License

This project is licensed under the MIT License. See `LICENSE` for details.