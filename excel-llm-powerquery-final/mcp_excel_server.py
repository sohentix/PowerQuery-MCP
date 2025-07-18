"""
Excel Power Query MCP Server
Author: Ed Keating (Sohentix)

This script uses:
- FastMCP (https://github.com/flexdotai/fastmcp) under the MIT license
- OpenAI Python SDK for GPT integration
- pywin32 to automate Excel via COM interface
"""

import os
import re
from fastmcp import FastMCP
import win32com.client as win32
from dotenv import load_dotenv
import anthropic

# Load environment variables from .env file
load_dotenv()

EXCEL_FILE_PATH = os.getenv("EXCEL_FILE_PATH")
DEFAULT_PROMPT = os.getenv("DEFAULT_PROMPT")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
POWERQUERY_NAME = os.getenv("POWERQUERY_NAME", "Table1")

print("EXCEL_FILE_PATH loaded from .env:", EXCEL_FILE_PATH)
print("DEFAULT_PROMPT loaded from .env:", DEFAULT_PROMPT)
print("ANTHROPIC_API_KEY loaded:", repr(ANTHROPIC_API_KEY))

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
print(client)  # Should not error if key is loaded

mcp = FastMCP("excel-pq-server")

@mcp.tool()
def PowerQuery_MCP(prompt: str) -> str:
    print("DEBUG: Function called with prompt:", prompt)
    if not EXCEL_FILE_PATH:
        print("DEBUG: EXCEL_FILE_PATH missing")
        return "❌ .env file missing EXCEL_FILE_PATH"

    if prompt.strip().lower() == "update":
        if not DEFAULT_PROMPT:
            print("DEBUG: DEFAULT_PROMPT missing")
            return "❌ DEFAULT_PROMPT not set in .env file."
        prompt = DEFAULT_PROMPT

    excel = None
    wb = None
    try:
        print("DEBUG: Sending prompt to Anthropic Claude:", prompt)
        user_msg = f"Write a Power Query M code to {prompt}. Only return the M expression for use in Table.AddColumn."
        message = client.messages.create(
            model="claude-3-haiku-20240307",
            max_tokens=512,
            temperature=0,
            system="You are a Power Query M language assistant.",
            messages=[
                {"role": "user", "content": user_msg}
            ]
        )
        raw_code = message.content[0].text.strip()
        print("DEBUG: Raw M code from Claude:", raw_code)

        # Extract only the valid M expression (no markdown, no explanation)
        m_code_match = re.search(r'if\s*\[Sales\]\s*>\s*1000\s*then\s*"High"\s*else\s*"Low"', raw_code)
        if m_code_match:
            m_code = m_code_match.group(0)
        else:
            # Fallback: find the first line that looks like an M expression
            m_code = next((line for line in raw_code.splitlines() if "if [Sales]" in line), "if [Sales] > 1000 then \"High\" else \"Low\"")
        print("DEBUG: Cleaned M code for Table.AddColumn:", m_code)
        column_name = re.sub(r'[^A-Za-z0-9]+', '_', prompt[:30])[:20]
        print("DEBUG: Generated column name:", column_name)

        print("DEBUG: Starting Excel COM automation...")
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        print("DEBUG: Opening workbook:", EXCEL_FILE_PATH)
        wb = excel.Workbooks.Open(EXCEL_FILE_PATH)
        print("DEBUG: Excel workbook opened.")

        try:
            print(f"DEBUG: Accessing Power Query '{POWERQUERY_NAME}'...")
            pq = wb.Queries.Item(POWERQUERY_NAME)
        except Exception as e:
            print(f"DEBUG: Power Query '{POWERQUERY_NAME}' not found:", str(e))
            wb.Close(False)
            if excel:
                excel.Quit()
            return f"❌ Power Query '{POWERQUERY_NAME}' not found in workbook."

        try:
            old_formula = pq.Formula
            print("DEBUG: Old Power Query formula:", old_formula)
            last_step_match = re.findall(r'#"(.*?)"\s*=', old_formula)
            if last_step_match:
                last_step = f'#"{last_step_match[-1]}"'
            else:
                last_step = "Source"
            print("DEBUG: Last step detected:", last_step)

            new_step_name = f'Added_{column_name}'
            new_step = f'#"{new_step_name}" = Table.AddColumn({last_step}, "{column_name}", each {m_code})'
            new_formula = old_formula + f',\n{new_step}'
            pq.Formula = new_formula
            print("DEBUG: Power Query updated.")
        except Exception as e:
            print("DEBUG: Exception updating Power Query:", str(e))
            wb.Close(False)
            if excel:
                excel.Quit()
            return f"❌ Could not update Power Query: {str(e)}"

        wb.Save()
        wb.Close(False)
        if excel:
            excel.Quit()
        print("DEBUG: Workbook saved and closed.")
        return f"✅ Column added with M Code: {m_code}"
    except Exception as e:
        import traceback
        print("DEBUG: Exception in main try block:", str(e))
        traceback.print_exc()
        if wb:
            wb.Close(False)
        if excel:
            excel.Quit()
        return f"❌ Error: {str(e)}"

if __name__ == "__main__":
    try:
        mcp.run()
    except Exception as e:
        import traceback
        print("❌ MCP server failed to start:")
        traceback.print_exc()