"""
Excel Power Query MCP Server with Claude Integration (Full Table.AddColumn Version)
Author: Ed Keating (Sohentix)
"""

import sys
sys.stdout.reconfigure(encoding='utf-8')

print("✅ Step 1: Importing libraries")
import os, re, traceback
from fastmcp import FastMCP
import win32com.client as win32
import anthropic

EXCEL_FILE_PATH = "C:\\Users\\edwar\\Documents\\Sales.xlsx"
ANTHROPIC_KEY = "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"
POWERQUERY_NAME = "Table1"

print("✅ Step 2: Initializing Claude  FastMCP")
client = anthropic.Anthropic(api_key=ANTHROPIC_KEY)
mcp = FastMCP("excel-pq-server")

@mcp.tool()
def PowerQuery_MCP(prompt: str) -> str:
    print("⚙️ PowerQuery_MCP invoked")

    if not (EXCEL_FILE_PATH and prompt.strip() and ANTHROPIC_KEY):
        return "❌ Missing required values"

    column_name = "Sales Status"
    step_name = "Added_" + re.sub(r'[^A-Za-z0-9]+', '_', column_name)
    print(f"📝 Column: {column_name}, Step: {step_name}")
    print(f"📁 Opening workbook: {EXCEL_FILE_PATH}")

    # Initialize COM first
    import pythoncom
    pythoncom.CoInitialize()

    excel = wb = None

    try:
        if not os.path.exists(EXCEL_FILE_PATH):
            return f"❌ Excel file not found: {EXCEL_FILE_PATH}"

        enhanced_prompt = f"""
{prompt}

Return ONLY a full, valid Power Query M expression using Table.AddColumn.
It must be a single line, like:

Table.AddColumn(#"Changed Type", "Sales Status", each if [Sales] > 1000 then "High" else "Low")

Do NOT include explanations, comments, or Markdown formatting. Just return the M code directly.
"""

        message = client.messages.create(
            model="claude-3-haiku-20240307",
            max_tokens=256,
            temperature=0,
            system="You are a Power Query M assistant. Return only valid full Table.AddColumn expressions.",
            messages=[{"role": "user", "content": enhanced_prompt}]
        )
        raw_response = message.content[0].text.strip()
        print("📤 Claude Response:", raw_response)

        m_code = raw_response.strip()

        # Fix COM cache issue - use Dispatch instead of gencache
        try:
            excel = win32.Dispatch("Excel.Application")
        except Exception:
            # If that fails, try clearing cache first
            import tempfile, shutil
            cache_dir = os.path.join(tempfile.gettempdir(), "gen_py")
            if os.path.exists(cache_dir):
                try:
                    shutil.rmtree(cache_dir)
                except:
                    pass
            excel = win32.Dispatch("Excel.Application")
        
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(EXCEL_FILE_PATH)
        if wb.ReadOnly:
            return f"❌ Workbook is read-only: {wb.FullName}"

        pq = wb.Queries.Item(POWERQUERY_NAME)
        old_formula = pq.Formula
        print("🧪 Old Power Query Formula:\n", old_formula)

        lines = old_formula.strip().splitlines()
        in_index = next((i for i, line in enumerate(lines) if line.strip().lower().startswith("in")), None)
        if in_index is None:
            return "❌ Invalid formula: missing 'in' clause."

        if not lines[in_index - 1].strip().endswith(","):
            lines[in_index - 1] = lines[in_index - 1].rstrip() + ","

        lines.insert(in_index, f'    #"{step_name}" = {m_code}')
        lines[in_index + 1] = f'in\n    #"{step_name}"'

        if len(lines) > in_index + 2 and lines[in_index + 2].strip().startswith('#"'):
            print("🧽 Removing leftover line after in clause:", lines[in_index + 2])
            del lines[in_index + 2]

        new_formula = "\n".join(lines)
        if old_formula.strip() == new_formula.strip():
            return "⚠️ No changes made. Formula is already up to date."

        print("🧪 New Power Query Formula:\n", new_formula)
        pq.Formula = new_formula
        wb.Save()
        print("📁 Workbook saved.")
        return f"✅ Column '{column_name}' added."

    except Exception as e:
        print(f"❌ Exception in PowerQuery_MCP: {e}")
        traceback.print_exc()
        return f"❌ Error: {str(e)}"

    finally:
        try:
            if wb:
                wb.Close(SaveChanges=False)
            if excel:
                excel.Quit()
        except Exception as e:
            print(f"⚠️ Cleanup error: {e}")
        import gc
        excel = None
        wb = None
        gc.collect()
        pythoncom.CoUninitialize()

@mcp.tool()
def ListPowerQueries() -> str:
    """List all Power Query names in the Excel workbook."""
    # Initialize COM first
    import pythoncom
    pythoncom.CoInitialize()
    
    excel = wb = None
    try:
        if not os.path.exists(EXCEL_FILE_PATH):
            return f"❌ Excel file not found: {EXCEL_FILE_PATH}"

        # Fix COM cache issue
        try:
            excel = win32.gencache.EnsureDispatch("Excel.Application")
        except AttributeError:
            # Clear cache and use regular Dispatch
            import tempfile, shutil
            cache_dir = os.path.join(tempfile.gettempdir(), "gen_py")
            if os.path.exists(cache_dir):
                shutil.rmtree(cache_dir)
            excel = win32.Dispatch("Excel.Application")
        
        excel.Visible = False
        wb = excel.Workbooks.Open(EXCEL_FILE_PATH)
        names = [q.Name for q in wb.Queries]
        wb.Close(False)
        excel.Quit()
        return f"📋 Available Power Queries: {names}"
    except Exception as e:
        return f"❌ Could not list queries: {e}"
    finally:
        try:
            if wb:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel:
                excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    print("🚀 Launching MCP Server: excel-pq-server")
    mcp.run()
