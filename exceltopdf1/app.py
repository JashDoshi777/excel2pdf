import streamlit as st
import pandas as pd
import requests
import json
from jinja2 import Template
from xhtml2pdf import pisa
import io
import datetime
import base64
import os

# --- CONFIGURATION ---
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")

OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
MODEL_NAME = "deepseek/deepseek-v3.2"

# --- ASSETS ---
LOCAL_LOGO_FILENAME = "logo.png"

# Fallback (Black Abstract)
FALLBACK_LOGO_B64 = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAYAAAAeP4ixAAAABmJLR0QA/wD/AP+gvaeTAAAACXBIWXMAAAsTAAALEwEAmpwYAAAAB3RJTUUH5QwWESEX2j0ADAAAAxpJREFUaN7tmr9rFEEQxz97l0TiFyxEQVAQY2OVRmzEzsJCsLTGwkrEQv+FvY2Ilb+FlcU/QKy0sfASxZQBpZBCsBAkXOSSe5uFvbvbvb2527uT+MHA7M3MfnznZ2Z2duA/FhAAq8AasAFsA5vARrW+BawBa8Cq+u0BfWAIGAPGgTHgMzCufwd4r34/B46B00Z7G9W6C+wDO8AecBioA31Vf8vo78fV+yFwBrwFLoB3wFf1exx4a7S/U617wCFwGDgKHAZ2K4Mto1+f1fsJcAm8Bi6BT+r3FPAJeG20v1Ote8AhcBg4ChwGdiuDLaNfn9X7CXAJvAYugU/q9xTwCXhttL9TrXvAIXAYOAocBnYrgy2jX5/V+wlwybH3CngDnAevgA/q9wTw0Wh/p1r3gEPgMHAUOAzz7J0C7402d6p1FzgEdoH9YF8Z7xn9+qzeT4Az4DVwAbwDvqrfc8B7o/2dat0DDoHDwFHgMLBbGWwZ/fqc2DsF3htt7lTrLnAI7AL7wb4y3jP69Vm9nwBnwGvgAngHfFW/54D3Rvs71boHHAKHgaPAYWC3Mtgy+vU5sXcq/d7cqcZ94AjYA/aB/crw0OjXZ/V+ApwBr4EL4B3wVf2eA94b7e9U6x5wCBwGjgKHgd3KYMvo1+fE3qn0e3OnGveBI2AP2Af2K8NDo1+f1fsJcAa8Bi6Ad8BX9XsOeG+0v1Ote8AhcBg4ChwGdiuDLaNfn9X7CXAJvAYugU/q9xTwCXhttL9TrXvAIXAYOAocBnYrgy2jX5/V+wlwybH3CngDnAevgA/q9wTw0Wh/p1r3gEPgMHAUOAzz7J0C7402d6p1FzgEdoH9YF8Z7xn9+qzeT4Az4DVwAbwDvqrfc8B7o/2dat0DDoHDwFHgMLBbGWwZ/fqc2DsF3htt7lTrLnAI7AL7wb4y3jP69Vm9nwBnwGvgAngHfFW/54D3Rvs71boHHAKHgaPAYWC3Mtgy+vU5sXcq/d7cqcZ94AjYA/aB/crw0OjXZ/V+ApwBr4EL4B3wVf2eA94b7e9U6x5wCBwGjgKHgd3KYMvo1+fE3qn0e3OnGveBI2AP2Af2K8NDo1+f1fsJcAa8Bi6Ad8BX9XsOeG+0v1Ote8AhcBg4ChwGdiuDLaNfnxN75z+I337vH5qk4QAAAABJRU5ErkJggg=="

# --- HTML TEMPLATE ---
# UPDATED: The <img> tag now uses {{ logo_path }} instead of base64 data
html_template_string = """
<!DOCTYPE html>
<html>
<head>
    <style>
        @page {
            size: A4;
            margin: 1.5cm;
            @frame footer_frame {
                -pdf-frame-content: footerContent;
                bottom: 1cm;
                margin-left: 1.5cm;
                margin-right: 1.5cm;
                height: 1cm;
            }
        }
        body { font-family: Helvetica; color: #333333; }
        
        /* --- TITLE PAGE STYLES --- */
        .title-page {
            text-align: center;
            padding-top: 120px;
            page-break-after: always;
        }
        
        .logo {
            width: 250px;
            height: auto;
            margin-bottom: 40px;
        }
        
        .main-title {
            font-size: 36px;
            color: #2c3e50; 
            margin-bottom: 15px;
            text-transform: uppercase;
            letter-spacing: 2px;
            font-weight: bold;
        }
        .sub-title {
            font-size: 14px;
            color: #7f8c8d;
            margin-top: 0;
            text-transform: uppercase;
        }
        
        /* --- REPORT CONTENT STYLES --- */
        table.header-layout { width: 100%; border-bottom: 2px solid #2c3e50; margin-bottom: 20px; }
        td.header-left { text-align: left; vertical-align: bottom; }
        td.header-right { text-align: right; vertical-align: bottom; color: #777; font-size: 10px; }
        
        h1.page-header { color: #2c3e50; font-size: 18px; margin: 0; padding-bottom: 5px; }
        
        h2 { color: #e67e22; font-size: 18px; margin-top: 30px; border-bottom: 1px solid #e67e22; padding-bottom: 5px; }
        
        table.data-table { width: 100%; border: 0.5px solid #ddd; margin-top: 10px; font-size: 11px; }
        
        th { background-color: #2c3e50; color: #ffffff; padding: 8px; text-align: left; font-weight: bold; }
        
        td { padding: 6px; border-bottom: 0.5px solid #ddd; vertical-align: top; }
        tr.even { background-color: #f9f9f9; }
    </style>
</head>
<body>

    <div class="title-page">
        <img src="{{ logo_path }}" class="logo"/>
        <h1 class="main-title">{{ report_title }}</h1>
        <div class="sub-title">GENERATED REPORT | {{ date }}</div>
    </div>

    <table class="header-layout">
        <tr>
            <td class="header-left"><h1 class="page-header">{{ report_title }} (Details)</h1></td>
            <td class="header-right">Generated via AI Reporter<br/>Date: {{ date }}</td>
        </tr>
    </table>

    {% for section in tables %}
        {% if section.headers and section.rows %}
        <h2>{{ section.title }}</h2>
        <table class="data-table" repeat="1">
            <thead>
                <tr>
                    {% for header in section.headers %}
                    <th>{{ header }}</th>
                    {% endfor %}
                </tr>
            </thead>
            <tbody>
                {% for row in section.rows %}
                <tr class="{{ loop.cycle('odd', 'even') }}">
                    {% for cell in row %}
                    <td>{{ cell }}</td>
                    {% endfor %}
                </tr>
                {% endfor %}
            </tbody>
        </table>
        {% endif %}
    {% endfor %}

    <div id="footerContent" style="text-align:center; color:#888; font-size:9pt;">
        Page <pdf:pagenumber>
    </div>
</body>
</html>
"""

# --- HELPER: Save Logo to Disk and Return PATH ---
def get_logo_path(uploaded_file):
    """
    Saves the best available logo to a temporary file and returns the ABSOLUTE PATH.
    xhtml2pdf prefers paths over base64 strings.
    """
    temp_filename = "temp_report_logo.png"
    abs_path = os.path.abspath(temp_filename)

    # 1. User Upload (Priority)
    if uploaded_file is not None:
        try:
            with open(abs_path, "wb") as f:
                f.write(uploaded_file.getvalue())
            return abs_path
        except Exception as e:
            print(f"Error saving uploaded logo: {e}")

    # 2. Local File (Preferred)
    if os.path.exists(LOCAL_LOGO_FILENAME):
        return os.path.abspath(LOCAL_LOGO_FILENAME)
        
    # 3. Fallback (Decode Base64 string to a file)
    try:
        # Split the header "data:image/png;base64," from the data
        if "base64," in FALLBACK_LOGO_B64:
            _, encoded = FALLBACK_LOGO_B64.split("base64,", 1)
        else:
            encoded = FALLBACK_LOGO_B64
            
        decoded_data = base64.b64decode(encoded)
        with open(abs_path, "wb") as f:
            f.write(decoded_data)
        return abs_path
    except Exception as e:
        print(f"Error processing fallback logo: {e}")
        return None

# --- HELPER: Normalize Data ---
def normalize_ai_output(data):
    if not isinstance(data, dict):
        return {"report_title": "Report", "date": "", "tables": []}
        
    data["report_title"] = str(data.get("report_title", "Report"))
    data["date"] = str(data.get("date", ""))
    
    valid_tables = []
    
    for table in data.get("tables", []):
        raw_headers = table.get("headers", [])
        if not raw_headers: continue
            
        headers = [str(h) if (h is not None and str(h).strip() != "") else " " for h in raw_headers]
        col_count = len(headers)
        
        raw_rows = table.get("rows", [])
        clean_rows = []
        
        for row in raw_rows:
            if not isinstance(row, list): continue
            
            str_row = [str(cell) if (cell is not None and str(cell).strip() != "") else " " for cell in row]
            
            if len(str_row) < col_count:
                str_row += [" "] * (col_count - len(str_row))
            elif len(str_row) > col_count:
                str_row = str_row[:col_count]
                
            clean_rows.append(str_row)
        
        if clean_rows:
            table["headers"] = headers
            table["rows"] = clean_rows
            valid_tables.append(table)
        
    data["tables"] = valid_tables
    return data

def get_ai_structure(excel_file, api_key):
    # 1. Read all sheets
    all_sheets = pd.read_excel(excel_file, sheet_name=None, header=None)
    
    master_tables = []
    final_report_title = "Consolidated Report"
    final_report_date = ""

    # UI Progress Bar
    progress_bar = st.progress(0)
    total_sheets = len(all_sheets)
    
    # 2. LOOP through each sheet individually to avoid Token Limits
    for i, (sheet_name, df) in enumerate(all_sheets.items()):
        if df.empty: 
            continue
            
        # Update progress
        progress_bar.progress((i + 1) / total_sheets)
        
        csv_text = df.to_csv(index=False)
        
        # Prompt for individual sheet
        prompt = f"""
        Analyze the raw Excel data for the sheet named "{sheet_name}".
        Format numbers with commas.
        
        RULES:
        1. Extract all meaningful tables. 
        2. Give each table a clear, descriptive title based on the data content.
        3. Do NOT prepend the "Sheet Name" to the table title.
        4. Return ONLY valid JSON.
        
        JSON Structure:
        {{
            "report_title": "Report Title", 
            "date": "YYYY-MM-DD",
            "tables": [ 
                {{ 
                    "title": "Financials", 
                    "headers": ["ColA","ColB"], 
                    "rows": [["1","2"], ["3", "4"]] 
                }}
            ]
        }}

        RAW DATA:
        {csv_text}
        """
        
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
        
        data = {
            "model": MODEL_NAME,
            "messages": [{"role": "user", "content": prompt}]
        }
        
        try:
            response = requests.post(OPENROUTER_URL, headers=headers, json=data)
            
            if response.status_code == 200:
                content = response.json()['choices'][0]['message']['content']
                clean_content = content.replace("```json", "").replace("```", "")
                
                # Parse JSON for this specific sheet
                sheet_data = json.loads(clean_content)
                
                # Normalize and clean
                safe_data = normalize_ai_output(sheet_data)
                
                # Add to Master List
                master_tables.extend(safe_data.get("tables", []))
                
                # Capture date/title from the first valid sheet found
                if not final_report_date and safe_data.get("date"):
                    final_report_date = safe_data["date"]
                if final_report_title == "Consolidated Report" and safe_data.get("report_title"):
                    final_report_title = safe_data["report_title"]
                    
        except Exception as e:
            print(f"Skipping sheet {sheet_name} due to error: {e}")
            continue

    # Return the combined result
    return {
        "report_title": final_report_title,
        "date": final_report_date,
        "tables": master_tables
    }

def create_pdf(data_context, logo_path):
    # Pass the FILE PATH, not the binary data
    data_context['logo_path'] = logo_path 
    template = Template(html_template_string)
    html_content = template.render(**data_context)
    
    pdf_buffer = io.BytesIO()
    # pisa.CreatePDF handles file paths better than base64 strings
    pisa_status = pisa.CreatePDF(io.StringIO(html_content), dest=pdf_buffer)
    
    if pisa_status.err:
        raise Exception("PDF generation failed.")
        
    pdf_buffer.seek(0)
    return pdf_buffer

# --- STREAMLIT UI ---
st.set_page_config(page_title="Professional Report Generator", page_icon="📄")

st.title("📄 AI Report Generator")
st.markdown("Upload Excel, customize title, get a professional PDF.")

with st.sidebar:
    st.header("Report Settings")
    if OPENROUTER_API_KEY:
        st.success("✅ API key loaded from environment")
    else:
        st.error("❌ OPENROUTER_API_KEY not set")
    
    st.markdown("---")
    st.subheader("Title Page Config")
    custom_title = st.text_input("Report Title", value="Annual Financial Report")
    
    if os.path.exists(LOCAL_LOGO_FILENAME):
        st.success(f"✅ Found local logo: {LOCAL_LOGO_FILENAME}")
    else:
        st.warning(f"⚠️ Local logo '{LOCAL_LOGO_FILENAME}' not found. Using fallback.")
    
    uploaded_logo = st.file_uploader("Override with new Logo (Optional)", type=['png', 'jpg', 'jpeg'])
    if uploaded_logo:
        st.image(uploaded_logo, caption="Logo Preview", width=100)

uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        st.info(f"Detected {len(sheet_names)} sheets: {', '.join(sheet_names)}")
        
        with st.expander("See raw data preview"):
            df_preview = pd.read_excel(uploaded_file, sheet_name=0, header=None)
            st.dataframe(df_preview.head().astype(str))
    except Exception:
        st.warning("Could not preview file.")

    if st.button("Generate Report", type="primary"):
        if not OPENROUTER_API_KEY:
            st.error("OpenRouter API key not found. Set it in Render Environment Variables.")
            st.stop()
        else:
            try:
                uploaded_file.seek(0)
                
                with st.spinner("Processing Logo..."):
                    # Process logo to get a valid FILE PATH
                    final_logo_path = get_logo_path(uploaded_logo)
                
                with st.spinner("AI is analyzing all sheets individually..."):
                    structured_data = get_ai_structure(uploaded_file, OPENROUTER_API_KEY)

                
                if custom_title:
                    structured_data["report_title"] = custom_title
                if not structured_data.get("date"):
                    structured_data["date"] = datetime.date.today().strftime("%Y-%m-%d")

                if not structured_data["tables"]:
                    st.warning("No valid data found in Excel.")
                    st.json(structured_data)
                else:
                    with st.spinner("Rendering Title Page & PDF..."):
                        pdf_bytes = create_pdf(structured_data, final_logo_path)
                    
                    st.success("Report generated successfully!")
                    
                    st.download_button(
                        label="Download Professional PDF 📥",
                        data=pdf_bytes,
                        file_name=f"{custom_title.replace(' ', '_')}.pdf",
                        mime="application/pdf"
                    )
                    
                    with st.expander("See JSON Data (Debug)"):
                        st.json(structured_data)

            except Exception as e:
                st.error(f"An error occurred: {e}")
                import traceback
                st.code(traceback.format_exc())
