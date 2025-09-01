# --- START OF FILE app.py ---

import streamlit as st
import datetime
import io
import os
import sys
import base64
from urllib.parse import quote
import re
import tempfile

# MODIFIED: Import PyPDF2 and reportlab components
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

# Platform-specific import for Windows COM initialization and Outlook control
if sys.platform == 'win32':
    import pythoncom
    import win32com.client as win32

# --- CUSTOM CSS STYLES (Unchanged) ---
def load_custom_css():
    """Apply modern custom styling to the Streamlit app"""
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        font-family: 'Inter', sans-serif;
    }
    .main .block-container {
        background: rgba(255, 255, 255, 0.95);
        border-radius: 20px; padding: 2rem; margin-top: 2rem;
        box-shadow: 0 20px 40px rgba(0, 0, 0, 0.1);
    }
    .main-title {
        background: linear-gradient(135deg, #667eea, #764ba2);
        -webkit-background-clip: text; -webkit-text-fill-color: black;
        background-clip: text; font-size: 2.5rem; font-weight: 700; text-align: center;
    }
    .section-header {
        color: #2c3e50; font-size: 1.5rem; font-weight: 600; margin: 2rem 0 1rem 0;
        padding-bottom: 0.5rem; border-bottom: 3px solid #667eea;
    }
    .form-card {
        background: white; padding: 1.5rem; border-radius: 15px;
        box-shadow: 0 8px 25px rgba(0, 0, 0, 0.08); margin: 1rem 0;
    }
    .category-header {
        color: #667eea; font-size: 1.2rem; font-weight: 600; margin: 1.5rem 0 1rem 0;
        padding: 0.5rem 1rem; border-left: 4px solid #667eea;
        background: linear-gradient(90deg, rgba(102, 126, 234, 0.1), transparent);
        border-radius: 0 10px 10px 0;
    }
    .pdf-viewer-container {
        border: 1px solid #ddd; border-radius: 10px; padding: 1rem; margin-top: 1.5rem;
        background-color: #f9f9f9; box-shadow: 0 4px 12px rgba(0,0,0,0.05);
    }
    .pdf-iframe {
        width: 100%; height: 800px; border: none; border-radius: 5px;
    }
    .stDeployButton {display:none;} footer {visibility: hidden;} .stApp > header {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

# --- Helper Functions (Unchanged) ---
def sanitize_filename(name):
    """Removes characters that are invalid in Windows filenames."""
    return re.sub(r'[<>:"/\\|?*]', '_', name)

# --- CALLBACK FUNCTIONS ---
def clear_form():
    """Resets all form inputs in session_state to their default values."""
    defaults = {
        'unit_name': "Kenda Park St.", 'unit_num': "BS-03/BS-04", 'tenant': "",
        'serial_no': "IR-MEP-MEC/ELE/PLM-000", 'inspection_date': datetime.date.today(), 'email': "marwa_sanadily@parkst-eg.com"
    }
    for key, value in defaults.items(): st.session_state[key] = value
    for key in st.session_state.keys():
        if key.startswith("chk_"): st.session_state[key] = False
    
    if 'pdf_bytes' in st.session_state: del st.session_state.pdf_bytes
    if 'preview_visible' in st.session_state: st.session_state.preview_visible = False
    st.success("✨ Form cleared successfully!")


# --- NEW PDF GENERATION FUNCTION ---
def generate_documents():
    """Generates a PDF by stamping form data onto a template PDF."""
    try:
        progress_bar = st.progress(0, text="🔄 Initializing PDF generation...")

        # --- COORDINATES MAPPING ---
        # Maps keys to (x, y) coordinates on the PDF. Origin (0,0) is bottom-left.
        # These will need fine-tuning to match your template.pdf perfectly.
        # A standard Letter page is 612 points wide and 792 points high.
        CHECKBOX_X_POS = 508
        coordinates = {
            # --- Header Info ---
            'unit_name': (140, 715), 'unit_num': (430, 715),
            'tenant': (100, 690), 'email': (380, 690),
            'serial_no': (100, 665), 'date_formatted': (380, 665), # Use formatted date
            'date_in_sentence': (245, 627),
            
            # --- Checkboxes ---
            # CIVIL & STRUCTURAL (Page 1)
            'chk_signage_installation': (CHECKBOX_X_POS, 520),
            'chk_signage_actual_sample': (CHECKBOX_X_POS, 497),
            'chk_facade_junction_installation': (CHECKBOX_X_POS, 474),
            'chk_floor_water_proofing_facade_line': (CHECKBOX_X_POS, 451),
            'chk_floor_water_proofing_wat_area_1st_fix': (CHECKBOX_X_POS, 428),
            'chk_floor_water_proofing_wat_area_2nd_fix': (CHECKBOX_X_POS, 405),
            'chk_kitchen_drainage_piping_installation_inspection': (CHECKBOX_X_POS, 382),
            'chk_kitchen_tiles_and_grout_installation_inspection': (CHECKBOX_X_POS, 359),
            'chk_releasing_ceiling_closure_inspection': (CHECKBOX_X_POS, 336),
            'chk_flooring_installation_release': (CHECKBOX_X_POS, 313),
            'chk_rcp_3rd_fix': (CHECKBOX_X_POS, 290),
            'chk_mep_final_inspection': (CHECKBOX_X_POS, 267),
            'chk_ad_shaft_ceiling': (CHECKBOX_X_POS, 244),
            'chk_facade_isolation': (CHECKBOX_X_POS, 221),
            'chk_upstand_isolation': (CHECKBOX_X_POS, 198),

            # ELECTRICAL SYSTEMS (Assumed to be on Page 2, adjust if not)
            # To handle multiple pages, a more complex logic would be needed.
            # This example assumes a very long single page or that you map to page 2 coords.
            # Let's assume the template is a single long page for simplicity here.
            
            # ELECTRICAL (Lighting & Power)
            'chk_electrical_1st_fix': (CHECKBOX_X_POS, 140),
            'chk_electrical_2nd_fix': (CHECKBOX_X_POS, 117),
            'chk_electrical_3rd_fix': (CHECKBOX_X_POS, 94),
            'chk_panel_board_installation_termination': (CHECKBOX_X_POS, 71),
            'chk_electrical_test_commission': (CHECKBOX_X_POS, 48),
            
            # This is where coordinates would continue onto a second page
            # For this example, we'll continue decrementing Y as if it were one long page.
            # This will require you to create a 2-page template.pdf
        }
        
        # A more complete mapping would be needed for all checkboxes.
        # This is a representative sample. You can add the rest following the pattern.

        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=letter)
        c.setFont("Helvetica", 10)

        progress_bar.progress(30, text="📝 Placing form data on PDF...")

        # Special handling for dates
        st.session_state['date_formatted'] = st.session_state.inspection_date.strftime("%Y-%m-%d")
        st.session_state['date_in_sentence'] = st.session_state.inspection_date.strftime("%B %d, %Y")

        # Add text fields
        for key, (x, y) in coordinates.items():
            if not key.startswith("chk_") and key in st.session_state:
                c.drawString(x, y, str(st.session_state[key]))
        
        # Add checkmarks ("X")
        c.setFont("Helvetica-Bold", 14)
        for key, (x, y) in coordinates.items():
            if key.startswith("chk_") and st.session_state.get(key, False):
                c.drawString(x, y + 2, "X") # Small offset for better centering

        c.save()
        packet.seek(0)
        
        progress_bar.progress(60, text=" merging data with template...")
        
        stamp_pdf = PdfReader(packet)
        with open("template.pdf", "rb") as f:
            template_pdf = PdfReader(f)
            writer = PdfWriter()
            page = template_pdf.pages[0]
            page.merge_page(stamp_pdf.pages[0])
            writer.add_page(page)
            
            # If your template has more pages, add them back
            for i in range(1, len(template_pdf.pages)):
                writer.add_page(template_pdf.pages[i])

            output_buffer = io.BytesIO()
            writer.write(output_buffer)
            st.session_state.pdf_bytes = output_buffer.getvalue()

        progress_bar.progress(100, text="✅ PDF generated successfully!")
        st.success("🎉 PDF is ready for download!")
        st.session_state.file_name_base = f"IR_{st.session_state.unit_name.replace(' ', '_')}_{st.session_state.date_formatted}"
        st.session_state.preview_visible = False 
        import time; time.sleep(1); progress_bar.empty()

    except FileNotFoundError:
        st.error("❌ Error: 'template.pdf' not found. Please ensure the template file is in the same directory.")
    except Exception as e:
        st.error(f"❌ An error occurred: {e}")


def email_with_attachment_local():
    """Saves PDF to a temp file and opens Outlook. Windows only."""
    if sys.platform != 'win32':
        st.error("This feature is only available on Windows.")
        return
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(st.session_state.pdf_bytes)
            attachment_path = tmp.name
        pythoncom.CoInitialize()
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = st.session_state.email_to
        mail.Subject = st.session_state.email_subject
        mail.Body = st.session_state.email_body
        serial_no = st.session_state.get("serial_no", "Inspection_Request").strip()
        sanitized_name = sanitize_filename(serial_no)
        attachment_display_name = f"{sanitized_name}.pdf"
        mail.Attachments.Add(attachment_path, DisplayName=attachment_display_name, Type=1)
        mail.Display(True)
        st.success("Outlook email created with attachment!")
    except Exception as e:
        st.error(f"Failed to create Outlook email: {e}")
    finally:
        if 'attachment_path' in locals() and os.path.exists(attachment_path):
            os.unlink(attachment_path)

# --- MAIN APPLICATION UI ---
st.set_page_config(layout="wide", page_title="Inspection Request Form", page_icon="📋")
load_custom_css()
st.markdown('<h1 class="main-title">📋 Kenda Park St. Inspection Request Form</h1>', unsafe_allow_html=True)

# --- Form Section ---
st.markdown('<div class="form-card">', unsafe_allow_html=True)
st.markdown('<h2 class="section-header">🏢 Unit and Tenant Information</h2>', unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    st.text_input("🏪 Unit Name", "Kenda", key="unit_name", help="Enter the name of the business unit")
    st.text_input("🔢 Unit Number", "BS-03", key="unit_num", help="Building and unit identifier")
    st.text_input("👤 Tenant/TAR", "", key="tenant", help="Tenant or Tenant Authorized Representative")
with col2:
    st.text_input("📄 Serial Number", "IR-MEP-MEC-001", key="serial_no", help="Unique inspection request identifier")
    st.date_input("📅 Inspection Date", datetime.date.today(), key="inspection_date", help="Preferred inspection date")
    st.text_input("📧 Email", "marwa_sanadily@parkst-eg.com", key="email", help="Contact email for notifications")
st.markdown('</div>', unsafe_allow_html=True)

# --- Inspection Items Section ---
st.markdown('<div class="form-card">', unsafe_allow_html=True)
st.markdown('<h2 class="section-header">🔍 Inspection Items</h2>', unsafe_allow_html=True)
st.markdown("**Please select all inspections that apply to your project:**")

# NOTE: The `key` of each checkbox must match a key in the `coordinates` dictionary above.
with st.expander("🏗️ **CIVIL & STRUCTURAL**", expanded=True):
    c1, c2 = st.columns(2)
    with c1:
        st.checkbox("🪧 Signage installation", key="chk_signage_installation")
        st.checkbox("📋 Signage actual sample", key="chk_signage_actual_sample")
        st.checkbox("🔗 Facade junction installation", key="chk_facade_junction_installation")
        st.checkbox("💧 Floor water proofing - façade line", key="chk_floor_water_proofing_facade_line")
        st.checkbox("🚿 Floor water proofing - wet area 1st fix @ 60 cm height", key="chk_floor_water_proofing_wat_area_1st_fix")
        st.checkbox("🔧 Floor water proofing - wet area 2nd fix", key="chk_floor_water_proofing_wat_area_2nd_fix")
        st.checkbox("🍽️ Kitchen drainage piping installation", key="chk_kitchen_drainage_piping_installation_inspection")
        st.checkbox("🔲 Kitchen tiles and grout installation", key="chk_kitchen_tiles_and_grout_installation_inspection")
    with c2:
        st.checkbox("🏠 Releasing Ceiling closure inspection", key="chk_releasing_ceiling_closure_inspection")
        st.checkbox("🔲 Flooring installation-release", key="chk_flooring_installation_release")
        st.checkbox("📐 RCP - 3rd fix", key="chk_rcp_3rd_fix")
        st.checkbox("⚡ MEP final inspection", key="chk_mep_final_inspection")
        st.checkbox("🏗️ AD - shaft & ceiling", key="chk_ad_shaft_ceiling")
        st.checkbox("🏢 Façade isolation", key="chk_facade_isolation")
        st.checkbox("⬆️ Upstand isolation", key="chk_upstand_isolation")

with st.expander("⚡ **ELECTRICAL SYSTEMS**", expanded=False):
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<p class="category-header">💡 Lighting & Power</p>', unsafe_allow_html=True)
        st.checkbox("1️⃣ 1st fix (Electrical)", key="chk_electrical_1st_fix")
        st.checkbox("2️⃣ 2nd fix (Electrical)", key="chk_electrical_2nd_fix")
        st.checkbox("3️⃣ 3rd fix (Electrical)", key="chk_electrical_3rd_fix")
        st.checkbox("📊 Panel board installation & Termination", key="chk_panel_board_installation_termination")
        st.checkbox("🔬 Test & commission (Electrical)", key="chk_electrical_test_commission")
        st.markdown('<p class="category-header">📡 Light Current (Data, Tel, etc.)</p>', unsafe_allow_html=True)
        st.checkbox("1️⃣ 1st fix (Light Current)", key="chk_light_current_1st_fix")
        st.checkbox("2️⃣ 2nd fix (Light Current)", key="chk_light_current_2nd_fix")
        st.checkbox("3️⃣ 3rd fix (Light Current)", key="chk_light_current_3rd_fix")
    with c2:
        st.markdown('<p class="category-header">🚨 Fire Alarm</p>', unsafe_allow_html=True)
        st.checkbox("1️⃣ First fix (Fire Alarm)", key="chk_fire_alarm_first_fix")
        st.checkbox("2️⃣ Second fix (Fire Alarm)", key="chk_fire_alarm_second_fix")
        st.checkbox("3️⃣ Third fix (Fire Alarm)", key="chk_fire_alarm_third_fix")
        st.checkbox("🔬 Test & commission (Fire Alarm)", key="chk_fire_alarm_test_commission")
        st.checkbox("🔗 Interface with Mall (Fire Alarm)", key="chk_interface_with_mall")

# ... (The rest of your expanders and checkboxes go here. Ensure each has a unique 'key') ...
# You will need to add the coordinates for these keys in the `generate_documents` function.

st.markdown('</div>', unsafe_allow_html=True)
st.markdown("---")

# --- ACTION BUTTONS ---
st.markdown('<h2 class="section-header">🎯 Actions</h2>', unsafe_allow_html=True)
btn_col1, btn_col2, _ = st.columns([3, 3, 6])
with btn_col1:
    st.button("🚀 Generate PDF Document", on_click=generate_documents, type="primary", use_container_width=True)
with btn_col2:
    st.button("🗑️ Clear Form", on_click=clear_form, type="secondary", use_container_width=True)

# --- DOWNLOAD & INTERACTION SECTION ---
if 'pdf_bytes' in st.session_state:
    st.markdown("---")
    st.markdown('<h2 class="section-header">📥 Download & Share</h2>', unsafe_allow_html=True)
    
    dl_col1, dl_col2 = st.columns(2)
    with dl_col1:
        st.download_button(label="📄 Download PDF (.pdf)", data=st.session_state.pdf_bytes,
                           file_name=f"{st.session_state.file_name_base}.pdf",
                           mime="application/pdf", use_container_width=True)
    with dl_col2:
        def toggle_preview():
            st.session_state.preview_visible = not st.session_state.get('preview_visible', False)
        preview_text = "🔼 Hide Preview" if st.session_state.get('preview_visible', False) else "👁️ Show Inline Preview"
        st.button(preview_text, on_click=toggle_preview, use_container_width=True)
    
    with st.expander("📧 Email Documents via Outlook"):
        st.text_input("Recipient's Email", st.session_state.get('email', ''), key="email_to")
        st.text_input("Subject", f"Inspection Request: {st.session_state.unit_name} - {st.session_state.serial_no}", key="email_subject")
        st.text_area("Email Body", f"Dear Team,\n\nPlease find the inspection request attached for your review.\n\nUnit Name: {st.session_state.unit_name}\nInspection Date: {st.session_state.inspection_date.strftime('%Y-%m-%d')}\n\nThank you.", height=200, key="email_body")
        email_btn_col1, email_btn_col2 = st.columns(2)
        with email_btn_col1:
            mailto_url = f"mailto:{st.session_state.email_to}?subject={quote(st.session_state.email_subject)}&body={quote(st.session_state.email_body)}"
            st.markdown(f'<a href="{mailto_url}" target="_blank" style="display:block;padding:0.75rem;background:#555;color:white;text-decoration:none;border-radius:8px;text-align:center;">📬 Open & Attach Manually</a>', unsafe_allow_html=True)
            st.caption("Best for web/deployed apps.")
        with email_btn_col2:
            if sys.platform == 'win32':
                st.button("📎 Open in Outlook (with Attachment)", on_click=email_with_attachment_local, use_container_width=True)
                st.caption("Requires local Windows & Outlook.")

# --- INLINE PDF PREVIEW ---
if st.session_state.get('preview_visible', False) and 'pdf_bytes' in st.session_state:
    base64_pdf = base64.b64encode(st.session_state.pdf_bytes).decode('utf-8')
    st.markdown(f"""
        <div class="pdf-viewer-container">
            <iframe src="data:application/pdf;base64,{base64_pdf}" class="pdf-iframe"></iframe>
        </div>
    """, unsafe_allow_html=True)

# --- FOOTER ---
st.markdown("---")
st.markdown("*Powered by Streamlit • Built for efficient inspection management*")