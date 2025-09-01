# --- START OF FILE app.py ---

import streamlit as st
# MODIFIED: Removed DocxTemplate, RichText, convert. Added Jinja2 and WeasyPrint.
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
import datetime
import io
import os
import sys
import base64
from urllib.parse import quote
import re
import tempfile

# Platform-specific import for Windows COM initialization and Outlook control
# This section remains unchanged and will be ignored on Streamlit Cloud (Linux)
if sys.platform == 'win32':
    import pythoncom
    import win32com.client as win32

# --- CUSTOM CSS STYLES ---
def load_custom_css():
    """Apply modern custom styling to the Streamlit app"""
    # This function is unchanged. The CSS is excellent.
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

    /* --- NEW: Styles for the INLINE PDF viewer --- */
    .pdf-viewer-container {
        border: 1px solid #ddd;
        border-radius: 10px;
        padding: 1rem;
        margin-top: 1.5rem;
        background-color: #f9f9f9;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
    }
    .pdf-iframe {
        width: 100%;
        height: 800px; /* A fixed height is good for inline display */
        border: none;
        border-radius: 5px;
    }
    
    /* Hide Streamlit branding */
    .stDeployButton {display:none;} footer {visibility: hidden;} .stApp > header {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

# --- Helper Functions ---
# REMOVED: The create_checkbox_rt function is no longer needed. Jinja2 handles booleans directly.

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
    
    # MODIFIED: Removed docx_bytes from cleanup
    if 'pdf_bytes' in st.session_state: del st.session_state.pdf_bytes
    if 'preview_visible' in st.session_state: st.session_state.preview_visible = False
    st.success("✨ Form cleared successfully!")

# MODIFIED: This is the new document generation function
def generate_documents():
    """Generates a PDF file from an HTML template and stores its bytes in session_state."""
    try:
        progress_bar = st.progress(0, text="🔄 Initializing...")
        
        # 1. Set up Jinja2 environment to load the HTML template
        env = Environment(loader=FileSystemLoader('.'))
        template = env.get_template("template.html")
        
        progress_bar.progress(30, text="📝 Processing form data...")
        # 2. Prepare the context dictionary for the template
        context = {key: value for key, value in st.session_state.items()}
        context['date'] = st.session_state.inspection_date.strftime("%Y-%m-%d")
        
        # 3. Render the HTML with the context data
        progress_bar.progress(50, text="📄 Rendering HTML from template...")
        rendered_html = template.render(context)
        
        # 4. Convert the rendered HTML to PDF using WeasyPrint
        progress_bar.progress(80, text=" Generating PDF document...")
        pdf_bytes = HTML(string=rendered_html).write_pdf()
        st.session_state.pdf_bytes = pdf_bytes

        progress_bar.progress(100, text="✅ PDF generated successfully!")
        st.success("🎉 PDF is ready for download!")
        st.session_state.file_name_base = f"IR_{st.session_state.unit_name.replace(' ', '_')}_{context['date']}"
        st.session_state.preview_visible = False # Hide preview on re-generation
        import time; time.sleep(1); progress_bar.empty()
    except Exception as e:
        st.error(f"❌ An error occurred: {e}")
        st.warning("⚠️ Please ensure 'template.html' is in the same directory as the app.")


def email_with_attachment_local():
    """
    Saves the PDF to a temporary file and opens Outlook with the file attached.
    WARNING: This ONLY works when running the script on a local Windows machine.
    """
    # This function is unchanged. It's correctly guarded for Windows-only execution.
    if sys.platform != 'win32':
        st.error("This feature is only available on Windows.")
        return

    try:
        # Create a temporary file to hold the PDF data
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
        st.warning("Ensure Outlook is installed and you have `pywin32` library (`pip install pywin32`).")
    finally:
        # Clean up the temporary file
        if 'attachment_path' in locals() and os.path.exists(attachment_path):
            os.unlink(attachment_path)


# --- MAIN APPLICATION ---
# This entire section remains the same, as it defines the user interface.
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
# This entire section is unchanged.
st.markdown('<div class="form-card">', unsafe_allow_html=True)
st.markdown('<h2 class="section-header">🔍 Inspection Items</h2>', unsafe_allow_html=True)
st.markdown("**Please select all inspections that apply to your project:**")

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

with st.expander("🔥 **FIRE PROTECTION SYSTEMS**", expanded=False):
    st.markdown('<p class="category-header">🚒 Firefighting Pipe Work</p>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.checkbox("🔧 Installation for F.F pipe work (with coating)", key="chk_ff_pipe_work_installation_with_coating")
        st.checkbox("💪 Pressure test for F.F @16 bar for 4 H", key="chk_ff_pressure_test_16bar_4h")
        st.checkbox("💧 Drops installation for sprinklers", key="chk_ff_drops_installation_for_sprinklers")
        st.checkbox("🌊 Flushing for F.F network", key="chk_ff_flushing_network")
    with c2:
        st.checkbox("🚿 Sprinkler installation", key="chk_ff_sprinkler_installation")
        st.checkbox("🔗 Connecting with Mall Tie-in & Opening Valve", key="chk_ff_connecting_with_mall_tie_in_opening_valve")
        st.checkbox("🍳 Hood wet chemical system (F&B)", key="chk_hood_wet_chemical_system_fnb")
        st.checkbox("🧯 FM200 & CO2 systems, fire extinguishers, etc.", key="chk_fm200_co2_systems_fire_extinguishers_fire_search")

with st.expander("❄️ **HVAC SYSTEMS**", expanded=False):
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<p class="category-header">🌪️ HVAC Duct Work</p>', unsafe_allow_html=True)
        st.checkbox("🔧 Installation for duct work, Dampers & FCU", key="chk_hvac_duct_installation_dampers_fcu")
        st.checkbox("💡 Light or smoke test for duct work", key="chk_hvac_duct_light_smoke_test")
        st.checkbox("🛡️ Insulation for duct work & VD", key="chk_hvac_duct_insulation")
        st.checkbox("3️⃣ Installation of 3rd fix for duct work", key="chk_hvac_duct_installation_3rd_fix")
        st.checkbox("🌬️ Volume dumper, Grill, diffusers", key="chk_hvac_duct_volume_dumper_grill_diffusers")
        st.checkbox("🔄 Air outlet installation", key="chk_hvac_duct_air_outlet_installation")
        st.checkbox("🔬 Test & commission (HVAC Duct)", key="chk_hvac_duct_test_commission")
        st.checkbox("📋 Test & Balance certificate", key="chk_hvac_duct_test_balance_certificate")
    with c2:
        st.markdown('<p class="category-header">🧊 HVAC Chilled Pipe</p>', unsafe_allow_html=True)
        st.checkbox("🔧 Installation of chilling pipes with coating", key="chk_hvac_chilled_pipe_installation_with_coating")
        st.checkbox("💪 Pressure test for chilled water @12 bar for 4 H", key="chk_hvac_chilled_pipe_pressure_test_12bar_4h")
        st.checkbox("🔗 Installation for hook-up", key="chk_hvac_chilled_pipe_installation_for_hook_up")
        st.checkbox("🧪 Chemical treatment for chilled water", key="chk_hvac_chilled_pipe_chemical_treatment")
        st.checkbox("🛡️ Insulation for all pipes & hook-up", key="chk_hvac_chilled_pipe_insulation_all_pipes_hook_up")
        st.checkbox("🔗 Connecting with Mall Tie-in & Opening Valve", key="chk_hvac_chilled_pipe_connecting_with_mall_tie_in_opening_valve")

with st.expander("🚿 **PLUMBING SYSTEMS**", expanded=False):
    c1, c2 = st.columns(2)
    with c1:
        st.checkbox("🔧 Installation of drainage pipes", key="chk_plumbing_drainage_pipes_installation")
        st.checkbox("💧 Water test for drainage pipes", key="chk_plumbing_drainage_pipes_water_test")
        st.checkbox("🛡️ Network protection before & after civil work", key="chk_plumbing_network_protection_before_after_civil_work")
        st.checkbox("3️⃣ 3rd fix (valves & plumbing fixtures)", key="chk_plumbing_3rd_fix_valves_fixtures")
        st.checkbox("❄️ Installation of A.C drainpipes", key="chk_plumbing_ac_drainpipes_installation")
        st.checkbox("💧 Water test for A.C drainpipes 16 bar for 4H", key="chk_plumbing_ac_drainpipes_water_test_16bar_4h")
    with c2:
        st.checkbox("🚰 Installation of water supply pipes", key="chk_plumbing_water_supply_pipes_installation")
        st.checkbox("💪 Pressure test for water supply pipes 16 bar for 4H", key="chk_plumbing_water_supply_pipes_pressure_test_16bar_4h")
        st.checkbox("🔧 Installation for drainage pipes (re-test)", key="chk_plumbing_drainage_pipes_re_installation")
        st.checkbox("💧 Water test for drainage pipes (re-test)", key="chk_plumbing_drainage_pipes_re_water_test")
        st.checkbox("🚽 Installation for flush tank - toilets only", key="chk_plumbing_flush_tank_toilets_only_installation")
        st.checkbox("🔥 Installation for EWH", key="chk_plumbing_ewh_installation")

with st.expander("📜 **CERTIFICATES REQUIRED**", expanded=False):
    c1, c2 = st.columns(2)
    with c1:
        st.checkbox("🧪 Chemical treatment flushing for chilled water", key="chk_certificate_chemical_treatment_flushing_chilled_water")
        st.checkbox("📊 Test & balance report for HVAC system (air & water)", key="chk_certificate_test_balance_report_hvac_air_water")
        st.checkbox("⚡ Panel board test certificated", key="chk_certificate_panel_board_test")
    with c2:
        st.checkbox("🚨 Fire alarm certificate", key="chk_certificate_fire_alarm")
        st.checkbox("🚿 Plumbing pipes certificate", key="chk_certificate_plumbing_pipes")
        st.checkbox("🧯 Hood fire suppression system MEP Testing sign-off", key="chk_certificate_hood_fire_suppression_system_mep_testing_sign_off_energization_commissioning")

st.markdown('</div>', unsafe_allow_html=True)
st.markdown("---")

# --- ACTION BUTTONS ---
st.markdown('<h2 class="section-header">🎯 Actions</h2>', unsafe_allow_html=True)
btn_col1, btn_col2, _ = st.columns([3, 3, 6])
with btn_col1:
    # MODIFIED: Changed label to reflect PDF-only output
    st.button("🚀 Generate PDF Document", on_click=generate_documents, type="primary", use_container_width=True)
with btn_col2:
    st.button("🗑️ Clear Form", on_click=clear_form, type="secondary", use_container_width=True)

# --- DOWNLOAD & INTERACTION SECTION ---
if 'pdf_bytes' in st.session_state:
    st.markdown("---")
    st.markdown('<h2 class="section-header">📥 Download & Share</h2>', unsafe_allow_html=True)
    
    # MODIFIED: Removed the DOCX download button and adjusted columns
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
    
    # This expander is unchanged and will work perfectly.
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
# This section is unchanged and works perfectly with the new PDF generation method.
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