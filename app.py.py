import streamlit as st
import os
from conservation_doc_formatter import format_document
import conservation_doc_formatter as cdf  

st.set_page_config(page_title="Report Formatter", layout="centered")

st.title("📄 Report Formatter")
st.write("Upload a draft Word document to automatically apply standard publication formatting.")

# --- UI Controls ---
uploaded_file = st.file_uploader("Upload .docx Draft", type="docx")

col1, col2 = st.columns(2)
with col1:
    theme_options = list(cdf.THEMES.keys()) + ["Custom Builder..."]
    theme_choice = st.selectbox("Select Theme", theme_options, index=0)
with col2:
    custom_label = st.text_input("Custom Footer Label", placeholder="e.g., Annual Summary")

# --- Advanced Theme Builder UI ---
active_theme = theme_choice

if theme_choice == "Custom Builder...":
    st.write("### 🎨 Advanced Theme Builder")
    
    tab1, tab2, tab3, tab4 = st.tabs(["Typography", "Core Branding", "Tables", "Text Elements"])
    
    with tab1:
        st.write("#### Document Typography")
        col_font1, col_font2 = st.columns(2)
        with col_font1:
            custom_font = st.selectbox("Primary Font", ["Arial", "Calibri", "Times New Roman", "Helvetica", "Garamond"])
        with col_font2:
            custom_size = st.slider("Base Font Size (pt)", min_value=9, max_value=14, value=11)
            # Convert standard pt size to Word's XML half-points
            xml_half_points = str(custom_size * 2)

    with tab2:
        col_a, col_b = st.columns(2)
        with col_a:
            primary = st.color_picker("Primary (Covers, Labels)", "#1E4D2B")
            accent = st.color_picker("Accent (Headings, Borders)", "#3D6B4F")
        with col_b:
            accent_light = st.color_picker("Light Accent (Subheadings)", "#6B9E7A")
            cover_sub = st.color_picker("Cover Subtitle", "#2B5F45")
            
    with tab3:
        col_c, col_d = st.columns(2)
        with col_c:
            table_header = st.color_picker("Table Header Background", "#D5E8D4")
            table_alt = st.color_picker("Table Alt Row Background", "#F2F8F0")
        with col_d:
            callout_bg = st.color_picker("Callout Box Background", "#1E4D2B")
            callout_text = st.color_picker("Callout Text Color", "#FFFFFF")
            
    with tab4:
        col_e, col_f = st.columns(2)
        with col_e:
            body = st.color_picker("Standard Body Text", "#1A1A1A")
            italic_quote = st.color_picker("Pull-Quote Text", "#404040")
        with col_f:
            caption = st.color_picker("Figure Captions", "#3D6B4F")
            footer_rule = st.color_picker("Footer Border Line", "#3D6B4F")

    # Map all custom variables, including fonts, to the formatter
    cdf.THEMES["custom_user"] = {
        "name": "Custom User Theme",
        "font": custom_font,
        "base_size": xml_half_points,
        "primary": primary.lstrip('#').upper(),
        "accent": accent.lstrip('#').upper(),
        "accent_light": accent_light.lstrip('#').upper(),
        "cover_sub": cover_sub.lstrip('#').upper(),
        "table_header": table_header.lstrip('#').upper(),
        "table_alt": table_alt.lstrip('#').upper(),
        "callout_bg": callout_bg.lstrip('#').upper(), 
        "callout_text": callout_text.lstrip('#').upper(),
        "caption": caption.lstrip('#').upper(), 
        "footer_rule": footer_rule.lstrip('#').upper(), 
        "body": body.lstrip('#').upper(),
        "italic_quote": italic_quote.lstrip('#').upper(),
    }
    active_theme = "custom_user"

# --- Document Processing ---
if uploaded_file is not None:
    if st.button("Format Document", type="primary"):
        input_path = f"temp_in_{uploaded_file.name}"
        output_path = f"Formatted_{uploaded_file.name}"
        
        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        label = custom_label if custom_label else uploaded_file.name.replace(".docx", "")
        
        with st.spinner("Applying formatting..."):
            try:
                format_document(input_path, output_path, theme=active_theme, report_label=label)
                
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="⬇️ Download Publication-Ready Report",
                        data=f,
                        file_name=output_path,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            except Exception as e:
                st.error(f"An error occurred: {e}")
            finally:
                if os.path.exists(input_path): os.remove(input_path)
                if os.path.exists(output_path): os.remove(output_path)
