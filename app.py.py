import streamlit as st
import os
import conservation_doc_formatter as cdf
from conservation_doc_formatter import format_document

# ── Graceful fallback if deployed formatter predates PAGE_LAYOUTS / INDENT_PROFILES ──
try:
    from conservation_doc_formatter import PAGE_LAYOUTS, INDENT_PROFILES
except ImportError:
    PAGE_LAYOUTS = {
        "a4_standard": {"page_width": 11906, "page_height": 16838,
                        "margin_top": 1440, "margin_bottom": 1440,
                        "margin_left": 1701, "margin_right": 1701,
                        "margin_header": 720, "margin_footer": 720,
                        "content_width": 8504, "gutter": 0},
        "a4_wide":     {"page_width": 11906, "page_height": 16838,
                        "margin_top": 1080, "margin_bottom": 1080,
                        "margin_left": 1080, "margin_right": 1080,
                        "margin_header": 540, "margin_footer": 540,
                        "content_width": 9746, "gutter": 0},
        "us_letter":   {"page_width": 12240, "page_height": 15840,
                        "margin_top": 1440, "margin_bottom": 1440,
                        "margin_left": 1440, "margin_right": 1440,
                        "margin_header": 720, "margin_footer": 720,
                        "content_width": 9360, "gutter": 0},
    }
    INDENT_PROFILES = {
        "formal":    {"body_first_line": 0,   "body_left": 0,  "body_right": 0,
                      "pullquote_left": 504,  "pullquote_right": 504,
                      "callout_left": 180,    "callout_right": 180,
                      "bullet_left": 504,     "bullet_hanging": 252,
                      "number_left": 504,     "number_hanging": 252,
                      "caption_left": 0,      "caption_right": 0,
                      "heading1_left": 0,     "heading2_left": 0, "heading3_left": 0},
        "editorial": {"body_first_line": 360, "body_left": 0,  "body_right": 0,
                      "pullquote_left": 720,  "pullquote_right": 720,
                      "callout_left": 216,    "callout_right": 216,
                      "bullet_left": 576,     "bullet_hanging": 288,
                      "number_left": 576,     "number_hanging": 288,
                      "caption_left": 0,      "caption_right": 0,
                      "heading1_left": 0,     "heading2_left": 288, "heading3_left": 432},
        "compact":   {"body_first_line": 0,   "body_left": 0,  "body_right": 0,
                      "pullquote_left": 360,  "pullquote_right": 360,
                      "callout_left": 144,    "callout_right": 144,
                      "bullet_left": 432,     "bullet_hanging": 216,
                      "number_left": 432,     "number_hanging": 216,
                      "caption_left": 0,      "caption_right": 0,
                      "heading1_left": 0,     "heading2_left": 0, "heading3_left": 216},
    }

# Patch onto the module so format_document can reach them if it looks them up
if not hasattr(cdf, "PAGE_LAYOUTS"):
    cdf.PAGE_LAYOUTS    = PAGE_LAYOUTS
if not hasattr(cdf, "INDENT_PROFILES"):
    cdf.INDENT_PROFILES = INDENT_PROFILES


st.set_page_config(page_title="WCT Report Formatter", layout="centered")

st.title("📄 Report Formatter")
st.write("Upload a draft Word document to automatically apply standard publication formatting.")

# --- UI Controls ---
uploaded_file = st.file_uploader("Upload .docx Draft", type="docx")

col1, col2 = st.columns(2)
with col1:
    theme_options = list(cdf.THEMES.keys()) + ["Custom Builder..."]
    theme_choice = st.selectbox("Select Theme", theme_options, index=0)
with col2:
    custom_label = st.text_input("Custom Footer Label", placeholder="e.g., Q2 Intervention Summary")

# --- Page Layout + Indent Profile row ---
col3, col4 = st.columns(2)
with col3:
    layout_options = list(PAGE_LAYOUTS.keys())
    layout_labels  = {
        "a4_standard": "A4 Standard (1.18\" margins)",
        "a4_wide":     "A4 Wide (0.75\" margins)",
        "us_letter":   "US Letter (1\" margins)",
    }
    # Pre-select theme default if not custom
    default_layout = cdf.THEMES.get(theme_choice, {}).get("page_layout", "a4_standard") \
        if theme_choice != "Custom Builder..." else "a4_standard"
    layout_choice = st.selectbox(
        "Page Layout",
        layout_options,
        index=layout_options.index(default_layout),
        format_func=lambda k: layout_labels.get(k, k),
    )
with col4:
    indent_options = list(INDENT_PROFILES.keys())
    indent_labels  = {
        "formal":    "Formal (block paragraphs)",
        "editorial": "Editorial (first-line indent)",
        "compact":   "Compact (field reports)",
    }
    default_indent = cdf.THEMES.get(theme_choice, {}).get("indent_profile", "formal") \
        if theme_choice != "Custom Builder..." else "formal"
    indent_choice = st.selectbox(
        "Indentation Style",
        indent_options,
        index=indent_options.index(default_indent),
        format_func=lambda k: indent_labels.get(k, k),
    )

# --- Custom Theme Builder UI ---
active_theme = theme_choice

if theme_choice == "Custom Builder...":
    st.write("### 🎨 Design Your Custom Theme")
    st.caption("Pick your hex colors below to match any specific branding requirements.")

    c1, c2 = st.columns(2)
    with c1:
        primary      = st.color_picker("Primary (Covers, Callouts)", "#1C2B3A")
        accent       = st.color_picker("Accent (Headings, Borders)", "#2E5077")
        table_header = st.color_picker("Table Header", "#D0DCE8")
    with c2:
        accent_light = st.color_picker("Light Accent (Subheadings)", "#4A7FAD")
        cover_sub    = st.color_picker("Cover Subtitle", "#243447")
        table_alt    = st.color_picker("Table Alt Row", "#F2F5F8")

    cdf.THEMES["custom_user"] = {
        "name":          "Custom User Theme",
        "primary":       primary.lstrip('#').upper(),
        "accent":        accent.lstrip('#').upper(),
        "accent_light":  accent_light.lstrip('#').upper(),
        "cover_sub":     cover_sub.lstrip('#').upper(),
        "table_header":  table_header.lstrip('#').upper(),
        "table_alt":     table_alt.lstrip('#').upper(),
        "callout_bg":    primary.lstrip('#').upper(),
        "callout_text":  "FFFFFF",
        "caption":       accent.lstrip('#').upper(),
        "footer_rule":   accent.lstrip('#').upper(),
        "body":          "1A1A1A",
        "italic_quote":  "404040",
        # Use the UI-selected layout + indent for custom themes
        "page_layout":   layout_choice,
        "indent_profile": indent_choice,
    }
    active_theme = "custom_user"
else:
    # Override the theme's defaults with the UI selections
    # (take a copy so we don't mutate the base dict permanently)
    import copy
    cdf.THEMES[active_theme] = dict(cdf.THEMES[active_theme])
    cdf.THEMES[active_theme]["page_layout"]    = layout_choice
    cdf.THEMES[active_theme]["indent_profile"] = indent_choice

# --- Document Processing ---
if uploaded_file is not None:
    if st.button("Format Document", type="primary"):
        input_path  = f"temp_in_{uploaded_file.name}"
        output_path = f"Formatted_{uploaded_file.name}"

        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        label = custom_label if custom_label else uploaded_file.name.replace(".docx", "")

        with st.spinner("Applying formatting…"):
            try:
                format_document(input_path, output_path, theme=active_theme, report_label=label)

                with open(output_path, "rb") as f:
                    st.download_button(
                        label="⬇️ Download Publication-Ready Report",
                        data=f,
                        file_name=output_path,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
                st.success(
                    f"Formatted with **{layout_labels.get(layout_choice, layout_choice)}** layout "
                    f"and **{indent_labels.get(indent_choice, indent_choice)}** indentation."
                )
            except Exception as e:
                st.error(f"An error occurred: {e}")
            finally:
                if os.path.exists(input_path):  os.remove(input_path)
                if os.path.exists(output_path): os.remove(output_path)
