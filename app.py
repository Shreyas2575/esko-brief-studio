import html
import io
import json
from datetime import datetime
from typing import List

import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Pt, RGBColor
import fitz  # PyMuPDF
import pandas as pd
from PIL import Image
from pydantic import BaseModel, Field
import streamlit as st
from google import genai
from google.genai import types

# ==========================================
# 1. PAGE CONFIG & ENTERPRISE SAAS STYLING
# ==========================================
st.set_page_config(
    page_title="Esko AI Packaging Brief Studio",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .hero-banner {
        background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 60%, #2563EB 100%);
        color: white;
        padding: 24px 30px;
        border-radius: 14px;
        margin-bottom: 20px;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
    }
    .hero-title {
        font-size: 1.9rem;
        font-weight: 800;
        margin: 0 0 6px 0;
        letter-spacing: -0.02em;
    }
    .hero-subtitle {
        font-size: 0.98rem;
        color: #DBEAFE;
        margin: 0;
        font-weight: 400;
    }
    .hero-pill {
        display: inline-block;
        background: rgba(255, 255, 255, 0.15);
        border: 1px solid rgba(255, 255, 255, 0.25);
        color: #F8FAFC;
        padding: 3px 12px;
        border-radius: 999px;
        font-size: 0.75rem;
        font-weight: 600;
        margin-bottom: 10px;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    .kpi-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 16px 18px;
        box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
        text-align: center;
    }
    .kpi-value {
        font-size: 1.75rem;
        font-weight: 800;
        color: #0F172A;
        line-height: 1.2;
    }
    .kpi-label {
        font-size: 0.8rem;
        font-weight: 600;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        margin-top: 4px;
    }
    .source-badge {
        display: inline-block;
        background-color: #EEF2FF;
        color: #3730A3;
        border: 1px solid #C7D2FE;
        border-radius: 999px;
        padding: 4px 12px;
        font-size: 0.76rem;
        font-weight: 600;
    }
    .user-badge {
        display: inline-block;
        background-color: #ECFDF5;
        color: #065F46;
        border: 1px solid #A7F3D0;
        border-radius: 999px;
        padding: 4px 12px;
        font-size: 0.76rem;
        font-weight: 600;
    }
    .conflict-card {
        border-left: 5px solid #EF4444;
        background: linear-gradient(90deg, #FEF2F2 0%, #FFFFFF 100%);
        border: 1px solid #FECACA;
        padding: 16px 20px;
        border-radius: 10px;
        margin-bottom: 12px;
        color: #1F2937;
    }
    .missing-card {
        border-left: 5px solid #F59E0B;
        background: linear-gradient(90deg, #FFFBEB 0%, #FFFFFF 100%);
        border: 1px solid #FDE68A;
        padding: 16px 20px;
        border-radius: 10px;
        margin-bottom: 12px;
        color: #1F2937;
    }
    .severity-pill-critical {
        background: #DC2626;
        color: white;
        padding: 2px 8px;
        border-radius: 6px;
        font-size: 0.72rem;
        font-weight: 700;
        margin-right: 8px;
        text-transform: uppercase;
    }
    .severity-pill-gap {
        background: #D97706;
        color: white;
        padding: 2px 8px;
        border-radius: 6px;
        font-size: 0.72rem;
        font-weight: 700;
        margin-right: 8px;
        text-transform: uppercase;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==========================================
# 2. PYDANTIC SCHEMAS
# ==========================================
class SourcedFact(BaseModel):
    field_name: str
    value: str
    source_file: str
    source_location: str
    confidence: str


class DiscrepancyOrGap(BaseModel):
    id: str
    category: str
    severity: str
    title: str
    explanation: str
    sources_involved: List[str]
    question_for_user: str
    suggested_options: List[str]


class PackagingAnalysisResult(BaseModel):
    product_summary: str
    extracted_facts: List[SourcedFact]
    issues_and_questions: List[DiscrepancyOrGap]


class BriefField(BaseModel):
    label: str
    value: str
    source_citation: str


class FinalAgencyBrief(BaseModel):
    project_title: str
    executive_summary: str
    project_and_brand_overview: List[BriefField]
    structural_and_pack_specs: List[BriefField]
    mandatory_copy_and_regulatory: List[BriefField]
    visual_and_creative_direction: List[BriefField]
    typography_and_print_corrections: List[str]
    open_questions_for_agency: List[str]


# ==========================================
# 3. DETERMINISTIC VALIDATORS
# ==========================================
def verify_upc_a(barcode_str: str) -> dict:
    digits = [int(c) for c in barcode_str if c.isdigit()]
    if len(digits) != 12:
        return {"valid": False, "msg": f"Non-standard length ({len(digits)} digits)"}
    odd_sum = sum(digits[0:11:2]) * 3
    even_sum = sum(digits[1:11:2])
    total = odd_sum + even_sum
    expected_check = (10 - (total % 10)) % 10
    actual_check = digits[11]
    if expected_check == actual_check:
        return {"valid": True, "msg": f"VALID UPC-A — Check digit {actual_check} verified."}
    return {"valid": False, "msg": f"INVALID UPC-A — Ends in {actual_check}, requires {expected_check}."}


# ==========================================
# 4. PARSERS & VISION EXTRACTORS
# ==========================================
def parse_uploaded_files(uploaded_files):
    gemini_parts = []
    file_summaries = []
    visual_assets = []

    for uf in uploaded_files:
        fname = uf.name
        raw_bytes = uf.read()

        if fname.lower().endswith(".pdf"):
            doc = fitz.open(stream=raw_bytes, filetype="pdf")
            pdf_text_blocks = []
            for page_idx in range(len(doc)):
                page = doc[page_idx]
                pdf_text_blocks.append(f"--- {fname} (Page {page_idx + 1}) ---\n{page.get_text()}")
                pix = page.get_pixmap(matrix=fitz.Matrix(2.2, 2.2))
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                gemini_parts.append(f"Visual artwork image from file '{fname}' Page {page_idx + 1}:")
                gemini_parts.append(img)
                visual_assets.append({
                    "title": f"{fname} — Page {page_idx + 1}",
                    "image": img,
                    "source": fname,
                })

            full_pdf_text = "\n".join(pdf_text_blocks)
            gemini_parts.append(f"Raw text layer from '{fname}':\n{full_pdf_text}")
            file_summaries.append({"name": fname, "type": "PDF Artwork", "details": f"{len(doc)} page(s)"})

        elif fname.lower().endswith((".xlsx", ".xls")):
            xls = pd.ExcelFile(io.BytesIO(raw_bytes))
            excel_dump = []
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet).fillna("")
                excel_dump.append(f"--- File: {fname} | Sheet: {sheet} ---\n{df.to_string(index=True)}")
            combined_excel = "\n\n".join(excel_dump)
            gemini_parts.append(f"Spreadsheet Data from '{fname}':\n{combined_excel}")
            file_summaries.append({"name": fname, "type": "Excel Spec", "details": f"{len(xls.sheet_names)} sheet(s)"})

        elif fname.lower().endswith((".png", ".jpg", ".jpeg", ".webp")):
            img = Image.open(io.BytesIO(raw_bytes))
            gemini_parts.append(f"Image '{fname}':")
            gemini_parts.append(img)
            visual_assets.append({"title": f"{fname}", "image": img, "source": fname})
            file_summaries.append({"name": fname, "type": "Image", "details": f"{img.size[0]}x{img.size[1]} px"})

        elif fname.lower().endswith(".docx"):
            doc = docx.Document(io.BytesIO(raw_bytes))
            text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
            gemini_parts.append(f"Word Document '{fname}':\n{text}")
            file_summaries.append({"name": fname, "type": "Word Doc", "details": f"{len(doc.paragraphs)} paragraphs"})

        else:
            text = raw_bytes.decode("utf-8", errors="ignore")
            gemini_parts.append(f"Text Document '{fname}':\n{text}")
            file_summaries.append({"name": fname, "type": "Text Note", "details": f"{len(text)} chars"})

    return gemini_parts, file_summaries, visual_assets


# ==========================================
# 5. GEMINI AI ENGINE WITH AUTO-FALLBACK
# ==========================================
def run_packaging_audit(api_key: str, model_name: str, gemini_parts: list, user_notes: str) -> PackagingAnalysisResult:
    client = genai.Client(api_key=api_key)
    system_prompt = """
    You are a Principal Packaging Technologist & Pre-Press Auditor at Esko.
    Inspect all uploaded packaging files and prepare verified data for a Design Agency Brief.
    Extract facts with exact source citations, detect contradictions/errors, and identify missing brief requirements.
    """
    contents = gemini_parts + [f"Additional Brand Manager Notes:\n{user_notes if user_notes else 'None.'}"]
    
    models_to_try = [model_name, "gemini-1.5-flash", "gemini-1.5-pro", "gemini-flash-latest"]
    last_error = None

    for m in models_to_try:
        try:
            response = client.models.generate_content(
                model=m,
                contents=contents,
                config=types.GenerateContentConfig(
                    system_instruction=system_prompt,
                    response_mime_type="application/json",
                    response_schema=PackagingAnalysisResult,
                    temperature=0.1,
                ),
            )
            return PackagingAnalysisResult.model_validate_json(response.text)
        except Exception as e:
            last_error = e
            continue
            
    raise last_error


def generate_final_brief(
    api_key: str,
    model_name: str,
    analysis: PackagingAnalysisResult,
    user_resolutions: dict,
    chat_history: list,
) -> FinalAgencyBrief:
    client = genai.Client(api_key=api_key)
    prompt = f"""
    Create a final, production-ready Design Agency Brief combining extracted facts with brand manager resolutions.
    --- EXTRACTED SOURCE FACTS ---
    {analysis.model_dump_json(indent=2)}
    --- BRAND MANAGER RESOLUTIONS ---
    {json.dumps(user_resolutions, indent=2)}
    --- CHAT ---
    {json.dumps(chat_history, indent=2)}
    """
    
    models_to_try = [model_name, "gemini-1.5-flash", "gemini-1.5-pro", "gemini-flash-latest"]
    last_error = None

    for m in models_to_try:
        try:
            response = client.models.generate_content(
                model=m,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=FinalAgencyBrief,
                    temperature=0.2,
                ),
            )
            return FinalAgencyBrief.model_validate_json(response.text)
        except Exception as e:
            last_error = e
            continue
            
    raise last_error


# ==========================================
# 6. EXPORT GENERATORS
# ==========================================
def brief_to_markdown(brief: FinalAgencyBrief) -> str:
    md = [f"# 📦 Packaging Design Brief: {brief.project_title}\n"]
    md.append(f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M')} | **Powered by:** Esko AI\n")
    md.append(f"**Executive Summary:** {brief.executive_summary}\n")
    sections = [
        ("1. Project & Brand Overview", brief.project_and_brand_overview),
        ("2. Structural & Pack Specifications", brief.structural_and_pack_specs),
        ("3. Mandatory Copy & Regulatory Content", brief.mandatory_copy_and_regulatory),
        ("4. Visual & Creative Direction", brief.visual_and_creative_direction),
    ]
    for title, fields in sections:
        md.append(f"## {title}")
        md.append("| Attribute | Verified Specification | Source Traceability |")
        md.append("| :--- | :--- | :--- |")
        for f in fields:
            md.append(f"| **{f.label}** | {f.value.replace(chr(10), ' ')} | `{f.source_citation}` |")
        md.append("")
    md.append("## 5. Mandatory Artwork Corrections")
    for item in brief.typography_and_print_corrections:
        md.append(f"- [ ] {item}")
    md.append("\n## 6. Open Questions for Agency")
    for q in brief.open_questions_for_agency:
        md.append(f"- ❓ {q}")
    return "\n".join(md)


def brief_to_docx_bytes(brief: FinalAgencyBrief) -> bytes:
    doc = docx.Document()
    title_para = doc.add_heading(f"PACKAGING DESIGN BRIEF: {brief.project_title}", level=0)
    for run in title_para.runs:
        run.font.color.rgb = RGBColor(15, 23, 42)
    meta = doc.add_paragraph(f"Prepared via Esko AI Packaging Brief Studio | Date: {datetime.now().strftime('%B %d, %Y')}")
    meta.runs[0].font.size = Pt(9.5)
    meta.runs[0].font.color.rgb = RGBColor(100, 116, 139)

    doc.add_heading("Executive Summary", level=1)
    doc.add_paragraph(brief.executive_summary)

    sections = [
        ("1. Project & Brand Overview", brief.project_and_brand_overview),
        ("2. Structural & Pack Specifications", brief.structural_and_pack_specs),
        ("3. Mandatory Copy & Regulatory Content", brief.mandatory_copy_and_regulatory),
        ("4. Visual & Creative Direction", brief.visual_and_creative_direction),
    ]
    for sec_title, fields in sections:
        doc.add_heading(sec_title, level=1)
        table = doc.add_table(rows=1, cols=3)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        hdr_cells = table.rows[0].cells
        for i, h in enumerate(["Attribute", "Verified Specification", "Source Traceability"]):
            hdr_cells[i].text = h
            hdr_cells[i].paragraphs[0].runs[0].font.bold = True
            hdr_cells[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(30, 58, 138)
        for item in fields:
            row_cells = table.add_row().cells
            row_cells[0].text = item.label
            row_cells[0].paragraphs[0].runs[0].font.bold = True
            row_cells[1].text = item.value
            row_cells[2].text = item.source_citation
            row_cells[2].paragraphs[0].runs[0].font.size = Pt(9)
            row_cells[2].paragraphs[0].runs[0].font.color.rgb = RGBColor(71, 85, 105)
        doc.add_paragraph()

    doc.add_heading("5. Mandatory Pre-Press & Artwork Corrections", level=1)
    for corr in brief.typography_and_print_corrections:
        doc.add_paragraph(f"☐  {corr}", style="List Bullet")

    doc.add_heading("6. Open Questions for the Design Agency", level=1)
    for q in brief.open_questions_for_agency:
        doc.add_paragraph(f"•  {q}", style="List Bullet")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def brief_to_styled_html(brief: FinalAgencyBrief) -> str:
    sections_html = ""
    sections = [
        ("1. Project & Brand Overview", brief.project_and_brand_overview),
        ("2. Structural & Pack Specifications", brief.structural_and_pack_specs),
        ("3. Mandatory Copy & Regulatory Content", brief.mandatory_copy_and_regulatory),
        ("4. Visual & Creative Direction", brief.visual_and_creative_direction),
    ]
    for sec_title, fields in sections:
        rows_html = ""
        for f in fields:
            is_user = "Brand Manager" in f.source_citation
            bg = "#ECFDF5" if is_user else "#EEF2FF"
            col = "#065F46" if is_user else "#3730A3"
            bdr = "#A7F3D0" if is_user else "#C7D2FE"
            rows_html += f"""
            <tr>
                <td style="width: 24%; font-weight: 700; color: #0F172A;">{html.escape(f.label)}</td>
                <td style="width: 50%; color: #1E293B; white-space: pre-line;">{html.escape(f.value)}</td>
                <td style="width: 26%;">
                    <span style="background:{bg}; color:{col}; border:1px solid {bdr}; padding:3px 10px; border-radius:999px; font-size:11px; font-weight:600;">
                        📌 {html.escape(f.source_citation)}
                    </span>
                </td>
            </tr>
            """
        sections_html += f"""
        <div style="margin-bottom: 24px;">
            <h2 style="font-size: 16px; color: #1E3A8A; border-bottom: 1px solid #E2E8F0; padding-bottom: 6px;">{html.escape(sec_title)}</h2>
            <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
                <thead>
                    <tr style="background: #F8FAFC; color: #475569; text-align: left;">
                        <th style="padding: 10px; border: 1px solid #CBD5E1;">Attribute</th>
                        <th style="padding: 10px; border: 1px solid #CBD5E1;">Verified Specification</th>
                        <th style="padding: 10px; border: 1px solid #CBD5E1;">Source Traceability</th>
                    </tr>
                </thead>
                <tbody>{rows_html}</tbody>
            </table>
        </div>
        """

    corr_items = "".join([f"<li><input type='checkbox'> {html.escape(c)}</li>" for c in brief.typography_and_print_corrections])
    open_q_items = "".join([f"<li>❓ {html.escape(q)}</li>" for q in brief.open_questions_for_agency])

    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>{html.escape(brief.project_title)}</title>
<style>
    body {{ font-family: sans-serif; color: #1E293B; background: #F8FAFC; padding: 28px; line-height: 1.5; }}
    .container {{ max-width: 960px; margin: 0 auto; background: #FFFFFF; padding: 36px; border-radius: 14px; border: 1px solid #E2E8F0; }}
    .print-btn {{ background: #2563EB; color: white; border: none; padding: 10px 18px; border-radius: 8px; font-weight: 700; cursor: pointer; }}
    @media print {{ .no-print {{ display: none !important; }} }}
</style>
</head>
<body>
<div class="container">
    <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 3px solid #1E3A8A; padding-bottom: 18px; margin-bottom: 24px;">
        <div>
            <div style="font-size: 12px; font-weight: 800; color: #2563EB; text-transform: uppercase;">ESKO AI BRIEF STUDIO</div>
            <h1 style="font-size: 24px; color: #0F172A; margin: 4px 0;">{html.escape(brief.project_title)}</h1>
        </div>
        <div class="no-print">
            <button class="print-btn" onclick="window.print()">🖨️ Print / Save as PDF</button>
        </div>
    </div>
    <div style="background: #F1F5F9; border-left: 4px solid #2563EB; padding: 14px 18px; border-radius: 6px; margin-bottom: 24px; font-size: 14px;">
        <strong>Executive Summary:</strong> {html.escape(brief.executive_summary)}
    </div>
    {sections_html}
    <div>
        <h2 style="font-size: 16px; color: #1E3A8A;">5. Mandatory Pre-Press Corrections</h2>
        <ul style="list-style: none; padding: 0;">{corr_items}</ul>
    </div>
    <div>
        <h2 style="font-size: 16px; color: #1E3A8A;">6. Open Questions for Agency</h2>
        <ul style="list-style: none; padding: 0;">{open_q_items}</ul>
    </div>
</div>
</body>
</html>"""


# ==========================================
# 7. MAIN STREAMLIT APPLICATION
# ==========================================
def main():
    for key, default in [
        ("analysis", None),
        ("user_resolutions", {}),
        ("final_brief", None),
        ("chat_history", []),
        ("visual_assets", []),
        ("file_summaries", []),
    ]:
        if key not in st.session_state:
            st.session_state[key] = default

    with st.sidebar:
        st.markdown("## 📦 Esko Brief Studio")
        st.caption("AI-Powered Packaging Pre-Press & Design Brief Engine")
        api_key = st.text_input("🔑 Gemini API Key", type="password", help="Enter your Google Gemini API key")
        model_name = st.selectbox(
            "🤖 AI Reasoning Model",
            options=["gemini-1.5-flash", "gemini-1.5-pro", "gemini-flash-latest"],
            index=0,
            help="Defaults to gemini-1.5-flash",
        )
        st.divider()
        if st.session_state.file_summaries:
            st.markdown("### 📂 Active Files")
            for fs in st.session_state.file_summaries:
                st.caption(f"✅ **{fs['name']}** ({fs['details']})")

    st.markdown(
        """
        <div class="hero-banner">
            <div class="hero-pill">✨ Esko AI Packaging Challenge</div>
            <div class="hero-title">From Messy Product Files to Agency-Ready Brief</div>
            <div class="hero-subtitle">Ingest dielines & spreadsheets • Audit regulatory contradictions • Export traceable briefs.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.session_state.analysis:
        res: PackagingAnalysisResult = st.session_state.analysis
        contradictions = [i for i in res.issues_and_questions if i.category == "CONTRADICTION"]
        missing_items = [i for i in res.issues_and_questions if i.category != "CONTRADICTION"]
        total_q = len(res.issues_and_questions)
        answered_q = len([v for v in st.session_state.user_resolutions.values() if str(v).strip()])
        readiness_pct = int((answered_q / total_q) * 100) if total_q > 0 else 100

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.markdown(f"<div class='kpi-card'><div class='kpi-value'>{len(st.session_state.file_summaries)}</div><div class='kpi-label'>Files</div></div>", unsafe_allow_html=True)
        k2.markdown(f"<div class='kpi-card'><div class='kpi-value' style='color:#2563EB;'>{len(res.extracted_facts)}</div><div class='kpi-label'>Facts</div></div>", unsafe_allow_html=True)
        k3.markdown(f"<div class='kpi-card'><div class='kpi-value' style='color:#DC2626;'>{len(contradictions)}</div><div class='kpi-label'>Conflicts</div></div>", unsafe_allow_html=True)
        k4.markdown(f"<div class='kpi-card'><div class='kpi-value' style='color:#D97706;'>{len(missing_items)}</div><div class='kpi-label'>Gaps</div></div>", unsafe_allow_html=True)
        k5.markdown(f"<div class='kpi-card'><div class='kpi-value' style='color:#059669;'>{readiness_pct}%</div><div class='kpi-label'>Ready</div></div>", unsafe_allow_html=True)
        st.write("")

    tab1, tab2, tab3, tab4 = st.tabs([
        "🖼️ 1. Upload & Visual Inspector",
        "🔍 2. Audit & Smart Questions",
        "💬 3. Copilot Refiner",
        "📋 4. Live Brief & Exports",
    ])

    with tab1:
        col_up, col_notes = st.columns([3, 2])
        with col_up:
            uploaded_files = st.file_uploader(
                "Upload Packaging PDF Dieline, Excel Spec, or Images",
                type=["pdf", "xlsx", "xls", "docx", "png", "jpg", "jpeg", "txt"],
                accept_multiple_files=True,
            )
        with col_notes:
            user_notes = st.text_area("Initial Launch Notes (Optional)", height=125)

        if st.button("🚀 Analyze Files & Launch Inspector", type="primary", use_container_width=True):
            if not api_key:
                st.error("Please enter your Gemini API Key in the left sidebar.")
            elif not uploaded_files:
                st.warning("Please upload at least one file.")
            else:
                with st.spinner("Processing dielines and running pre-press audit..."):
                    parts, summaries, visuals = parse_uploaded_files(uploaded_files)
                    st.session_state.file_summaries = summaries
                    st.session_state.visual_assets = visuals
                    st.session_state.analysis = run_packaging_audit(api_key, model_name, parts, user_notes)
                    st.session_state.final_brief = None
                st.success("Analysis complete! Proceed to Tab 2.")
                st.rerun()

        if st.session_state.analysis:
            res: PackagingAnalysisResult = st.session_state.analysis
            st.divider()
            st.info(res.product_summary)

            insp_left, insp_right = st.columns([1.15, 1.35], gap="large")
            with insp_left:
                st.markdown("#### 🎨 Visual Dieline Inspector")
                if st.session_state.visual_assets:
                    asset_titles = [a["title"] for a in st.session_state.visual_assets]
                    selected_title = st.selectbox("Select Page", options=asset_titles)
                    selected_asset = next(a for a in st.session_state.visual_assets if a["title"] == selected_title)

                    rotation = st.radio("🔄 Rotate Orientation:", options=[0, 90, 180, 270], horizontal=True)
                    img_to_show = selected_asset["image"]
                    if rotation != 0:
                        img_to_show = img_to_show.rotate(-rotation, expand=True)
                    st.image(img_to_show, use_container_width=True)
            with insp_right:
                st.markdown(f"#### 📑 Traced Facts ({len(res.extracted_facts)})")
                search_q = st.text_input("🔎 Filter Facts", placeholder="Search keyword...")
                rows = [
                    {"Attribute": f.field_name, "Value": f.value, "Source": f.source_file, "Location": f.source_location}
                    for f in res.extracted_facts
                    if not search_q or search_q.lower() in f"{f.field_name} {f.value} {f.source_file}".lower()
                ]
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True, height=380)

                barcodes = [f for f in res.extracted_facts if "barcode" in f.field_name.lower() or "upc" in f.field_name.lower()]
                if barcodes:
                    for b in barcodes:
                        check = verify_upc_a(b.value)
                        if check["valid"]:
                            st.success(f"**{b.value}**: {check['msg']}")
                        else:
                            st.error(f"**{b.value}**: {check['msg']}")

    with tab2:
        if not st.session_state.analysis:
            st.info("Analyze files in Tab 1 first.")
        else:
            res: PackagingAnalysisResult = st.session_state.analysis
            contradictions = [i for i in res.issues_and_questions if i.category == "CONTRADICTION"]
            missing_items = [i for i in res.issues_and_questions if i.category != "CONTRADICTION"]

            col_conf, col_gap = st.columns(2, gap="large")
            with col_conf:
                st.subheader(f"🔴 Contradictions ({len(contradictions)})")
                for issue in contradictions:
                    st.markdown(
                        f"""<div class="conflict-card"><span class="severity-pill-critical">{issue.severity}</span>
                        <strong>{issue.title}</strong><br>{issue.explanation}<br><small>📌 {', '.join(issue.sources_involved)}</small></div>""",
                        unsafe_allow_html=True,
                    )
                    choice = st.radio(f"❓ **{issue.question_for_user}**", options=issue.suggested_options + ["Custom..."], key=f"r_{issue.id}")
                    if choice == "Custom...":
                        st.session_state.user_resolutions[issue.title] = st.text_input("Custom input:", key=f"c_{issue.id}")
                    else:
                        st.session_state.user_resolutions[issue.title] = choice
                    st.divider()

            with col_gap:
                st.subheader(f"🟡 Missing Specs ({len(missing_items)})")
                for gap in missing_items:
                    st.markdown(
                        f"""<div class="missing-card"><span class="severity-pill-gap">GAP</span>
                        <strong>{gap.title}</strong><br>{gap.explanation}</div>""",
                        unsafe_allow_html=True,
                    )
                    choice = st.radio(f"❓ **{gap.question_for_user}**", options=gap.suggested_options + ["Custom..."], key=f"r_{gap.id}")
                    if choice == "Custom...":
                        st.session_state.user_resolutions[gap.title] = st.text_input("Custom input:", key=f"c_{gap.id}")
                    else:
                        st.session_state.user_resolutions[gap.title] = choice
                    st.divider()

            if st.button("✨ Synthesize Final Agency Brief", type="primary", use_container_width=True):
                with st.spinner("Synthesizing brief..."):
                    st.session_state.final_brief = generate_final_brief(
                        api_key, model_name, st.session_state.analysis, st.session_state.user_resolutions, st.session_state.chat_history
                    )
                st.success("Brief generated! Go to Tab 4.")

    with tab3:
        st.subheader("💬 Brief Copilot")
        for msg in st.session_state.chat_history:
            with st.chat_message(msg["role"]):
                st.write(msg["content"])
        user_msg = st.chat_input("Ask a question or give instructions...")
        if user_msg:
            st.session_state.chat_history.append({"role": "user", "content": user_msg})
            with st.chat_message("user"):
                st.write(user_msg)
            if api_key and st.session_state.analysis:
                client = genai.Client(api_key=api_key)
                reply = client.models.generate_content(
                    model=model_name,
                    contents=f"Extracted: {st.session_state.analysis.model_dump_json()}\nUser: {user_msg}",
                )
                bot_text = reply.text
            else:
                bot_text = "Recorded for brief synthesis."
            st.session_state.chat_history.append({"role": "assistant", "content": bot_text})
            with st.chat_message("assistant"):
                st.write(bot_text)

    with tab4:
        if not st.session_state.final_brief:
            st.info("Complete Tab 2 to unlock the live brief and exports.")
        else:
            brief: FinalAgencyBrief = st.session_state.final_brief
            st.markdown("#### 📤 Exports")
            exp1, exp2, exp3, exp4 = st.columns(4)
            with exp1:
                st.download_button("📄 Word (.docx)", data=brief_to_docx_bytes(brief), file_name="Esko_Brief.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, type="primary")
            with exp2:
                st.download_button("🖨️ HTML / PDF", data=brief_to_styled_html(brief), file_name="Esko_Brief.html", mime="text/html", use_container_width=True)
            with exp3:
                st.download_button("📝 Markdown", data=brief_to_markdown(brief), file_name="Esko_Brief.md", mime="text/markdown", use_container_width=True)
            with exp4:
                st.download_button("🗂️ JSON Audit", data=brief.model_dump_json(indent=2), file_name="Esko_Audit.json", mime="application/json", use_container_width=True)

            st.divider()
            view_mode = st.radio("View Mode:", options=["✏️ Editor", "👁️ Preview"], horizontal=True)
            if view_mode == "👁️ Preview":
                st.components.v1.html(brief_to_styled_html(brief), height=900, scrolling=True)
            else:
                brief.project_title = st.text_input("Project Title", value=brief.project_title)
                brief.executive_summary = st.text_area("Executive Summary", value=brief.executive_summary, height=85)

                def render_sec(title, fields, key):
                    st.subheader(title)
                    for idx, item in enumerate(fields):
                        c1, c2 = st.columns([3.2, 1.2])
                        with c1:
                            item.value = st.text_input(item.label, value=item.value, key=f"{key}_{idx}")
                        with c2:
                            cls = "user-badge" if "Brand Manager" in item.source_citation else "source-badge"
                            st.markdown(f"<div style='margin-top:28px;'><span class='{cls}'>📌 {html.escape(item.source_citation)}</span></div>", unsafe_allow_html=True)

                render_sec("1. Overview", brief.project_and_brand_overview, "s1")
                render_sec("2. Specs", brief.structural_and_pack_specs, "s2")
                render_sec("3. Regulatory", brief.mandatory_copy_and_regulatory, "s3")
                render_sec("4. Creative", brief.visual_and_creative_direction, "s4")

                st.subheader("5. Pre-Press Corrections")
                corr = st.text_area("Corrections (one per line)", value="\n".join(brief.typography_and_print_corrections), height=150)
                brief.typography_and_print_corrections = [l.strip() for l in corr.split("\n") if l.strip()]

                st.subheader("6. Open Questions")
                qs = st.text_area("Questions (one per line)", value="\n".join(brief.open_questions_for_agency), height=120)
                brief.open_questions_for_agency = [l.strip() for l in qs.split("\n") if l.strip()]


if __name__ == "__main__":
    main()