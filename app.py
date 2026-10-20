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
    /* Main Hero Header */
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

    /* KPI Metric Cards */
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

    /* Source & User Traceability Badges */
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

    /* Issue & Conflict Resolution Cards */
    .conflict-card {
        border-left: 5px solid #EF4444;
        background: linear-gradient(90deg, #FEF2F2 0%, #FFFFFF 100%);
        border-top: 1px solid #FECACA;
        border-right: 1px solid #FECACA;
        border-bottom: 1px solid #FECACA;
        padding: 16px 20px;
        border-radius: 10px;
        margin-bottom: 12px;
        color: #1F2937;
    }
    .missing-card {
        border-left: 5px solid #F59E0B;
        background: linear-gradient(90deg, #FFFBEB 0%, #FFFFFF 100%);
        border-top: 1px solid #FDE68A;
        border-right: 1px solid #FDE68A;
        border-bottom: 1px solid #FDE68A;
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
# 2. PYDANTIC SCHEMAS FOR STRUCTURED AI OUTPUT
# ==========================================
class SourcedFact(BaseModel):
    field_name: str = Field(description="Packaging attribute name, e.g., Brand, Net Weight, UPC Barcode, Ingredients")
    value: str = Field(description="Extracted value from the file")
    source_file: str = Field(description="Exact file name where this value was found")
    source_location: str = Field(description="Panel name (e.g., Back Panel, Top Flap) or Excel Row/Column")
    confidence: str = Field(description="High, Medium, or Low")


class DiscrepancyOrGap(BaseModel):
    id: str = Field(description="Short unique ID, e.g., CONF-1 or GAP-1")
    category: str = Field(description="'CONTRADICTION' for conflicting/erroneous data, or 'MISSING_INFO' for missing brief essentials")
    severity: str = Field(description="'Critical', 'High', or 'Medium'")
    title: str = Field(description="Short title of the conflict or missing requirement")
    explanation: str = Field(description="Detailed explanation referencing exact conflicting values and sources")
    sources_involved: List[str] = Field(description="List of file names and panels/rows involved")
    question_for_user: str = Field(description="Clear, actionable question asking the Brand Manager how to resolve this")
    suggested_options: List[str] = Field(description="2 to 4 concrete resolution options the user can click to resolve")


class PackagingAnalysisResult(BaseModel):
    product_summary: str = Field(description="Concise executive overview of the product and existing packaging state")
    extracted_facts: List[SourcedFact]
    issues_and_questions: List[DiscrepancyOrGap]


class BriefField(BaseModel):
    label: str
    value: str
    source_citation: str = Field(description="Exact source citation, e.g. 'Carton PDF (Back Panel)' or 'Confirmed by Brand Manager'")


class FinalAgencyBrief(BaseModel):
    project_title: str
    executive_summary: str
    project_and_brand_overview: List[BriefField]
    structural_and_pack_specs: List[BriefField]
    mandatory_copy_and_regulatory: List[BriefField]
    visual_and_creative_direction: List[BriefField]
    typography_and_print_corrections: List[str] = Field(description="Specific artwork typos, barcode fixes, or regulatory formatting fixes required")
    open_questions_for_agency: List[str]


# ==========================================
# 3. DETERMINISTIC VALIDATORS (BARCODE & MATH)
# ==========================================
def verify_upc_a(barcode_str: str) -> dict:
    """Deterministic modulo-10 check for 12-digit UPC-A barcodes."""
    digits = [int(c) for c in barcode_str if c.isdigit()]
    if len(digits) != 12:
        return {"valid": False, "msg": f"Non-standard length ({len(digits)} digits)"}
    odd_sum = sum(digits[0:11:2]) * 3
    even_sum = sum(digits[1:11:2])
    total = odd_sum + even_sum
    expected_check = (10 - (total % 10)) % 10
    actual_check = digits[11]
    if expected_check == actual_check:
        return {
            "valid": True,
            "msg": f"VALID UPC-A — Check digit {actual_check} is mathematically verified.",
        }
    return {
        "valid": False,
        "msg": f"INVALID UPC-A — Ends in {actual_check}, but Modulo-10 math requires check digit {expected_check}.",
    }


# ==========================================
# 4. MULTI-MODAL FILE PARSER & VISUAL RENDERER
# ==========================================
def parse_uploaded_files(uploaded_files):
    """Extracts high-res images for visual inspection alongside structured text/tables for Gemini."""
    gemini_parts = []
    file_summaries = []
    visual_assets = []  # Stores {"title": str, "image": PIL.Image, "source": str}

    for uf in uploaded_files:
        fname = uf.name
        raw_bytes = uf.read()

        if fname.lower().endswith(".pdf"):
            doc = fitz.open(stream=raw_bytes, filetype="pdf")
            pdf_text_blocks = []
            for page_idx in range(len(doc)):
                page = doc[page_idx]
                pdf_text_blocks.append(
                    f"--- {fname} (Page {page_idx + 1} Extracted Text) ---\n{page.get_text()}"
                )
                # Render at 2.2x zoom for crisp visual inspection in Tab 1
                pix = page.get_pixmap(matrix=fitz.Matrix(2.2, 2.2))
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                gemini_parts.append(f"Visual artwork image from file '{fname}' Page {page_idx + 1}:")
                gemini_parts.append(img)
                visual_assets.append({
                    "title": f"{fname} — Page {page_idx + 1} (Dieline / Artwork)",
                    "image": img,
                    "source": fname,
                })

            full_pdf_text = "\n".join(pdf_text_blocks)
            gemini_parts.append(f"Raw text layer from '{fname}':\n{full_pdf_text}")
            file_summaries.append({"name": fname, "type": "PDF Artwork / Dieline", "details": f"{len(doc)} page(s)"})

        elif fname.lower().endswith((".xlsx", ".xls")):
            xls = pd.ExcelFile(io.BytesIO(raw_bytes))
            excel_dump = []
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet).fillna("")
                excel_dump.append(f"--- File: {fname} | Sheet: {sheet} ---\n{df.to_string(index=True)}")
            combined_excel = "\n\n".join(excel_dump)
            gemini_parts.append(f"Structured Spreadsheet Data from '{fname}':\n{combined_excel}")
            file_summaries.append({"name": fname, "type": "Excel Spec Sheet", "details": f"{len(xls.sheet_names)} sheet(s)"})

        elif fname.lower().endswith((".png", ".jpg", ".jpeg", ".webp")):
            img = Image.open(io.BytesIO(raw_bytes))
            gemini_parts.append(f"Uploaded Packaging Image '{fname}':")
            gemini_parts.append(img)
            visual_assets.append({
                "title": f"{fname} (Packaging Image)",
                "image": img,
                "source": fname,
            })
            file_summaries.append({"name": fname, "type": "Image Asset", "details": f"{img.size[0]}x{img.size[1]} px"})

        elif fname.lower().endswith(".docx"):
            doc = docx.Document(io.BytesIO(raw_bytes))
            text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
            gemini_parts.append(f"Word Document '{fname}':\n{text}")
            file_summaries.append({"name": fname, "type": "Word Document", "details": f"{len(doc.paragraphs)} paragraphs"})

        else:
            text = raw_bytes.decode("utf-8", errors="ignore")
            gemini_parts.append(f"Text Document '{fname}':\n{text}")
            file_summaries.append({"name": fname, "type": "Text Note", "details": f"{len(text)} chars"})

    return gemini_parts, file_summaries, visual_assets


# ==========================================
# 5. GEMINI AI ENGINE (AUDIT & BRIEF SYNTHESIS)
# ==========================================
def run_packaging_audit(api_key: str, model_name: str, gemini_parts: list, user_notes: str) -> PackagingAnalysisResult:
    client = genai.Client(api_key=api_key)

    system_prompt = """
    You are a Principal Packaging Technologist & Pre-Press Auditor at Esko.
    Your job is to inspect all uploaded packaging files (artwork dieline PDFs, spreadsheets, specs, and notes)
    and prepare verified data for a Design Agency Brief.

    PERFORM A RIGOROUS 3-PART AUDIT:
    1. EXTRACT BASELINE FACTS WITH EXACT CITATIONS:
       - Extract Brand, Product Name, Label ID, BPN, Net Weight, Serving Size, Servings per Container,
         Ingredients (with exact % values), Allergens, Claims/Certifications, Barcodes/UPCs, Storage Copy,
         Marketing Copy, and Visual Elements shown in photography/illustrations.
       - Cite the exact file name AND panel/row (e.g., 'DancingFish_Granola_Berry_US_Carton.pdf -> Back Panel'
         or 'Content_Sheet.xlsx -> Row 24').

    2. DETECT ALL INTERNAL & CROSS-FILE CONTRADICTIONS ('CONTRADICTION'):
       - Cross-check numbers between the PDF artwork and the Excel sheet (e.g., Oat % 75% vs 79%, Honey % 1% vs 9%,
         Barcode ending in 5 vs 6, missing Dried Strawberries/Pumpkin Seeds, duplicate Sunflower Seeds).
       - Check claims against ingredients (e.g., 'Certified Vegan' logo vs Honey in ingredients;
         'No Added Sugar' / '0g Added Sugars' vs Honey, Coconut Sugar, and Cane Sugar).
       - Check math & nutrition consistency:
         * Does Servings Per Container * Serving Size equal Net Weight? (e.g., 8 * 85g = 680g != 480g).
         * Do Back Panel nutrition lozenges match the Side Panel Nutrition Facts?
         * Did the Excel sheet mislabel the 85g serving column as '250g' and erroneously scale the 480g and 1000g columns?
         * Do Total Sugars + Dietary Fiber exceed Total Carbohydrate? Do fat/carb/protein macros match stated Calories?
       - Check Allergen formatting: Are main ingredients (Almonds, Pecans, Oats) mistakenly listed as 'May Contain'?
       - Check Artwork Photography vs Ingredients: Does the front/back bowl photo depict fruits (like fresh raspberries
         and blackberries) that are NOT in the ingredient list?
       - Check Regulatory & Typos: US Carton missing fluid/avoirdupois oz weight on Principal Display Panel (e.g. 480g / 16.9 oz),
         UK traffic light lozenges on a US carton, and spelling errors ('RECYCABLE', 'DIETRAY INFORMATION').

    3. IDENTIFY MISSING DESIGN BRIEF REQUIREMENTS ('MISSING_INFO'):
       - Ask about essential design brief fields not present in the files:
         * Target audience & positioning
         * Structural dieline dimensions / carton material finish (matte/gloss)
         * Whether the existing Dancing Fish logo & brand visual identity must remain unchanged or be refreshed
         * Whether this brief is correcting the existing Berry Crunch US carton or launching a new flavor variant.
    """

    contents = gemini_parts + [f"Additional Brand Manager Notes:\n{user_notes if user_notes else 'None provided.'}"]

    response = client.models.generate_content(
        model=model_name,
        contents=contents,
        config=types.GenerateContentConfig(
            system_instruction=system_prompt,
            response_mime_type="application/json",
            response_schema=PackagingAnalysisResult,
            temperature=0.1,
        ),
    )
    return PackagingAnalysisResult.model_validate_json(response.text)


def generate_final_brief(
    api_key: str,
    model_name: str,
    analysis: PackagingAnalysisResult,
    user_resolutions: dict,
    chat_history: list,
) -> FinalAgencyBrief:
    client = genai.Client(api_key=api_key)

    prompt = f"""
    You are a Senior Packaging Design Director creating a final, production-ready Design Agency Brief.
    Combine the extracted facts from the source files with the Brand Manager's verified resolutions below.

    --- EXTRACTED SOURCE FACTS ---
    {analysis.model_dump_json(indent=2)}

    --- BRAND MANAGER RESOLUTIONS TO CONFLICTS & QUESTIONS ---
    {json.dumps(user_resolutions, indent=2)}

    --- ADDITIONAL CLARIFYING CHAT ---
    {json.dumps(chat_history, indent=2)}

    RULES FOR THE FINAL BRIEF:
    1. Every single field MUST include an accurate `source_citation` showing where the final decision came from
       (e.g., 'Content Sheet Excel (Row 10) + Verified UPC Math', 'Carton PDF (Back Panel)', or 'Confirmed by Brand Manager').
    2. Where the Brand Manager resolved a conflict, use their chosen resolution and cite both the original source and
       'Confirmed by Brand Manager'.
    3. List all mandatory artwork corrections (typos like 'RECYCABLE' -> 'RECYCLABLE', 'DIETRAY' -> 'DIETARY',
       removing print guide text like 'Front panel Orientation' / 'Copy free', fixing US net weight to '480g (16.9 oz)',
       and updating fruit photography if requested).
    4. Any items the Brand Manager marked as 'Ask Agency / TBD' must be placed in `open_questions_for_agency`.
    """

    response = client.models.generate_content(
        model=model_name,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=FinalAgencyBrief,
            temperature=0.2,
        ),
    )
    return FinalAgencyBrief.model_validate_json(response.text)


# ==========================================
# 6. EXPORT GENERATORS (.DOCX, HTML/PDF, .MD)
# ==========================================
def brief_to_markdown(brief: FinalAgencyBrief) -> str:
    md = [f"# 📦 Packaging Design Brief: {brief.project_title}\n"]
    md.append(f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M')} | **Powered by:** Esko AI Brief Studio\n")
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
            val_clean = f.value.replace("\n", " ")
            md.append(f"| **{f.label}** | {val_clean} | `{f.source_citation}` |")
        md.append("")

    md.append("## 5. Mandatory Artwork & Pre-Press Corrections")
    for item in brief.typography_and_print_corrections:
        md.append(f"- [ ] {item}")

    md.append("\n## 6. Open Questions for the Design Agency")
    for q in brief.open_questions_for_agency:
        md.append(f"- ❓ {q}")

    return "\n".join(md)


def brief_to_docx_bytes(brief: FinalAgencyBrief) -> bytes:
    """Builds a formatted, corporate-styled Microsoft Word (.docx) document in memory."""
    doc = docx.Document()

    # Title Banner
    title_para = doc.add_heading(f"PACKAGING DESIGN BRIEF: {brief.project_title}", level=0)
    for run in title_para.runs:
        run.font.color.rgb = RGBColor(15, 23, 42)

    meta_para = doc.add_paragraph(
        f"Prepared via Esko AI Packaging Brief Studio  |  Date: {datetime.now().strftime('%B %d, %Y')}  |  Status: Approved for Agency Handover"
    )
    meta_para.runs[0].font.size = Pt(9.5)
    meta_para.runs[0].font.color.rgb = RGBColor(100, 116, 139)

    # Executive Summary
    doc.add_heading("Executive Summary", level=1)
    summary_p = doc.add_paragraph(brief.executive_summary)
    summary_p.paragraph_format.space_after = Pt(12)

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
        headers = ["Packaging Attribute", "Verified Agency Specification", "Source Traceability"]
        for i, h_text in enumerate(headers):
            hdr_cells[i].text = h_text
            for p in hdr_cells[i].paragraphs:
                for r in p.runs:
                    r.font.bold = True
                    r.font.size = Pt(10)
                    r.font.color.rgb = RGBColor(30, 58, 138)

        for item in fields:
            row_cells = table.add_row().cells
            row_cells[0].text = item.label
            if row_cells[0].paragraphs and row_cells[0].paragraphs[0].runs:
                row_cells[0].paragraphs[0].runs[0].font.bold = True
            row_cells[1].text = item.value
            row_cells[2].text = item.source_citation
            if row_cells[2].paragraphs and row_cells[2].paragraphs[0].runs:
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
    """Generates a print-ready HTML/PDF document with Esko branding and a 1-click Print-to-PDF button."""
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
            badge_bg = "#ECFDF5" if is_user else "#EEF2FF"
            badge_col = "#065F46" if is_user else "#3730A3"
            badge_bdr = "#A7F3D0" if is_user else "#C7D2FE"
            rows_html += f"""
            <tr>
                <td style="width: 24%; font-weight: 700; color: #0F172A;">{html.escape(f.label)}</td>
                <td style="width: 50%; color: #1E293B; white-space: pre-line;">{html.escape(f.value)}</td>
                <td style="width: 26%;">
                    <span style="display:inline-block; background:{badge_bg}; color:{badge_col}; border:1px solid {badge_bdr};
                                 padding:3px 10px; border-radius:999px; font-size:11px; font-weight:600;">
                        📌 {html.escape(f.source_citation)}
                    </span>
                </td>
            </tr>
            """
        sections_html += f"""
        <div class="section-block">
            <h2>{html.escape(sec_title)}</h2>
            <table>
                <thead>
                    <tr>
                        <th>Attribute</th>
                        <th>Verified Agency Specification</th>
                        <th>Source Traceability</th>
                    </tr>
                </thead>
                <tbody>{rows_html}</tbody>
            </table>
        </div>
        """

    corr_items = "".join([f"<li><input type='checkbox'> {html.escape(c)}</li>" for c in brief.typography_and_print_corrections])
    open_q_items = "".join([f"<li><strong>❓</strong> {html.escape(q)}</li>" for q in brief.open_questions_for_agency])

    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>{html.escape(brief.project_title)} - Esko Packaging Design Brief</title>
<style>
    @page {{ size: A4; margin: 16mm; }}
    body {{
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
        color: #1E293B;
        background: #F8FAFC;
        margin: 0;
        padding: 28px;
        line-height: 1.5;
    }}
    .page-container {{
        max-width: 960px;
        margin: 0 auto;
        background: #FFFFFF;
        padding: 36px 44px;
        border-radius: 14px;
        box-shadow: 0 10px 30px rgba(15, 23, 42, 0.08);
        border: 1px solid #E2E8F0;
    }}
    .top-bar {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 3px solid #1E3A8A;
        padding-bottom: 18px;
        margin-bottom: 24px;
    }}
    .brand-tag {{
        font-size: 12px;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: #2563EB;
    }}
    h1 {{
        font-size: 24px;
        font-weight: 800;
        color: #0F172A;
        margin: 4px 0 0 0;
    }}
    .exec-box {{
        background: #F1F5F9;
        border-left: 4px solid #2563EB;
        padding: 14px 18px;
        border-radius: 6px;
        margin-bottom: 24px;
        font-size: 14px;
    }}
    h2 {{
        font-size: 16px;
        font-weight: 700;
        color: #1E3A8A;
        margin: 24px 0 10px 0;
        border-bottom: 1px solid #E2E8F0;
        padding-bottom: 6px;
    }}
    table {{
        width: 100%;
        border-collapse: collapse;
        font-size: 13px;
        margin-bottom: 14px;
    }}
    th {{
        background: #F8FAFC;
        color: #475569;
        text-align: left;
        font-weight: 700;
        padding: 10px 12px;
        border: 1px solid #CBD5E1;
        text-transform: uppercase;
        font-size: 11px;
    }}
    td {{
        padding: 10px 12px;
        border: 1px solid #E2E8F0;
        vertical-align: top;
    }}
    ul.checklist {{
        list-style: none;
        padding-left: 0;
    }}
    ul.checklist li {{
        background: #FFFBEB;
        border: 1px solid #FDE68A;
        padding: 8px 12px;
        border-radius: 6px;
        margin-bottom: 6px;
        font-size: 13px;
    }}
    ul.questions {{
        list-style: none;
        padding-left: 0;
    }}
    ul.questions li {{
        background: #EEF2FF;
        border: 1px solid #C7D2FE;
        padding: 8px 12px;
        border-radius: 6px;
        margin-bottom: 6px;
        font-size: 13px;
    }}
    .print-btn {{
        background: #2563EB;
        color: white;
        border: none;
        padding: 10px 18px;
        border-radius: 8px;
        font-weight: 700;
        cursor: pointer;
        font-size: 13px;
    }}
    .print-btn:hover {{ background: #1D4ED8; }}
    @media print {{
        body {{ background: white; padding: 0; }}
        .page-container {{ box-shadow: none; border: none; padding: 0; }}
        .no-print {{ display: none !important; }}
    }}
</style>
</head>
<body>
<div class="page-container">
    <div class="top-bar">
        <div>
            <div class="brand-tag">PACKAGING POWERED BY ESKO • AI BRIEF STUDIO</div>
            <h1>{html.escape(brief.project_title)}</h1>
            <div style="font-size:12px; color:#64748B; margin-top:4px;">
                Generated: {datetime.now().strftime('%B %d, %Y')} • Full Source Traceability Enabled
            </div>
        </div>
        <div class="no-print">
            <button class="print-btn" onclick="window.print()">🖨️ Print / Save as PDF</button>
        </div>
    </div>

    <div class="exec-box">
        <strong>Executive Summary:</strong> {html.escape(brief.executive_summary)}
    </div>

    {sections_html}

    <div class="section-block">
        <h2>5. Mandatory Pre-Press & Artwork Corrections</h2>
        <ul class="checklist">{corr_items}</ul>
    </div>

    <div class="section-block">
        <h2>6. Open Questions for the Design Agency</h2>
        <ul class="questions">{open_q_items}</ul>
    </div>
</div>
</body>
</html>"""


# ==========================================
# 7. MAIN STREAMLIT WORKSPACE
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

    # ------------------------------------------
    # SIDEBAR: CONFIGURATION & ARCHITECTURE
    # ------------------------------------------
    with st.sidebar:
        st.markdown("## 📦 Esko Brief Studio")
        st.caption("AI-Powered Packaging Pre-Press & Design Brief Engine")

        api_key = st.text_input("🔑 Gemini API Key", type="password", help="Enter your Google Gemini API key")
        model_name = st.selectbox(
            "🤖 AI Reasoning Model",
            options=["gemini-2.5-flash", "gemini-3.8-pro", "gemini-2.5-flash", "gemini-flash-latest"],
            index=0,
            help="Defaults to gemini-2.5-flash",
        )

        st.divider()
        st.markdown("### 🏆 Challenge Capabilities")
        st.markdown(
            "- **Multi-Modal Dieline Vision:** Reads rotated carton panels & Excel sheets.\n"
            "- **Deterministic + AI Audit:** Verifies UPC-A check digits, serving math & vegan/sugar claims.\n"
            "- **1-Click Conflict Cards:** Resolves cross-file contradictions fast.\n"
            "- **Full Source Traceability:** Every spec links to its exact panel or row.\n"
            "- **Agency Exports:** Word (`.docx`), Print-Ready PDF/HTML & Markdown."
        )

        if st.session_state.file_summaries:
            st.divider()
            st.markdown("### 📂 Active Project Files")
            for fs in st.session_state.file_summaries:
                st.caption(f"✅ **{fs['name']}** ({fs['details']})")

    # ------------------------------------------
    # TOP HERO BANNER
    # ------------------------------------------
    st.markdown(
        """
        <div class="hero-banner">
            <div class="hero-pill">✨ Esko AI Packaging Challenge • End-to-End Brief Studio</div>
            <div class="hero-title">From Messy Product Files to Agency-Ready Packaging Brief</div>
            <div class="hero-subtitle">
                Ingest artwork dielines, spreadsheets, and notes • Automatically audit regulatory & cross-file contradictions • Export a traceable brief in Word, PDF/HTML, or Markdown.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ------------------------------------------
    # EXECUTIVE KPI DASHBOARD BAR
    # ------------------------------------------
    if st.session_state.analysis:
        res: PackagingAnalysisResult = st.session_state.analysis
        contradictions = [i for i in res.issues_and_questions if i.category == "CONTRADICTION"]
        missing_items = [i for i in res.issues_and_questions if i.category != "CONTRADICTION"]
        total_q = len(res.issues_and_questions)
        answered_q = len([v for v in st.session_state.user_resolutions.values() if str(v).strip()])
        readiness_pct = int((answered_q / total_q) * 100) if total_q > 0 else 100

        k1, k2, k3, k4, k5 = st.columns(5)
        with k1:
            st.markdown(
                f"<div class='kpi-card'><div class='kpi-value'>{len(st.session_state.file_summaries)}</div>"
                f"<div class='kpi-label'>Files Ingested</div></div>",
                unsafe_allow_html=True,
            )
        with k2:
            st.markdown(
                f"<div class='kpi-card'><div class='kpi-value' style='color:#2563EB;'>{len(res.extracted_facts)}</div>"
                f"<div class='kpi-label'>Traced Facts</div></div>",
                unsafe_allow_html=True,
            )
        with k3:
            st.markdown(
                f"<div class='kpi-card'><div class='kpi-value' style='color:#DC2626;'>{len(contradictions)}</div>"
                f"<div class='kpi-label'>Contradictions Flagged</div></div>",
                unsafe_allow_html=True,
            )
        with k4:
            st.markdown(
                f"<div class='kpi-card'><div class='kpi-value' style='color:#D97706;'>{len(missing_items)}</div>"
                f"<div class='kpi-label'>Missing Brief Inputs</div></div>",
                unsafe_allow_html=True,
            )
        with k5:
            st.markdown(
                f"<div class='kpi-card'><div class='kpi-value' style='color:#059669;'>{readiness_pct}%</div>"
                f"<div class='kpi-label'>Brief Readiness</div></div>",
                unsafe_allow_html=True,
            )
        st.write("")

    # ------------------------------------------
    # 4 WORKFLOW TABS
    # ------------------------------------------
    tab1, tab2, tab3, tab4 = st.tabs([
        "🖼️ 1. Upload & Visual Artwork Inspector",
        "🔍 2. Audit & Smart Questions",
        "💬 3. Copilot Refiner",
        "📋 4. Live Editable Brief & Exports (.DOCX / PDF)",
    ])

    # ==========================================
    # TAB 1: UPLOAD & VISUAL ARTWORK INSPECTOR
    # ==========================================
    with tab1:
        col_up, col_notes = st.columns([3, 2])
        with col_up:
            uploaded_files = st.file_uploader(
                "Upload Packaging Dieline PDF, Excel Content Sheet, Word Spec, or Images",
                type=["pdf", "xlsx", "xls", "docx", "png", "jpg", "jpeg", "txt"],
                accept_multiple_files=True,
            )
        with col_notes:
            user_notes = st.text_area(
                "Initial Brand Manager Launch Notes (Optional)",
                placeholder="e.g., Prepare updated US Berry Crunch Carton brief for design agency. Ensure FDA compliance and resolve spec discrepancies...",
                height=125,
            )

        if st.button("🚀 Analyze Packaging Files & Launch Visual Inspector", type="primary", use_container_width=True):
            if not api_key:
                st.error("Please enter your Gemini API Key in the left sidebar first.")
            elif not uploaded_files:
                st.warning("Please upload at least one file (e.g., Carton PDF and/or Excel Content Sheet).")
            else:
                with st.spinner("Rendering high-res dieline artwork, parsing spreadsheets, and running pre-press audit..."):
                    gemini_parts, summaries, visuals = parse_uploaded_files(uploaded_files)
                    st.session_state.file_summaries = summaries
                    st.session_state.visual_assets = visuals
                    st.session_state.analysis = run_packaging_audit(api_key, model_name, gemini_parts, user_notes)
                    st.session_state.final_brief = None
                st.success("Analysis complete! Inspect your artwork vs. extracted facts below, then proceed to Tab 2.")
                st.rerun()

        if st.session_state.analysis:
            res: PackagingAnalysisResult = st.session_state.analysis
            st.divider()
            st.markdown("#### 📌 Executive Product Summary")
            st.info(res.product_summary)

            # SIDE-BY-SIDE VISUAL ARTWORK INSPECTOR + TRACED DATA TABLE
            insp_left, insp_right = st.columns([1.15, 1.35], gap="large")

            with insp_left:
                st.markdown("#### 🎨 Visual Dieline & Artwork Inspector")
                if st.session_state.visual_assets:
                    asset_titles = [a["title"] for a in st.session_state.visual_assets]
                    selected_title = st.selectbox("Select Artwork / Dieline Page", options=asset_titles)
                    selected_asset = next(a for a in st.session_state.visual_assets if a["title"] == selected_title)

                    rot_col1, rot_col2 = st.columns([2, 1])
                    with rot_col1:
                        rotation = st.radio(
                            "🔄 Rotate Dieline Orientation (to read flaps & side panels):",
                            options=[0, 90, 180, 270],
                            format_func=lambda x: {
                                0: "0° (Front/Back Normal)",
                                90: "90° CW (Side Panel)",
                                180: "180° (Upside-Down Top Flap)",
                                270: "270° CCW (Side Panel)",
                            }[x],
                            horizontal=True,
                        )
                    with rot_col2:
                        zoom_full = st.checkbox("🔍 Expand Full Width", value=True)

                    img_to_show = selected_asset["image"]
                    if rotation != 0:
                        img_to_show = img_to_show.rotate(-rotation, expand=True)

                    st.image(
                        img_to_show,
                        caption=f"{selected_title} (Orientation: {rotation}°)",
                        use_container_width=zoom_full,
                    )
                else:
                    st.info("Upload a PDF dieline or image file to inspect visual artwork panels here.")

            with insp_right:
                st.markdown(f"#### 📑 Extracted Facts with Source Traceability ({len(res.extracted_facts)})")
                search_q = st.text_input(
                    "🔎 Filter Extracted Facts by Keyword or Source (e.g., 'Honey', 'Barcode', 'Back Panel')",
                    placeholder="Type to filter...",
                )

                facts_rows = []
                for f in res.extracted_facts:
                    combined_str = f"{f.field_name} {f.value} {f.source_file} {f.source_location}".lower()
                    if not search_q or search_q.lower() in combined_str:
                        facts_rows.append({
                            "Attribute": f.field_name,
                            "Extracted Value": f.value,
                            "Source File": f.source_file,
                            "Panel / Row": f.source_location,
                            "Confidence": f.confidence,
                        })

                st.dataframe(pd.DataFrame(facts_rows), use_container_width=True, hide_index=True, height=380)

                # Automated Deterministic Barcode Check Card
                barcodes = [
                    f for f in res.extracted_facts
                    if "barcode" in f.field_name.lower() or "upc" in f.field_name.lower()
                ]
                if barcodes:
                    st.markdown("##### 🔢 Automated Pre-Press UPC-A Modulo-10 Validator")
                    for b in barcodes:
                        check = verify_upc_a(b.value)
                        if check["valid"]:
                            st.success(f"**{b.value}** (`{b.source_file} → {b.source_location}`): {check['msg']}")
                        else:
                            st.error(f"**{b.value}** (`{b.source_file} → {b.source_location}`): {check['msg']}")

    # ==========================================
    # TAB 2: AUDIT & SMART QUESTIONS
    # ==========================================
    with tab2:
        if not st.session_state.analysis:
            st.info("Upload and analyze your files in Tab 1 first.")
        else:
            res: PackagingAnalysisResult = st.session_state.analysis
            contradictions = [i for i in res.issues_and_questions if i.category == "CONTRADICTION"]
            missing_items = [i for i in res.issues_and_questions if i.category != "CONTRADICTION"]

            total_q = len(res.issues_and_questions)
            answered_q = len([v for v in st.session_state.user_resolutions.values() if str(v).strip()])
            st.progress(
                answered_q / total_q if total_q else 1.0,
                text=f"Resolution Progress: {answered_q} of {total_q} items reviewed by Brand Manager",
            )

            col_conf, col_gap = st.columns(2, gap="large")

            with col_conf:
                st.subheader(f"🔴 Contradictions & Pre-Press Errors ({len(contradictions)})")
                st.caption("Cross-checked across PDF artwork panels, Excel specs, and FDA/packaging rules.")

                for issue in contradictions:
                    st.markdown(
                        f"""<div class="conflict-card">
                            <span class="severity-pill-critical">{issue.severity}</span>
                            <strong>{issue.title}</strong><br>
                            <div style="margin-top:6px; font-size:0.92rem;">{issue.explanation}</div>
                            <div style="margin-top:8px; font-size:0.8rem; color:#475569;">
                                📌 <strong>Traced Sources:</strong> {', '.join(issue.sources_involved)}
                            </div>
                        </div>""",
                        unsafe_allow_html=True,
                    )
                    options = issue.suggested_options + [
                        "Custom resolution...",
                        "Ask Design Agency / Leave as Open Question",
                    ]
                    choice = st.radio(
                        f"❓ **{issue.question_for_user}**",
                        options=options,
                        key=f"radio_{issue.id}",
                    )
                    if choice == "Custom resolution...":
                        custom_val = st.text_input("Enter your custom decision:", key=f"custom_{issue.id}")
                        st.session_state.user_resolutions[issue.title] = custom_val
                    else:
                        st.session_state.user_resolutions[issue.title] = choice
                    st.divider()

            with col_gap:
                st.subheader(f"🟡 Missing Design Brief Specs ({len(missing_items)})")
                st.caption("Smart follow-up questions for essential agency brief inputs missing from your files.")

                for gap in missing_items:
                    st.markdown(
                        f"""<div class="missing-card">
                            <span class="severity-pill-gap">MISSING INPUT</span>
                            <strong>{gap.title}</strong><br>
                            <div style="margin-top:6px; font-size:0.92rem;">{gap.explanation}</div>
                        </div>""",
                        unsafe_allow_html=True,
                    )
                    options = gap.suggested_options + [
                        "Custom specification...",
                        "Ask Design Agency / Leave as Open Question",
                    ]
                    choice = st.radio(
                        f"❓ **{gap.question_for_user}**",
                        options=options,
                        key=f"radio_{gap.id}",
                    )
                    if choice == "Custom specification...":
                        custom_val = st.text_input("Enter your custom specification:", key=f"custom_{gap.id}")
                        st.session_state.user_resolutions[gap.title] = custom_val
                    else:
                        st.session_state.user_resolutions[gap.title] = choice
                    st.divider()

            if st.button("✨ Synthesize Final Traceable Design Agency Brief", type="primary", use_container_width=True):
                with st.spinner("Synthesizing verified facts, source citations, and your decisions into the final brief..."):
                    st.session_state.final_brief = generate_final_brief(
                        api_key,
                        model_name,
                        st.session_state.analysis,
                        st.session_state.user_resolutions,
                        st.session_state.chat_history,
                    )
                st.success("Final Brief ready! Open Tab 4 (`📋 Live Editable Brief & Exports`) to review and download.")

    # ==========================================
    # TAB 3: COPILOT REFINER (CONVERSATIONAL)
    # ==========================================
    with tab3:
        st.subheader("💬 Packaging Brief Copilot")
        st.caption(
            "Ask follow-up questions about your files or give natural-language instructions (e.g., "
            "'Replace honey with organic agave syrup so we can keep the Certified Vegan logo')."
        )

        for msg in st.session_state.chat_history:
            with st.chat_message(msg["role"]):
                st.write(msg["content"])

        user_msg = st.chat_input("Ask a question or add a custom packaging brief instruction...")
        if user_msg:
            st.session_state.chat_history.append({"role": "user", "content": user_msg})
            with st.chat_message("user"):
                st.write(user_msg)

            if api_key and st.session_state.analysis:
                client = genai.Client(api_key=api_key)
                reply = client.models.generate_content(
                    model=model_name,
                    contents=(
                        f"Extracted Data: {st.session_state.analysis.model_dump_json()}\n"
                        f"Current Resolutions: {json.dumps(st.session_state.user_resolutions)}\n"
                        f"User message: {user_msg}\n"
                        "Acknowledge how this updates the brief or answer the question concisely with exact source citations."
                    ),
                )
                bot_text = reply.text
            else:
                bot_text = "Recorded! This instruction will be included when you generate the final brief."

            st.session_state.chat_history.append({"role": "assistant", "content": bot_text})
            with st.chat_message("assistant"):
                st.write(bot_text)

    # ==========================================
    # TAB 4: LIVE EDITABLE BRIEF & MULTI-FORMAT EXPORTS
    # ==========================================
    with tab4:
        if not st.session_state.final_brief:
            st.info("Complete Tab 2 and click **'✨ Synthesize Final Traceable Design Agency Brief'** to unlock the live editor and exports.")
        else:
            brief: FinalAgencyBrief = st.session_state.final_brief

            # MULTI-FORMAT EXPORT BAR (.DOCX, PRINT-READY HTML/PDF, .MD, .JSON)
            st.markdown("#### 📤 Export Agency-Ready Brief")
            exp1, exp2, exp3, exp4 = st.columns(4)

            docx_bytes = brief_to_docx_bytes(brief)
            html_report = brief_to_styled_html(brief)
            md_report = brief_to_markdown(brief)
            json_report = brief.model_dump_json(indent=2)

            with exp1:
                st.download_button(
                    "📄 Download Word (.docx)",
                    data=docx_bytes,
                    file_name="Esko_Packaging_Design_Brief.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    type="primary",
                )
            with exp2:
                st.download_button(
                    "🖨️ Download Print-Ready PDF/HTML",
                    data=html_report,
                    file_name="Esko_Packaging_Design_Brief_Print.html",
                    mime="text/html",
                    use_container_width=True,
                )
            with exp3:
                st.download_button(
                    "📝 Download Markdown (.md)",
                    data=md_report,
                    file_name="Esko_Packaging_Design_Brief.md",
                    mime="text/markdown",
                    use_container_width=True,
                )
            with exp4:
                st.download_button(
                    "🗂️ Download Audit Trail (.json)",
                    data=json_report,
                    file_name="Esko_Brief_Traceability.json",
                    mime="application/json",
                    use_container_width=True,
                )

            st.divider()

            view_mode = st.radio(
                "Select View Mode:",
                options=["✏️ Interactive Live Brief Editor", "👁️ Styled Agency Print Preview (PDF Layout)"],
                horizontal=True,
            )

            if view_mode == "👁️ Styled Agency Print Preview (PDF Layout)":
                st.caption("Tip: Click the blue **'🖨️ Print / Save as PDF'** button inside the preview below to save directly as a PDF.")
                st.components.v1.html(html_report, height=900, scrolling=True)
            else:
                brief.project_title = st.text_input("Project Title", value=brief.project_title)
                brief.executive_summary = st.text_area("Executive Summary", value=brief.executive_summary, height=85)

                def render_editable_section(section_title: str, fields_list: List[BriefField], sec_key: str):
                    st.subheader(section_title)
                    for idx, item in enumerate(fields_list):
                        c1, c2 = st.columns([3.2, 1.2])
                        with c1:
                            item.value = st.text_input(
                                f"**{item.label}**",
                                value=item.value,
                                key=f"{sec_key}_{idx}",
                            )
                        with c2:
                            badge_class = "user-badge" if "Brand Manager" in item.source_citation else "source-badge"
                            st.markdown(
                                f"<div style='margin-top: 28px;'><span class='{badge_class}'>📌 {html.escape(item.source_citation)}</span></div>",
                                unsafe_allow_html=True,
                            )

                render_editable_section("1. Project & Brand Overview", brief.project_and_brand_overview, "sec1")
                st.divider()
                render_editable_section("2. Structural & Pack Specifications", brief.structural_and_pack_specs, "sec2")
                st.divider()
                render_editable_section("3. Mandatory Copy & Regulatory Content", brief.mandatory_copy_and_regulatory, "sec3")
                st.divider()
                render_editable_section("4. Visual & Creative Direction", brief.visual_and_creative_direction, "sec4")
                st.divider()

                st.subheader("5. Mandatory Pre-Press & Artwork Corrections")
                corrections_text = st.text_area(
                    "Edit Pre-Press Checklist (one item per line)",
                    value="\n".join(brief.typography_and_print_corrections),
                    height=150,
                )
                brief.typography_and_print_corrections = [
                    line.strip() for line in corrections_text.split("\n") if line.strip()
                ]

                st.subheader("6. Open Questions for the Design Agency")
                questions_text = st.text_area(
                    "Edit Open Agency Questions (one item per line)",
                    value="\n".join(brief.open_questions_for_agency),
                    height=120,
                )
                brief.open_questions_for_agency = [
                    line.strip() for line in questions_text.split("\n") if line.strip()
                ]


if __name__ == "__main__":
    main()