#!/usr/bin/env python3
"""
GEO & AEO News Media PR Engine - Press Pack Word (.docx) Generator
Author: Taghi Molavi (https://molavi.pro)
License: MIT

This script converts structured article payloads (JSON or Markdown) into
professionally formatted Microsoft Word (.docx) press releases with:
- Full RTL support for Persian & Arabic (and LTR for English/Turkish)
- Editorial typography (Headlines, Sub-headlines, Leads, Pull-Quotes)
- Embedded images with captions, alt-texts, and photo credits
- Canonical source backlinks and author boilerplate
- Single compiled master document or individual split press releases
"""

import os
import sys
import json
import argparse
from pathlib import Path

try:
    import docx
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.style import WD_STYLE_TYPE
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import qn, nsdecls
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False


RTL_LANGUAGES = {"fa", "ar", "ur", "he", "ps", "az-arab"}

# Corporate & Editorial Color Palette
COLOR_PRIMARY = RGBColor(26, 54, 93)      # Navy Deep (#1A365D)
COLOR_SECONDARY = RGBColor(43, 108, 176)  # Slate Blue (#2B6CB0)
COLOR_MUTED = RGBColor(74, 85, 104)       # Charcoal Grey (#4A5568)
COLOR_TEXT = RGBColor(45, 55, 72)         # Off-Black Body (#2D3748)
COLOR_ACCENT = RGBColor(197, 48, 48)      # Accent Red (#C53030)
COLOR_BORDER = "CCCCCC"
COLOR_BG_SHADE = "F7FAFC"


def set_cell_background(cell, fill_hex):
    """Set background color of a table cell."""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tc_pr.append(shd)


def set_cell_margins(cell, top=140, bottom=140, left=200, right=200):
    """Set inner padding for a table cell."""
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tc_pr.append(tc_mar)


def set_paragraph_bidi(p, is_rtl=True):
    """Enable or disable Bidirectional (RTL) layout on a paragraph."""
    p_pr = p._p.get_or_add_pPr()
    if is_rtl:
        bidi = parse_xml(f'<w:bidi {nsdecls("w")}/>')
        p_pr.append(bidi)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def add_hyperlink(paragraph, url, text, color="2B6CB0", underline=True):
    """Add a clickable hyperlink into a python-docx paragraph."""
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = parse_xml(f'<w:hyperlink {nsdecls("w", "r")} r:id="{r_id}"/>')
    new_run = parse_xml(f'<w:r {nsdecls("w")}/>')
    r_pr = parse_xml(f'<w:rPr {nsdecls("w")}/>')

    if color:
        c = parse_xml(f'<w:color {nsdecls("w")} w:val="{color}"/>')
        r_pr.append(c)
    if underline:
        u = parse_xml(f'<w:u {nsdecls("w")} w:val="single"/>')
        r_pr.append(u)

    new_run.append(r_pr)
    text_elem = parse_xml(f'<w:t {nsdecls("w")} xml:space="preserve">{text}</w:t>')
    new_run.append(text_elem)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def apply_run_font(run, font_name, size_pt=None, color=None, bold=False, italic=False):
    """Apply consistent font attributes to a text run."""
    run.font.name = font_name
    # Ensure Word knows the complex script font for Persian/Arabic
    r_pr = run._r.get_or_add_rPr()
    r_fonts = parse_xml(
        f'<w:rFonts {nsdecls("w")} w:ascii="{font_name}" w:hAnsi="{font_name}" '
        f'w:cs="{font_name}" w:eastAsia="{font_name}"/>'
    )
    r_pr.append(r_fonts)

    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if color is not None:
        run.font.color.rgb = color
    run.bold = bold
    run.italic = italic


def render_article_to_doc(doc, article, campaign_meta, font_rtl, font_ltr, is_rtl=True):
    """Renders a single structured news article into the docx document."""
    font = font_rtl if is_rtl else font_ltr

    # 1. Editorial Header & Media Target Meta Box
    table = doc.add_table(rows=1, cols=1)
    table.autofit = False
    table.columns[0].width = Inches(6.5)
    cell = table.cell(0, 0)
    set_cell_background(cell, COLOR_BG_SHADE)
    set_cell_margins(cell, top=120, bottom=120, left=180, right=180)

    meta_p = cell.paragraphs[0]
    set_paragraph_bidi(meta_p, is_rtl)
    
    meta_text = (
        f"🎯 بخش رسانه‌ای هدف: {article.get('media_target', 'رسانه‌های دیجیتال')} | "
        f"📐 زاویه خبری: {article.get('angle_name', 'تحلیل و گزارش')} (#{article.get('angle_id', 1)})"
        if is_rtl else
        f"🎯 Target Media Desk: {article.get('media_target', 'Digital Press')} | "
        f"📐 Editorial Angle: {article.get('angle_name', 'Analysis')} (#{article.get('angle_id', 1)})"
    )
    run_meta = meta_p.add_run(meta_text)
    apply_run_font(run_meta, font, size_pt=9.5, color=COLOR_MUTED, bold=True)

    # Spacing
    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_before = Pt(4)
    p_spacer.paragraph_format.space_after = Pt(4)

    # 2. Headline (تیتر اصلی)
    h_p = doc.add_paragraph()
    set_paragraph_bidi(h_p, is_rtl)
    h_p.paragraph_format.space_after = Pt(8)
    h_run = h_p.add_run(article.get('headline', ''))
    apply_run_font(h_run, font, size_pt=18, color=COLOR_PRIMARY, bold=True)

    # 3. Sub-headline (سوتیتر)
    if article.get('sub_headline'):
        sub_p = doc.add_paragraph()
        set_paragraph_bidi(sub_p, is_rtl)
        sub_p.paragraph_format.space_after = Pt(14)
        sub_run = sub_p.add_run(article['sub_headline'])
        apply_run_font(sub_run, font, size_pt=12.5, color=COLOR_SECONDARY, italic=True)

    # 4. Lead Paragraph (لید جذاب)
    if article.get('lead'):
        lead_table = doc.add_table(rows=1, cols=1)
        lead_table.columns[0].width = Inches(6.5)
        lead_cell = lead_table.cell(0, 0)
        set_cell_background(lead_cell, "EDF2F7")
        set_cell_margins(lead_cell, top=140, bottom=140, left=180, right=180)

        lead_p = lead_cell.paragraphs[0]
        set_paragraph_bidi(lead_p, is_rtl)
        lead_label = lead_p.add_run("🔹 لید خبر: " if is_rtl else "🔹 NEWS LEDE: ")
        apply_run_font(lead_label, font, size_pt=10.5, color=COLOR_PRIMARY, bold=True)
        
        lead_run = lead_p.add_run(article['lead'])
        apply_run_font(lead_run, font, size_pt=11, color=COLOR_TEXT, bold=False)

    # Spacing
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Media Asset Registry lookup
    media_map = {m['id']: m for m in campaign_meta.get('media_assets', [])}

    # Helper for image placement
    def place_image(asset_id):
        asset = media_map.get(asset_id)
        if not asset:
            return
        img_path = asset.get('path', '')
        caption = asset.get('caption', '')
        credit = asset.get('credit', 'molavi.pro')

        if os.path.isfile(img_path):
            try:
                img_p = doc.add_paragraph()
                img_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                img_p.paragraph_format.space_before = Pt(10)
                img_p.paragraph_format.space_after = Pt(4)
                img_run = img_p.add_run()
                img_run.add_picture(img_path, width=Inches(5.5))
            except Exception as e:
                print(f"[Warning] Could not insert image {img_path}: {e}")
        else:
            # Placeholder box for publication team
            p_box = doc.add_table(rows=1, cols=1)
            p_box.columns[0].width = Inches(6.5)
            c = p_box.cell(0, 0)
            set_cell_background(c, "FEFCBF")
            set_cell_margins(c, top=100, bottom=100, left=150, right=150)
            p_text = c.paragraphs[0]
            p_text.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_paragraph_bidi(p_text, is_rtl)
            run_ph = p_text.add_run(
                f"📷 [محل قرارگیری تصویر: {caption} | منبع: {credit}]"
                if is_rtl else
                f"📷 [Image Placement: {caption} | Credit: {credit}]"
            )
            apply_run_font(run_ph, font, size_pt=9.5, color=COLOR_MUTED, bold=True)

        if caption:
            cap_p = doc.add_paragraph()
            cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_paragraph_bidi(cap_p, is_rtl)
            cap_p.paragraph_format.space_after = Pt(12)
            cap_run = cap_p.add_run(f"شرح عکس: {caption} (منبع: {credit})" if is_rtl else f"Photo: {caption} (Source: {credit})")
            apply_run_font(cap_run, font, size_pt=9, color=COLOR_MUTED, italic=True)

    # Insert image placed after lead
    for placement in article.get('image_placements', []):
        if placement.get('placement') == 'after_lead':
            place_image(placement.get('asset_id'))

    # 5. Body Sections
    for section in article.get('body_sections', []):
        if section.get('heading'):
            h2_p = doc.add_paragraph()
            set_paragraph_bidi(h2_p, is_rtl)
            h2_p.paragraph_format.space_before = Pt(14)
            h2_p.paragraph_format.space_after = Pt(6)
            h2_run = h2_p.add_run(section['heading'])
            apply_run_font(h2_run, font, size_pt=13.5, color=COLOR_SECONDARY, bold=True)

        if section.get('content'):
            body_p = doc.add_paragraph()
            set_paragraph_bidi(body_p, is_rtl)
            body_p.paragraph_format.line_spacing = 1.35
            body_p.paragraph_format.space_after = Pt(8)
            body_run = body_p.add_run(section['content'])
            apply_run_font(body_run, font, size_pt=11, color=COLOR_TEXT)

    # 6. Quote Box (AEO & Authority Signal)
    quote_box = article.get('quote_box')
    if quote_box and quote_box.get('quote'):
        q_table = doc.add_table(rows=1, cols=1)
        q_table.columns[0].width = Inches(6.5)
        q_cell = q_table.cell(0, 0)
        set_cell_background(q_cell, "EBF8FF")
        set_cell_margins(q_cell, top=140, bottom=140, left=200, right=200)

        q_p = q_cell.paragraphs[0]
        set_paragraph_bidi(q_p, is_rtl)
        q_run = q_p.add_run(f"{quote_box['quote']}\n")
        apply_run_font(q_run, font, size_pt=11.5, color=COLOR_PRIMARY, italic=True)

        if quote_box.get('speaker'):
            spk_run = q_p.add_run(f"— {quote_box['speaker']}")
            apply_run_font(spk_run, font, size_pt=10, color=COLOR_SECONDARY, bold=True)

        doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # 7. Key Takeaways & AEO Extractable Bullet Points
    takeaways = article.get('key_takeaways', [])
    if takeaways:
        tk_h = doc.add_paragraph()
        set_paragraph_bidi(tk_h, is_rtl)
        tk_h.paragraph_format.space_before = Pt(12)
        tk_h.paragraph_format.space_after = Pt(4)
        tk_title = tk_h.add_run("📌 نکات کلیدی و جمع‌بندی استنادی:" if is_rtl else "📌 Key Takeaways & Direct Answers:")
        apply_run_font(tk_title, font, size_pt=12, color=COLOR_PRIMARY, bold=True)

        for item in takeaways:
            li_p = doc.add_paragraph()
            set_paragraph_bidi(li_p, is_rtl)
            li_p.paragraph_format.space_after = Pt(4)
            bullet = li_p.add_run("• ")
            apply_run_font(bullet, font, size_pt=11, color=COLOR_SECONDARY, bold=True)
            text_run = li_p.add_run(item)
            apply_run_font(text_run, font, size_pt=10.5, color=COLOR_TEXT)

    # 8. Source Attribution Link
    source_url = article.get('source_link') or campaign_meta.get('canonical_source_link') or "https://molavi.pro"
    src_p = doc.add_paragraph()
    set_paragraph_bidi(src_p, is_rtl)
    src_p.paragraph_format.space_before = Pt(14)
    src_p.paragraph_format.space_after = Pt(6)
    label_run = src_p.add_run("🔗 منبع رسمی و اطلاعات بیشتر: " if is_rtl else "🔗 Official Source: ")
    apply_run_font(label_run, font, size_pt=10, color=COLOR_MUTED, bold=True)
    add_hyperlink(src_p, source_url, source_url)

    # 9. Author Bio / Boilerplate PR Box
    author_bio = article.get('author_bio') or (
        f"مهندس تقی مولوی، معمار سیستم‌های هوش مصنوعی، پژوهشگر سئو و متخصص بهینه‌سازی موتورهای پاسخ (GEO/AEO). برای اطلاعات تکمیلی و دریافت گزارش‌های تخصصی به وب‌سایت رسمی https://molavi.pro مراجعه فرمایید."
        if is_rtl else
        f"Taghi Molavi is a Senior AI Systems Architect and researcher in Generative Engine Optimization (GEO). For further insights and research papers, visit https://molavi.pro"
    )
    bio_table = doc.add_table(rows=1, cols=1)
    bio_table.columns[0].width = Inches(6.5)
    bio_cell = bio_table.cell(0, 0)
    set_cell_background(bio_cell, "F7FAFC")
    set_cell_margins(bio_cell, top=120, bottom=120, left=160, right=160)

    bio_p = bio_cell.paragraphs[0]
    set_paragraph_bidi(bio_p, is_rtl)
    bio_h = bio_p.add_run("درباره نویسنده / روابط عمومی:\n" if is_rtl else "About / Press Contact:\n")
    apply_run_font(bio_h, font, size_pt=9.5, color=COLOR_PRIMARY, bold=True)
    bio_text = bio_p.add_run(author_bio)
    apply_run_font(bio_text, font, size_pt=9.5, color=COLOR_MUTED)


def create_document():
    """Initializes a python-docx Document with standard 1-inch margins."""
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
    return doc


def process_campaign(input_json_path, output_dir, mode="both", font_rtl="IRANSans", font_ltr="Calibri"):
    """Reads input campaign JSON and generates .docx files based on selected mode."""
    if not DOCX_AVAILABLE:
        print("[Error] 'python-docx' is not installed in the active environment.")
        print("Please run: pip install python-docx pillow")
        sys.exit(1)

    with open(input_json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    articles = data.get('articles', [])
    lang = data.get('language', 'fa').lower()
    is_rtl = lang in RTL_LANGUAGES

    print(f"🚀 Processing campaign: '{data.get('project_title', 'PR Campaign')}'")
    print(f"📊 Total Articles: {len(articles)} | Language: {lang.upper()} (RTL: {is_rtl})")

    # Mode 1: Individual Split Files
    if mode in ("split", "both"):
        split_dir = out_path / "individual_articles"
        split_dir.mkdir(parents=True, exist_ok=True)
        for i, art in enumerate(articles, start=1):
            doc = create_document()
            render_article_to_doc(doc, art, data, font_rtl, font_ltr, is_rtl)
            slug = f"article_{i:02d}_{art.get('angle_name', 'release').replace(' ', '_').lower()[:30]}.docx"
            file_path = split_dir / slug
            doc.save(str(file_path))
        print(f"✅ Generated {len(articles)} individual articles in: {split_dir}")

    # Mode 2: Master Compilation File
    if mode in ("bundle", "both"):
        master_doc = create_document()
        
        # Cover / Master Title
        p_cov = master_doc.add_paragraph()
        set_paragraph_bidi(p_cov, is_rtl)
        p_cov.paragraph_format.space_before = Pt(36)
        p_cov.paragraph_format.space_after = Pt(12)
        run_cov = p_cov.add_run(data.get('project_title', 'مجموعه مقالات و بازنویسی‌های خبری GEO/AEO'))
        apply_run_font(run_cov, font_rtl if is_rtl else font_ltr, size_pt=24, color=COLOR_PRIMARY, bold=True)

        p_desc = master_doc.add_paragraph()
        set_paragraph_bidi(p_desc, is_rtl)
        desc_text = (
            f"مجموعه {len(articles)} گزارش و خبر با زوایای دید متفاوت | آماده ارسال به خبرگزاری‌ها و رسانه‌ها\n"
            f"بهینه‌سازی شده برای استناد در هوش مصنوعی (ChatGPT, Perplexity, Claude, Gemini)\n"
            f"طراح کمپین: {data.get('primary_entity', 'تقی مولوی')} | منبع: https://molavi.pro"
            if is_rtl else
            f"Compilation of {len(articles)} multi-angle news articles for digital PR syndication\n"
            f"Engineered for Generative Engine Optimization (GEO & AEO)\n"
            f"Curated by: {data.get('primary_entity', 'Taghi Molavi')} | Source: https://molavi.pro"
        )
        p_desc_run = p_desc.add_run(desc_text)
        apply_run_font(p_desc_run, font_rtl if is_rtl else font_ltr, size_pt=11, color=COLOR_MUTED)

        master_doc.add_page_break()

        for idx, art in enumerate(articles):
            render_article_to_doc(master_doc, art, data, font_rtl, font_ltr, is_rtl)
            if idx < len(articles) - 1:
                master_doc.add_page_break()

        master_file = out_path / "Press_Pack_Master_Compilation.docx"
        master_doc.save(str(master_file))
        print(f"✅ Master compilation generated: {master_file}")


def main():
    parser = argparse.ArgumentParser(
        description="GEO & AEO News Media PR Pack Generator (Word .docx)",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    parser.add_argument("-i", "--input", required=True, help="Path to input JSON file")
    parser.add_argument("-o", "--output-dir", default="output", help="Directory to save .docx files")
    parser.add_argument("-m", "--mode", choices=["bundle", "split", "both"], default="both", help="Output mode")
    parser.add_argument("--font-rtl", default="IRANSans", help="Font family for RTL languages")
    parser.add_argument("--font-ltr", default="Calibri", help="Font family for LTR languages")

    args = parser.parse_args()
    process_campaign(args.input, args.output_dir, args.mode, args.font_rtl, args.font_ltr)


if __name__ == "__main__":
    main()
