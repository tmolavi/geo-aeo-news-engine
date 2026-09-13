#!/usr/bin/env python3
"""
GEO & AEO News Media PR Engine - Press Pack Word (.docx) Generator
Author: Taghi Molavi (https://molavi.pro)
License: MIT

Professional, 100% publication-ready Microsoft Word (.docx) press releases:
- Native, flawless UTF-8 & Right-to-Left (RTL) Complex Script typography (zero garbled text)
- Built-in automatic web image crawler & high-res downloader
- Real embedded photos with natural journalistic captions
- ZERO meta-instructions, prompt artifacts, or editorial brackets in the copy
- Ready for immediate publication by newsrooms, journalists, and media desks
"""

import os
import sys
import json
import io
import argparse
from pathlib import Path
from urllib.parse import urljoin

try:
    import docx
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import parse_xml
    from docx.oxml.ns import qn, nsdecls
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import requests
    from bs4 import BeautifulSoup
    from PIL import Image
    SCRAPER_AVAILABLE = True
except ImportError:
    SCRAPER_AVAILABLE = False

RTL_LANGUAGES = {"fa", "ar", "ur", "he", "ps", "az-arab"}

# Professional Journalistic Color Palette
COLOR_HEADLINE = RGBColor(15, 23, 42)    # Slate Black (#0F172A)
COLOR_SUBTITLE = RGBColor(71, 85, 105)   # Slate Grey (#475569)
COLOR_BODY = RGBColor(30, 41, 59)        # Rich Body Text (#1E293B)
COLOR_ACCENT = RGBColor(30, 58, 138)     # Classic Press Blue (#1E3A8A)
COLOR_CAPTION = RGBColor(100, 116, 139)  # Caption Slate (#64748B)


def crawl_website_images(url, output_dir="assets/crawled", min_width=200, min_height=150, max_images=10):
    """
    Crawls a target webpage, discovers high-res images (og:image, article images),
    converts them into clean JPEGs via Pillow, and saves them locally.
    """
    if not SCRAPER_AVAILABLE:
        print("[Notice] requests/beautifulsoup4/Pillow not available for live scraping.")
        return []

    os.makedirs(output_dir, exist_ok=True)
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }

    print(f"🔍 Crawling images from: {url}")
    try:
        resp = requests.get(url, headers=headers, timeout=12)
        if resp.status_code != 200:
            print(f"[Warning] Received HTTP {resp.status_code} from {url}")
            return []
        soup = BeautifulSoup(resp.text, "html.parser")
    except Exception as e:
        print(f"[Warning] Failed to fetch {url}: {e}")
        return []

    candidate_urls = []
    # 1. OpenGraph & Twitter Cards
    og = soup.find("meta", property="og:image") or soup.find("meta", attrs={"name": "og:image"})
    if og and og.get("content"):
        candidate_urls.append(og.get("content"))
    tw = soup.find("meta", attrs={"name": "twitter:image"})
    if tw and tw.get("content"):
        candidate_urls.append(tw.get("content"))

    # 2. Page Images
    for img in soup.find_all("img"):
        src = img.get("src") or img.get("data-src")
        if src and not src.startswith("data:") and not any(ext in src.lower() for ext in [".svg", ".ico", "pixel"]):
            candidate_urls.append(src)

    saved_files = []
    seen = set()

    for idx, raw_url in enumerate(candidate_urls):
        if len(saved_files) >= max_images:
            break
        full_url = urljoin(url, raw_url)
        if full_url in seen:
            continue
        seen.add(full_url)

        try:
            r = requests.get(full_url, headers=headers, timeout=8)
            if r.status_code == 200 and len(r.content) > 5000:
                im = Image.open(io.BytesIO(r.content))
                if im.width >= min_width and im.height >= min_height:
                    out_name = f"crawl_{idx}_{im.width}x{im.height}.jpg"
                    out_path = os.path.join(output_dir, out_name)
                    if im.mode in ("RGBA", "P", "LA"):
                        im = im.convert("RGB")
                    im.save(out_path, "JPEG", quality=92)
                    saved_files.append(out_path)
                    print(f"   📸 Downloaded high-res image: {out_name} ({im.width}x{im.height})")
        except Exception:
            continue

    print(f"✅ Total real images cached: {len(saved_files)}")
    return saved_files


def apply_rtl_complex_run(run, font_name="Tahoma", size_pt=11, bold=False, italic=False, color_rgb=None, is_rtl=True):
    """
    Applies strict OpenXML formatting for Persian/Arabic Complex Scripts:
    Sets w:rFonts (ascii, hAnsi, cs), w:rtl, w:lang, w:sz, w:szCs, w:b, w:bCs, w:i, w:iCs.
    This guarantees 100% native glyph shaping and zero garbled/disconnected text in Word.
    """
    rPr = run._r.get_or_add_rPr()

    # Clean existing font tags to avoid XML schema duplicates
    for child in list(rPr):
        if child.tag.endswith("rFonts"):
            rPr.remove(child)

    # 1. Font Family
    rPr.append(parse_xml(
        f'<w:rFonts {nsdecls("w")} w:ascii="{font_name}" w:hAnsi="{font_name}" w:cs="{font_name}"/>'
    ))

    # 2. RTL & Language
    if is_rtl:
        rPr.append(parse_xml(f'<w:rtl {nsdecls("w")} w:val="1"/>'))
        rPr.append(parse_xml(f'<w:lang {nsdecls("w")} w:val="fa-IR" w:bidi="fa-IR"/>'))
    else:
        rPr.append(parse_xml(f'<w:lang {nsdecls("w")} w:val="en-US"/>'))

    # 3. Size (Half-points)
    if size_pt:
        half_pts = int(size_pt * 2)
        rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="{half_pts}"/>'))
        if is_rtl:
            rPr.append(parse_xml(f'<w:szCs {nsdecls("w")} w:val="{half_pts}"/>'))

    # 4. Bold
    if bold:
        rPr.append(parse_xml(f'<w:b {nsdecls("w")}/>'))
        if is_rtl:
            rPr.append(parse_xml(f'<w:bCs {nsdecls("w")}/>'))

    # 5. Italic
    if italic:
        rPr.append(parse_xml(f'<w:i {nsdecls("w")}/>'))
        if is_rtl:
            rPr.append(parse_xml(f'<w:iCs {nsdecls("w")}/>'))

    # 6. Color
    if color_rgb:
        hex_color = f"{color_rgb[0]:02X}{color_rgb[1]:02X}{color_rgb[2]:02X}"
        rPr.append(parse_xml(f'<w:color {nsdecls("w")} w:val="{hex_color}"/>'))


def set_paragraph_formatting(p, is_rtl=True, line_spacing=1.35, space_after=6, space_before=0, justify=True):
    """Configures paragraph alignment, BiDi flow, and spacing."""
    pPr = p._p.get_or_add_pPr()
    if is_rtl:
        pPr.append(parse_xml(f'<w:bidi {nsdecls("w")} w:val="1"/>'))
        if justify:
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="both"/>'))
        else:
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="right"/>'))
    else:
        if justify:
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="both"/>'))
        else:
            pPr.append(parse_xml(f'<w:jc {nsdecls("w")} w:val="left"/>'))

    p.paragraph_format.line_spacing = line_spacing
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)


def add_hyperlink(paragraph, url, text, font_name="Tahoma", size_pt=10.5, color_hex="1E3A8A", is_rtl=True):
    """Inserts a clickable hyperlink with RTL complex script support."""
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = parse_xml(f'<w:hyperlink {nsdecls("w", "r")} r:id="{r_id}"/>')
    new_run = parse_xml(f'<w:r {nsdecls("w")}/>')
    rPr = parse_xml(f'<w:rPr {nsdecls("w")}/>')

    rPr.append(parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{font_name}" w:hAnsi="{font_name}" w:cs="{font_name}"/>'))
    if is_rtl:
        rPr.append(parse_xml(f'<w:rtl {nsdecls("w")} w:val="1"/>'))
    if color_hex:
        rPr.append(parse_xml(f'<w:color {nsdecls("w")} w:val="{color_hex}"/>'))
    rPr.append(parse_xml(f'<w:u {nsdecls("w")} w:val="single"/>'))

    if size_pt:
        half_pts = int(size_pt * 2)
        rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="{half_pts}"/>'))
        if is_rtl:
            rPr.append(parse_xml(f'<w:szCs {nsdecls("w")} w:val="{half_pts}"/>'))

    new_run.append(rPr)
    text_elem = parse_xml(f'<w:t {nsdecls("w")} xml:space="preserve">{text}</w:t>')
    new_run.append(text_elem)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def render_clean_news_article(doc, article, available_images, font_name="Tahoma", is_rtl=True, article_idx=0):
    """
    Renders a 100% publication-ready press release into the document.
    NO meta-instructions, NO prompt brackets, NO 'لید خبر:' or 'زاویه:' labels.
    """
    # 1. Headline (تیتر یک مطبوعاتی)
    headline_text = article.get("headline", "").strip()
    if headline_text:
        hp = doc.add_paragraph()
        set_paragraph_formatting(hp, is_rtl=is_rtl, line_spacing=1.2, space_after=6, space_before=14, justify=False)
        hrun = hp.add_run(headline_text)
        apply_rtl_complex_run(hrun, font_name=font_name, size_pt=18, bold=True, color_rgb=COLOR_HEADLINE, is_rtl=is_rtl)

    # 2. Sub-headline (سوتیتر)
    sub_text = article.get("sub_headline", "").strip()
    if sub_text:
        sp = doc.add_paragraph()
        set_paragraph_formatting(sp, is_rtl=is_rtl, line_spacing=1.25, space_after=12, space_before=0, justify=False)
        srun = sp.add_run(sub_text)
        apply_rtl_complex_run(srun, font_name=font_name, size_pt=12.5, italic=True, color_rgb=COLOR_SUBTITLE, is_rtl=is_rtl)

    # 3. Lede Paragraph (لید جذاب خبری - مستقیماً شروع می‌شود بدون هیچ برچسبی)
    lead_text = article.get("lead", "").strip()
    if lead_text:
        lp = doc.add_paragraph()
        set_paragraph_formatting(lp, is_rtl=is_rtl, line_spacing=1.4, space_after=12, space_before=4, justify=True)
        # Bold first few words for journalistic punch
        lrun = lp.add_run(lead_text)
        apply_rtl_complex_run(lrun, font_name=font_name, size_pt=11.5, bold=False, color_rgb=COLOR_BODY, is_rtl=is_rtl)

    # 4. Embedded Photo with Natural Caption (if images are available)
    if available_images:
        # Select image based on article index
        img_to_use = available_images[article_idx % len(available_images)]
        if os.path.isfile(img_to_use):
            try:
                img_p = doc.add_paragraph()
                img_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                img_p.paragraph_format.space_before = Pt(8)
                img_p.paragraph_format.space_after = Pt(4)
                img_run = img_p.add_run()
                img_run.add_picture(img_to_use, width=Inches(5.5))

                # Natural Journalistic Caption
                caption_text = article.get("photo_caption") or (
                    f"تصویر شماره {article_idx + 1}: بررسی رویکردهای نوین هوش مصنوعی و بهینه‌سازی موتورهای جستجو"
                    if is_rtl else
                    f"Figure {article_idx + 1}: Advanced AI Architecture and Search Engine Optimization"
                )
                cap_p = doc.add_paragraph()
                cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cap_p.paragraph_format.space_after = Pt(14)
                cap_p.paragraph_format.space_before = Pt(2)
                cap_run = cap_p.add_run(caption_text)
                apply_rtl_complex_run(cap_run, font_name=font_name, size_pt=9.5, italic=True, color_rgb=COLOR_CAPTION, is_rtl=is_rtl)
            except Exception as e:
                print(f"[Warning] Failed to insert picture {img_to_use}: {e}")

    # 5. Body Sections with Journalistic Interheadings
    for sec in article.get("body_sections", []):
        sec_h = sec.get("heading", "").strip()
        if sec_h:
            h2_p = doc.add_paragraph()
            set_paragraph_formatting(h2_p, is_rtl=is_rtl, line_spacing=1.2, space_after=6, space_before=14, justify=False)
            h2_run = h2_p.add_run(sec_h)
            apply_rtl_complex_run(h2_run, font_name=font_name, size_pt=13.5, bold=True, color_rgb=COLOR_ACCENT, is_rtl=is_rtl)

        sec_body = sec.get("content", "").strip()
        if sec_body:
            bp = doc.add_paragraph()
            set_paragraph_formatting(bp, is_rtl=is_rtl, line_spacing=1.35, space_after=8, space_before=0, justify=True)
            brun = bp.add_run(sec_body)
            apply_rtl_complex_run(brun, font_name=font_name, size_pt=11, color_rgb=COLOR_BODY, is_rtl=is_rtl)

    # 6. Natural Direct Quote Box (نقل‌قول طبیعی درون متن)
    quote_box = article.get("quote_box")
    if quote_box and quote_box.get("quote"):
        qp = doc.add_paragraph()
        set_paragraph_formatting(qp, is_rtl=is_rtl, line_spacing=1.35, space_after=10, space_before=10, justify=True)
        qp.paragraph_format.left_indent = Inches(0.4)
        qp.paragraph_format.right_indent = Inches(0.4)
        
        quote_str = quote_box["quote"].strip()
        if not quote_str.startswith("«") and not quote_str.startswith('"'):
            quote_str = f"«{quote_str}»" if is_rtl else f'"{quote_str}"'

        qrun = qp.add_run(quote_str)
        apply_rtl_complex_run(qrun, font_name=font_name, size_pt=11.5, italic=True, bold=True, color_rgb=COLOR_ACCENT, is_rtl=is_rtl)

        if quote_box.get("speaker"):
            spk_run = qp.add_run(f"\n— {quote_box['speaker'].strip()}")
            apply_rtl_complex_run(spk_run, font_name=font_name, size_pt=10, color_rgb=COLOR_SUBTITLE, is_rtl=is_rtl)

    # 7. Natural Concluding Section & Source Attribution
    source_url = article.get("source_link") or "https://molavi.pro"
    end_p = doc.add_paragraph()
    set_paragraph_formatting(end_p, is_rtl=is_rtl, line_spacing=1.35, space_after=14, space_before=12, justify=True)
    
    concl_text = (
        "علاقه‌مندان برای کسب اطلاعات تکمیلی، مطالعه مقالات پژوهشی و آشنایی با پروژه‌های تخصصی این حوزه می‌توانند به نشانی "
        if is_rtl else
        "For additional information, research papers, and technical insights, visit "
    )
    crun = end_p.add_run(concl_text)
    apply_rtl_complex_run(crun, font_name=font_name, size_pt=10.5, color_rgb=COLOR_BODY, is_rtl=is_rtl)
    add_hyperlink(end_p, source_url, source_url, font_name=font_name, size_pt=10.5, color_hex="1E3A8A", is_rtl=is_rtl)
    dot_run = end_p.add_run(" مراجعه فرمایند." if is_rtl else ".")
    apply_rtl_complex_run(dot_run, font_name=font_name, size_pt=10.5, color_rgb=COLOR_BODY, is_rtl=is_rtl)


def create_blank_press_doc():
    """Initializes a document with 1-inch margins."""
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)
    return doc


def process_campaign(input_json_path, output_dir, mode="both", font_rtl="Tahoma", font_ltr="Calibri", crawl_url=None):
    """
    Main compilation workflow:
    - Automatically scrapes real photos from source URL or campaign link
    - Sets up native Complex Script fonts (Tahoma/Arial/Vazirmatn)
    - Generates 100% publication-ready news Word (.docx) files without meta-labels
    """
    if not DOCX_AVAILABLE:
        print("[Error] 'python-docx' is not installed.")
        sys.exit(1)

    with open(input_json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    articles = data.get("articles", [])
    lang = data.get("language", "fa").lower()
    is_rtl = lang in RTL_LANGUAGES
    font = font_rtl if is_rtl else font_ltr

    # Discover and prepare real images
    target_url = crawl_url or data.get("source_url") or data.get("canonical_source_link")
    cached_images = []

    # Check local asset folder first
    asset_dir = Path("assets/crawled")
    if asset_dir.is_dir():
        for ext in ("*.jpg", "*.jpeg", "*.png"):
            cached_images.extend([str(p) for p in asset_dir.glob(ext)])

    # If no images cached and URL is available, crawl automatically!
    if not cached_images and target_url:
        cached_images = crawl_website_images(target_url, output_dir="assets/crawled")

    print(f"🚀 Processing campaign: '{data.get('project_title', 'PR Campaign')}'")
    print(f"📊 Total Articles: {len(articles)} | Language: {lang.upper()} | Font: {font}")
    print(f"🖼️ Real Photos Available: {len(cached_images)}")

    # Mode: Split Individual Articles
    if mode in ("split", "both"):
        split_dir = out_path / "individual_articles"
        split_dir.mkdir(parents=True, exist_ok=True)
        for idx, art in enumerate(articles):
            doc = create_blank_press_doc()
            render_clean_news_article(doc, art, cached_images, font_name=font, is_rtl=is_rtl, article_idx=idx)
            slug = f"article_{idx + 1:02d}_{art.get('angle_name', 'news').replace(' ', '_').lower()[:25]}.docx"
            file_path = split_dir / slug
            doc.save(str(file_path))
        print(f"✅ Generated {len(articles)} clean individual Word files in: {split_dir}")

    # Mode: Master Compilation
    if mode in ("bundle", "both"):
        master_doc = create_blank_press_doc()

        # Cover Title
        cov_p = master_doc.add_paragraph()
        set_paragraph_formatting(cov_p, is_rtl=is_rtl, space_after=12, space_before=36, justify=False)
        cov_run = cov_p.add_run(data.get("project_title", "بسته جامع رپرتاژ و مقالات خبری"))
        apply_rtl_complex_run(cov_run, font_name=font, size_pt=22, bold=True, color_rgb=COLOR_HEADLINE, is_rtl=is_rtl)

        sub_p = master_doc.add_paragraph()
        set_paragraph_formatting(sub_p, is_rtl=is_rtl, space_after=24, space_before=0, justify=False)
        sub_desc = (
            f"مجموعه {len(articles)} خبر و گزارش مطبوعاتی آماده ارسال به رسانه‌ها و خبرگزاری‌ها\n"
            f"طراح و معمار: {data.get('primary_entity', 'مهندس تقی مولوی')} | منبع رسمی: https://molavi.pro"
            if is_rtl else
            f"Complete compilation of {len(articles)} news releases ready for press syndication\n"
            f"Author: {data.get('primary_entity', 'Taghi Molavi')} | Source: https://molavi.pro"
        )
        srun = sub_p.add_run(sub_desc)
        apply_rtl_complex_run(srun, font_name=font, size_pt=11, color_rgb=COLOR_SUBTITLE, is_rtl=is_rtl)

        master_doc.add_page_break()

        for idx, art in enumerate(articles):
            render_clean_news_article(master_doc, art, cached_images, font_name=font, is_rtl=is_rtl, article_idx=idx)
            if idx < len(articles) - 1:
                master_doc.add_page_break()

        master_file = out_path / "Press_Pack_Master_Compilation.docx"
        master_doc.save(str(master_file))
        print(f"✅ Master compilation generated: {master_file}")


def main():
    parser = argparse.ArgumentParser(
        description="GEO & AEO News Media PR Pack Generator - Clean, Publication-Ready Word (.docx)",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    parser.add_argument("-i", "--input", required=True, help="Path to input JSON file")
    parser.add_argument("-o", "--output-dir", default="output/docx", help="Directory to save .docx files")
    parser.add_argument("-m", "--mode", choices=["bundle", "split", "both"], default="both", help="Output mode")
    parser.add_argument("--font-rtl", default="Tahoma", help="Font family for Persian/Arabic (Tahoma, Arial, Vazirmatn)")
    parser.add_argument("--font-ltr", default="Calibri", help="Font family for English/Turkish")
    parser.add_argument("--crawl", default=None, help="Optional URL to scrape images from")

    args = parser.parse_args()
    process_campaign(args.input, args.output_dir, args.mode, args.font_rtl, args.font_ltr, args.crawl)


if __name__ == "__main__":
    main()
