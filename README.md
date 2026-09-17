# GEO & AEO News Media PR Engine 📰🚀
### Generative Engine Optimization (GEO) & Answer Engine Optimization (AEO) Syndication Engine

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python: 3.10+](https://img.shields.io/badge/python-3.10+-brightgreen.svg)](https://python.org)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-emerald.svg)](https://github.com/tmolavi/geo-aeo-news-engine/pulls)
[![Maintained by Taghi Molavi](https://img.shields.io/badge/Architect-Taghi%20Molavi-1A365D.svg)](https://molavi.pro)

> **Autonomous multi-angle news rewriting, digital PR syndication, and AI answer engine optimization engine.**  
> Ingests a single source (URL or text) and media assets (photos) to generate **1 to 100 distinctive, publication-ready news and PR articles**, complete with press-ready headlines, source links, photo placements, and automated export to Microsoft Word (`.docx`).

---

## 🌍 Languages / مستندات به زبان‌های مختلف
[English](README.md) | [فارسی](README.fa.md) | [Türkçe](README.tr.md) | [العربية](README.ar.md)

---

## 💡 Why GEO & AEO News Engine?

Traditional digital PR and SEO suffer from critical modern failures:
1. **Duplicate Content Penalties**: Sending identical press releases to 40 news agencies causes search engines to filter them out as syndicated duplicates.
2. **Invisible to AI Search**: When users ask **ChatGPT, Perplexity, Claude, or Gemini** for recommendations or expert insights, conventional keyword-stuffed articles are ignored by Retrieval-Augmented Generation (RAG) pipelines.
3. **Friction in Newsrooms**: Journalists and editors reject raw copy that lacks editorial headlines, inverted pyramid leads, formatted image placements, and structured metadata.

**GEO & AEO News Engine solves this completely**:
- **1 to 100 Unique Journalistic Angles**: Derived from a 100-angle editorial matrix (executive vision, benchmark data, contrarian debates, deep tech architecture, case studies, and Q&As).
- **High Entity Density & Information Gain**: Engineered specifically so AI answer engines register your primary entity as an authoritative source and cite it in conversational responses.
- **Automated Word (.docx) Packaging**: Instantly converts generated articles into cleanly styled, publication-ready `.docx` files with embedded images, captions, photo credits, and source links.
- **Cross-Lingual Native Output**: Native fluency in Persian (RTL), Turkish, Arabic (RTL), and English without mechanical translation artifacts.

---

## 🚀 Key Features

| Feature | Description |
|---|---|
| 📐 **100-Angle Editorial Matrix** | Pre-built taxonomy spanning 10 news pillars to ensure every single article has a fresh, engaging perspective. |
| 🤖 **GEO Optimization** | Entity-first architecture, citable definitions, and statistical proof points designed for RAG retrieval in Perplexity & ChatGPT. |
| 🎯 **AEO Direct Extraction** | Standalone TL;DR summaries and 50-word FAQ answers tailored for featured snippets and AI answer cards. |
| 📄 **Automated Word (`.docx`) Export** | Outputs either a bundled master compilation or individual press files ready to attach to pitch emails. |
| 🖼️ **Smart Media Integration** | Automatic image placement, captions, alt-texts, and photo credits. |
| 🌐 **Full RTL & Multilingual Support** | Proper bidirectional text, fonts (IRANSans, Vazirmatn, Calibri), and cultural press phrasing. |

---

## 📦 Installation & Setup

### 1. Clone the Repository
```bash
git clone https://github.com/tmolavi/geo-aeo-news-engine.git
cd geo-aeo-news-engine
```

### 2. Set Up Virtual Environment & Dependencies
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

---

## 🛠️ Usage

### Using as an AI Agent Skill (Antigravity / Claude Code / Codex / Cursor)
This repository functions out-of-the-box as an AI Agent Skill. Simply install `SKILL.md` into your agent's configuration:

```bash
# For Google Antigravity
mkdir -p ~/.gemini/config/skills/geo-aeo-news-engine
cp SKILL.md ~/.gemini/config/skills/geo-aeo-news-engine/

# For Claude Code
mkdir -p .claude/skills
cp SKILL.md .claude/skills/geo-aeo-news-engine.md
```

Then prompt your agent:
> *"Using the geo-aeo-news-engine skill, take this article link [URL] and generate 40 distinct news pieces with photos and compile them into a Word docx file."*

### Generating Word (.docx) Files via CLI
Run the bundled script to compile your articles into Microsoft Word documents:

```bash
# Generate both a master compilation and individual press files:
python scripts/generate_press_pack.py \
  --input templates/sample_input.json \
  --output-dir output/campaign_pack \
  --mode both
```

#### CLI Options:
- `-i, --input`: Path to structured JSON campaign data.
- `-o, --output-dir`: Output directory for generated `.docx` documents.
- `-m, --mode`: `bundle` (single master doc), `split` (individual files per article), or `both`.
- `--font-rtl`: Font name for Persian/Arabic text (default: `IRANSans`).
- `--font-ltr`: Font name for English/Turkish text (default: `Calibri`).

---

## 📁 Repository Structure

```
geo-aeo-news-engine/
├── SKILL.md                     # Universal AI Agent Skill definition
├── README.md                    # English documentation
├── README.fa.md                 # راهنمای جامع به زبان فارسی
├── README.tr.md                 # Türkçe kullanım kılavuzu
├── README.ar.md                 # الدليل الشامل باللغة العربية
├── LICENSE                      # MIT License
├── requirements.txt             # Python dependencies (python-docx, pillow)
├── scripts/
│   └── generate_press_pack.py   # Word (.docx) generator with RTL & styling
├── templates/
│   ├── angle_matrix.json        # 100-angle editorial matrix across 10 categories
│   └── sample_input.json        # Sample JSON input payload
└── examples/
    └── sample_40_angles_persian.md # 40 Persian PR angles & full sample text
```

---

## 🤝 Contributing

Contributions, bug reports, and new editorial angle templates are welcome! Feel free to submit a pull request or open an issue.

---

---

## 🔗 Related Projects

Part of the **Molavi AI Engineering Ecosystem**:

* [**geo-scope**](https://github.com/tmolavi/geo-scope): Multi-model empirical AI visibility benchmark engine.
* [**answerpath-geo**](https://github.com/tmolavi/answerpath-geo): Privacy-first question discovery and intent stratification engine.
* [**sage-audit**](https://github.com/tmolavi/sage-audit): 3-Pillar static audit engine for SEO, AEO, and GEO.
* [**siteprobe**](https://github.com/tmolavi/siteprobe): Autonomous crawler and safe source code fixer.
* [**laravel-ai-summary**](https://github.com/tmolavi/laravel-ai-summary): Provider-agnostic AI summarization package.
* [**Ecosystem Map**](https://github.com/tmolavi/geo-scope/blob/main/docs/GITHUB_ECOSYSTEM.md): Complete architecture and evidence flow.

---

## 📄 License & Author

Developed by **Taghi Molavi** — [molavi.pro](https://molavi.pro)  
Distributed under the open-source [MIT License](LICENSE).

