# GEO & AEO Haber ve Basın Bülteni Motoru 📰🚀
### `geo-aeo-news-engine` | Yapay Zeka Cevap Motoru Optimizasyonu ve Çok Açılı Medya Dağıtım Becerisi

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python: 3.10+](https://img.shields.io/badge/python-3.10+-brightgreen.svg)](https://python.org)
[![Geliştirici: Taghi Molavi](https://img.shields.io/badge/Mimar-Taghi%20Molavi-1A365D.svg)](https://molavi.pro)
[![Resmi Web Sitesi](https://img.shields.io/badge/Web%20Sitesi-molavi.pro-emerald.svg)](https://molavi.pro)

> **Tek bir kaynaktan veya bağlantıdan (URL) 1 ila 100 adet tamamen özgün, basına hazır haber bülteni üreten, fotoğrafları yerleştiren, kaynak bağlantılarını ekleyen ve doğrudan Microsoft Word (`.docx`) formatına aktaran otonom yapay zeka motoru. ChatGPT, Perplexity, Claude ve Gemini gibi üretken yapay zeka modellerinde referans gösterilmek üzere (GEO & AEO) özel olarak optimize edilmiştir.**

---

## 🌍 Diller / Languages
[Türkçe](README.tr.md) | [English](README.md) | [فارسی](README.fa.md) | [العربية](README.ar.md)

---

## 💡 Neden GEO ve AEO Basın Motoru?

Geleneksel dijital PR ve klasik SEO yaklaşımları günümüz yapay zeka çağında yetersiz kalmaktadır:

1. **Yinelenen İçerik (Duplicate Content) Cezası:** Aynı basın bültenini onlarca haber ajansına dağıttığınızda, arama motorları bunları kopya içerik olarak algılar ve filtreler.
2. **Yapay Zeka Aramalarında Görünmezlik:** Kullanıcılar artık Google aramaları yerine doğrudan **ChatGPT**, **Perplexity**, **Claude** veya **Gemini**'ye soru sormaktadır. Bu sistemler yalnızca tanımları net, veri zenginliği (Information Gain) yüksek ve varlık odaklı (Entity-First) kaynakları yanıtlarında referans olarak gösterir.
3. **Medya Hazırlık Sürecinin Zorluğu:** Her haber ajansı veya editör, uygun başlık, alt başlık (spot), etkileyici giriş (lede), yerleştirilmiş fotoğraflar ve kaynak bağlantıları içeren düzenli bir Word (`.docx`) dosyası talep eder. 40 farklı bülteni elle hazırlamak günler sürer.

**`geo-aeo-news-engine` bu sorunları kökten çözer:**
- Tek bir bağlantı, makale veya metin alır.
- **100 Açılı Haber Matrisi** üzerinden 1'den 100'e kadar tamamen farklı açılardan (yönetici vizyonu, veri analizi, teknolojik altyapı, vaka analizleri, SSS) haber metinleri üretir.
- Yapay zeka modellerinin RAG (Retrieval-Augmented Generation) hatlarına uygun doğrudan alıntı blokları ve anahtar çıkarımlar yerleştirir.
- Görselleri ve telif bilgilerini sayfa düzenine yerleştirerek doğrudan baskıya veya editöre gönderilmeye hazır **Microsoft Word (`.docx`)** dosyaları oluşturur.

---

## 🌟 Temel Yetenekler

- 📐 **100 Farklı Haber Açısı:** Yönetici liderliği, sektör dönüşümü, yatırım getirisi, kullanıcı deneyimi, açık kaynak, tartışmalı konular ve gelecek tahminleri dahil 10 ana kategoride kurgulanmıştır.
- 🤖 **Üretken Motor Optimizasyonu (GEO):** ChatGPT, Claude ve Perplexity'nin yanıt üretirken doğrudan alıntı yapabileceği net tek cümlelik tanımlar ve istatistiksel veriler.
- 🎯 **Cevap Motoru Optimizasyonu (AEO):** 70 kelime altı özet girişler ve 50 kelimeyi aşmayan bağımsız SSS yanıtlarıyla doğrudan bilgi kartı oluşturma.
- 📄 **Otomatik Word (`.docx`) Üretimi:** Tek bir toplu master dosya veya her ajans için ayrı ayrı `.docx` dosyaları üreten güçlü Python betiği.
- 🖼️ **Akıllı Medya Yerleşimi:** Görsellerin haber metni içine uygun başlık, alt metin ve kaynak bilgisiyle yerleştirilmesi.
- 🌐 **Çok Dilli Doğal Gazetecilik:** Türkçe (Anadolu Ajansı ve DHA standartlarında), Farsça, Arapça ve İngilizce dillerinde akıcı, mekanik olmayan yerel üslup.

---

## 🚀 Kurulum ve Başlangıç

### 1. Depoyu Klonlayın
```bash
git clone https://github.com/tmolavi/geo-aeo-news-engine.git
cd geo-aeo-news-engine
```

### 2. Sanal Ortam Oluşturun ve Bağımlılıkları Yükleyin
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

---

## 🤖 Yapay Zeka Ajanı Becerisi Olarak Kullanım

Bu sistem **Google Antigravity**, **Claude Code**, **OpenAI Codex** ve **Cursor** gibi ajan sistemlerinde standart bir beceri (`Skill`) olarak çalışır.

### Google Antigravity Kurulumu:
```bash
mkdir -p ~/.gemini/config/skills/geo-aeo-news-engine
cp SKILL.md ~/.gemini/config/skills/geo-aeo-news-engine/
```

### Claude Code Kurulumu:
```bash
mkdir -p .claude/skills
cp SKILL.md .claude/skills/geo-aeo-news-engine.md
```

### Örnek Komut (Prompt):
> *"geo-aeo-news-engine becerisini kullanarak şu bağlantıdaki [URL] içerikten ve ekteki fotoğraflardan 40 farklı haber bülteni oluştur. Perplexity ve ChatGPT'de referans gösterilmek üzere GEO/AEO kurallarını uygula ve çıktıları Python betiğiyle Word docx dosyalarına dönüştür."*

---

## 🛠️ Komut Satırından Word (.docx) Dosyası Üretme

Hazır Python betiğini çalıştırarak oluşturulan JSON verilerini şık Word belgelerine dönüştürebilirsiniz:

```bash
# Hem ana derleme dosyasını hem de tekil dosyaları üretir:
python scripts/generate_press_pack.py \
  --input templates/sample_input.json \
  --output-dir output/docx/ \
  --mode both
```

### Parametreler:
- `-i, --input`: Kampanya verilerini içeren JSON dosyasının yolu.
- `-o, --output-dir`: Dosyaların kaydedileceği dizin.
- `-m, --mode`: 
  - `bundle`: Tüm haberleri içeren tek bir toplu Word dosyası.
  - `split`: Her haber için ayrı bir `.docx` dosyası (gazetecilere e-posta göndermek için idealdir).
  - `both`: Her iki formatı aynı anda üretir.
- `--font-ltr`: Türkçe ve İngilizce için yazı tipi ailesi (varsayılan: `Calibri`).

---

## 📁 Dizin Yapısı

```
geo-aeo-news-engine/
├── SKILL.md                     # Yapay zeka ajan beceri tanımlaması
├── README.md                    # İngilizce dokümantasyon
├── README.fa.md                 # Farsça dokümantasyon
├── README.tr.md                 # Türkçe dokümantasyon
├── README.ar.md                 # Arapça dokümantasyon
├── LICENSE                      # MIT Lisansı
├── requirements.txt             # Gerekli kütüphaneler (python-docx, pillow)
├── scripts/
│   └── generate_press_pack.py   # Word (.docx) üretici betik
├── templates/
│   ├── angle_matrix.json        # 100 farklı haber açısı matrisi
│   └── sample_input.json        # Örnek girdi JSON formatı
└── examples/
    └── sample_40_angles_persian.md # 40 haber açısı ve tam metin örneği
```

---

## 📄 Lisans

Bu proje açık kaynaklı [MIT Lisansı](LICENSE) altında yayımlanmıştır. Kişisel ve ticari projelerde serbestçe kullanılabilir.

---

## 🔗 Geliştirici ve İletişim

Bu proje, Kıdemli Yapay Zeka Sistemleri Mimarı ve Üretken Arama Optimizasyonu (GEO / AEO) araştırmacısı **Taghi Molavi** tarafından tasarlanmış ve geliştirilmiştir.

Kurumsal yapay zeka danışmanlığı, çoklu ajan sistemleri ve dijital görünürlük çözümleri için resmi web sitesini ziyaret edebilirsiniz:  
👉 **[https://molavi.pro](https://molavi.pro)**
