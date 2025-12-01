# 📘 DocMaster Formatter  
**Automated Research Paper Formatter | APA · IEEE · MLA · Custom Styles**

DocMaster is an intelligent desktop application built using **Python + PyQt5** that automatically formats academic documents—including DOCX, TXT, and PDF—into clean, professional, publication-ready research papers.

It eliminates hours of manual formatting by detecting structure, extracting images/tables, identifying sections (with optional NLP), and applying academic styles automatically.

---

## 🚀 Features

### 🔹 1. Multi-Format Import
DocMaster supports:
- **DOCX** (full text, tables, inline images)
- **TXT** (raw text)
- **PDF** (smart extraction: text, tables, images)

### 🔹 2. Automatic Section Detection
Using a hybrid detection system:
- Rule-based heading recognition  
- Optional ML classifier (SVM + TF-IDF)  
- Virtual image detection (for flattened PDF figures)

Recognized sections include:
Abstract, Introduction, Literature Review, Methodology, Results,
Discussion, Conclusion, Future Work, References, Acknowledgement,
Certificate, Declaration, Appendix, Custom Sections


### 🔹 3. Smart Image Preservation
DocMaster:
- Extracts inline DOCX images correctly  
- Extracts PDF images using bounding-box cropping  
- Keeps **exact original order**  
- Maintains proportional size  
- Detects figure captions automatically  
- Creates placeholders for non-extractable PDF figures

### 🔹 4. Complete Academic Formatting
Supports:
- **APA**
- **IEEE**
- **MLA**
- **Custom**

Formatting includes:
- Fonts  
- Margins  
- Line spacing  
- Headings  
- Alignment  
- Title page  
- Table formatting (T-Grid style)  
- Image scaling + captions  

### 🔹 5. Clean PyQt5 User Interface
Tabs:
1. Upload & Detect  
2. Structured Preview  
3. Formatting & Style  
4. Formatted Preview  
5. Export  

### 🔹 6. Export Options
- Export to **DOCX**  
- Export to **PDF**  

---

## 🧠 NLP Section Classifier (Optional)

DocMaster includes a trainable SVM-based classifier:

- TF-IDF vectorizer (1–3 n-grams)  
- Linear SVM  
- 70–80% accuracy on boosted dataset  
- Hybrid rule-based + ML pipeline  
- Handles ambiguous cases better than headings alone  

Model files:


vectorizer.pkl
classifier.pkl


---

## 📁 Project Structure



DocMaster/
│
├── main.py # UI Application
├── document_handler.py # Parsing engine (DOCX/PDF/TXT)
├── export_service.py # DOCX/PDF generator
├── section_predictor.py # NLP classifier (optional)
│
├── trained_tfidf_model/ # ML model directory
│ ├── vectorizer.pkl
│ └── classifier.pkl
│
├── resources/ # UI icons, assets (optional)
