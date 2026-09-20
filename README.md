# 📚 StudyEdge AI

## Private Local AI Study Assistant

StudyEdge AI is a local document intelligence assistant designed to help students understand and study from academic documents.

Students can upload academic documents, ask questions, retrieve evidence, and generate study material using local AI.

---

## 🚀 Features

- 📄 PDF document understanding
- 📊 Excel/XLSX spreadsheet understanding
- 📑 PowerPoint/PPTX understanding
- 📝 DOCX document support
- 🖼️ Image and OCR support
- 🔎 Evidence-grounded document retrieval
- 🤖 Local AI reasoning
- 🎓 Study Mode
- 📝 MCQ generation
- 🧠 Flashcards
- 🎤 Viva questions
- 📌 Key concepts
- 📚 Summaries
- ✍️ Typo-aware queries
- 💾 Cached document indexing

---

## 📂 Supported Formats

- PDF
- PPTX
- XLSX
- XLSM
- DOCX
- CSV
- TXT
- MD
- JSON
- XML
- HTML
- PNG
- JPG
- JPEG
- WEBP
- BMP
- TIF
- TIFF

---

## 🏗️ Architecture

```text
                    StudyEdge AI
                         │
                         ▼
                   Streamlit UI
                         │
                         ▼
                Document Ingestion
                         │
          ┌──────────────┼──────────────┐
          ▼              ▼              ▼
         PDF            PPTX           XLSX
          │              │              │
          └──────────────┼──────────────┘
                         ▼
              Text / Table Extraction
                         │
                         ▼
                      Chunking
                         │
                         ▼
                  TF-IDF Retrieval
                         │
                         ▼
                   Evidence Layer
                         │
                         ▼
                 Local LLM Reasoning
                         │
                         ▼
                    User Answer
