# StudyEdge AI

### Private Local AI Study Assistant

StudyEdge AI is a document-intelligence study assistant designed to help students
understand academic PDFs using fast retrieval, evidence-grounded answers, and
local AI reasoning.

The application allows students to upload academic documents and interact with
them through question answering and study-material generation.

---

## 🚀 Features

- 📄 PDF document upload
- 🔎 Fast document retrieval
- 💬 Ask questions about uploaded documents
- 🔄 Follow-up question understanding
- 📚 Evidence-grounded answers
- 🧠 Local LLM reasoning
- 📝 MCQ generation
- 🧠 Flashcard generation
- 🎤 Viva question generation
- 📌 Key concept generation
- 📖 Document summarization
- 🔤 OCR support for scanned PDFs
- 📑 Source/page evidence display
- ⚡ TF-IDF based fast retrieval

---

## 🏗️ Architecture

```text
                 ┌─────────────────────┐
                 │       Student       │
                 └──────────┬──────────┘
                            │
                            ▼
                 ┌─────────────────────┐
                 │    Streamlit UI     │
                 └──────────┬──────────┘
                            │
                            ▼
                 ┌─────────────────────┐
                 │    PDF Extraction   │
                 │      + OCR          │
                 └──────────┬──────────┘
                            │
                            ▼
                 ┌─────────────────────┐
                 │ Chunking + TF-IDF   │
                 │      Indexing       │
                 └──────────┬──────────┘
                            │
                            ▼
                 ┌─────────────────────┐
                 │ Evidence Retrieval  │
                 └──────────┬──────────┘
                            │
                            ▼
                 ┌─────────────────────┐
                 │ Local LLM Reasoning │
                 │    Llama 3.2 3B    │
                 └──────────┬──────────┘
                            │
                            ▼
                 ┌─────────────────────┐
                 │ Answer / Study Mode │
                 └─────────────────────┘
