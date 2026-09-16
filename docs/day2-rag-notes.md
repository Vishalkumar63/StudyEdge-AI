# Day 2 – Local RAG Pipeline

## StudyEdge AI

StudyEdge AI is a private, local AI study assistant that retrieves relevant information from uploaded academic documents and uses a local language model to generate grounded answers.

## Day 2 Objective

The objective of Day 2 was to build the first working Retrieval-Augmented Generation (RAG) pipeline.

## Implemented Components

### 1. PDF Text Extraction
PyMuPDF is used to extract selectable text from uploaded PDF documents.

### 2. Text Chunking
Extracted document text is divided into smaller chunks using sentence-boundary-aware chunking.

- Chunk size: 1200 characters
- Chunk overlap: 200 characters

### 3. Semantic Retrieval
Sentence Transformers with the `all-MiniLM-L6-v2` embedding model are used to convert document chunks and user queries into embeddings.

### 4. Keyword Retrieval
Important query terms are matched against document chunks to improve retrieval when exact terminology is important.

### 5. Hybrid Retrieval
Semantic and keyword scores are combined:

- Semantic retrieval: 70%
- Keyword retrieval: 30%

The top relevant document chunks are provided as evidence to the language model.

### 6. Local Language Model
StudyEdge AI currently uses:

- Model: Llama 3.2 3B
- Runtime: Ollama
- API: Local Ollama API

The model runs locally rather than sending document content to a cloud language-model API.

### 7. Grounded Answer Generation
The language model is instructed to answer using retrieved document evidence and avoid inventing information that is not present in the document.

### 8. Evidence Display
The application displays the retrieved document chunks used to generate the answer, allowing users to verify the source information.

## Testing

The system has been tested with:

- Direct factual questions
- Incorrect assumptions
- Missing-information questions
- Multi-part questions
- Numerical questions
- Rule-based calculations
- Exception-based questions
- Cross-chunk questions

## Current Limitation

Follow-up questions containing references such as "that", "this", or "it" are not yet fully resolved using conversation context.

For example:

> What is the submission deadline?

followed by:

> Is that before 10 September?

requires conversation-aware query rewriting.

## Day 2 Outcome

A working local RAG prototype has been created that can:

1. Accept an academic PDF.
2. Extract its text.
3. Split the text into chunks.
4. Retrieve relevant evidence.
5. Send the evidence to a local Llama model.
6. Generate a grounded answer.
7. Display the evidence used for the answer.

This establishes the core intelligence pipeline for StudyEdge AI.
