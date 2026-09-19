# StudyEdge AI — Project Status

## Completed

- Project concept finalized
- GitHub repository created
- Python environment configured
- Streamlit interface implemented
- PDF upload implemented
- PDF text extraction implemented
- PyMuPDF integration implemented
- OCR support added
- Document chunking implemented
- TF-IDF retrieval implemented
- Local Ollama integration implemented
- Llama 3.2 3B integration implemented
- Evidence/source display implemented
- Follow-up question handling implemented
- Study Mode implemented
- MCQ generation implemented
- Flashcard generation implemented
- Viva generation implemented
- Key Concept generation implemented
- Summary generation implemented
- Qualcomm AI Hub client configured
- Snapdragon target devices verified

## Current Development

The following areas are being improved:

1. Follow-up numerical reasoning
2. Intent classification
3. Multi-part numerical questions
4. Hypothesis-test reasoning
5. Document-family-aware retrieval
6. Study Mode evidence consistency
7. Generated MCQ validation
8. Generated flashcard validation
9. Generated viva-answer validation
10. Summary factual consistency

## Validation Examples

| Test | Expected |
|---|---:|
| Total vehicles | 4,750 |
| 3-W speed variance | 148.411 (km/h)² |
| 3-W variance χ² | 436.330 |
| 3-W variance p-value | 1.3e-07 |
| Welch speed p-value | 2.613e-81 |
| Welch acceleration p-value | 2.187e-66 |
| Compliance χ² | 119.892 |
| Compliance p-value | 8.142e-26 |
| Bus erraticness index | 1.941 |

## Important Development Principle

StudyEdge AI should prefer evidence-grounded answers over unsupported
generation.

When the uploaded documents do not contain sufficient information, the
system should explicitly state that the information is unavailable rather
than inventing an answer.
