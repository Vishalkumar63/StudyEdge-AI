# StudyEdge AI Test Plan

## Ask Documents

### Test 1 — Direct numerical retrieval

Question:

> What is the total number of vehicles in the dataset?

Expected:

> 4,750

---

### Test 2 — Vehicle variance

Question:

> What is the speed variance for 3-W?

Expected:

> 148.411 (km/h)²

---

### Test 3 — Variance statistic

Question:

> What is the chi-square statistic for the 3-W speed variance test?

Expected:

> 436.330

---

### Test 4 — Variance p-value

Question:

> What was the p-value for the 3-W speed variance test?

Expected:

> 1.3e-07

---

### Test 5 — Welch ANOVA

Question:

> What is the Welch ANOVA p-value for speed?

Expected:

> 2.613e-81

---

### Test 6 — Acceleration

Question:

> What is the Welch ANOVA p-value for acceleration?

Expected:

> 2.187e-66

---

### Test 7 — Compliance

Question:

> What does the chi-square test tell us about speed-limit compliance and vehicle group?

Expected:

The test indicates a statistically significant association between
speed-limit compliance and vehicle group.

---

### Test 8 — Hallucination protection

Question:

> What was the exact p-value of the t-test comparing electric cars with hydrogen buses in Assignment 1?

Expected:

The system should state that the documents do not provide enough
information.

---

## Follow-up Test

Start with:

> What is the speed variance for 3-W?

Then:

> Is that greater than 100?

Then:

> What was its p-value?

Then:

> Does that mean we reject the null hypothesis?

The system should preserve the previous context while answering the
specific new question.

---

## Study Mode

Test the following modes independently:

- MCQs
- Flashcards
- Viva
- Key Concepts
- Summary

Study Mode should use evidence from the selected document/topic and should
not mix unrelated documents.

---

## Known Issues Under Development

- Follow-up intent classification
- Multi-part numerical questions
- Document-family-aware retrieval
- Generated study-material consistency
- Numerical answer validation
