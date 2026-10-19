# Change Log: Analyser Reads the Full Document

19 Oct 2026 · Sai Praneeth Meduri · Branch `analyser-truncation-chnage`

This is Step 1 of [PROMPT_IMPROVEMENTS.md](PROMPT_IMPROVEMENTS.md).

## What was wrong

When scoring, the Analyser sent the AI only the **first 4,000–6,000 characters** of the case study (about 500–780 words) and the **first 5,000 characters** of the teaching note. It never read the rest of the document, and scored it as "missing".

In the 12 benchmark reports, this meant:
- Long cases (Lesson-Drawing, Decision-Forcing) were only 16–32% read and scored 42–52%. Short caselets scored 66–71%.
- Citation Quality, Theories & Frameworks and Additional Readings scored 0 in all 12 reports. The content they check for was in the documents, after the cut-off.

## What changed

The cut-offs were removed. Every Analyser step now sends the **full** case study and teaching note to the AI.

| Analyser step | Before | Now | Line in `app.py` |
|---|---|---|---|
| Executive summary | first 4,000 characters of the case | full case | 707 |
| Sector tags and keywords ("Generate Summary" button) | first 5,000 characters of the case | full case | 724 |
| Informational criteria (competency, SDG) | first 5,000 characters of the case and of the teaching note | full case and teaching note | 780, 783 |
| Scored criteria in Area 3 (teaching-note alignment) | first 5,000 characters of the case and of the teaching note | full case and teaching note | 811, 814 |
| Scored criteria in Areas 1, 2 and 4 | first 6,000 characters of the case | full case | 847 |
| KCM competency mapping | first 5,000 characters of the case | full case | 954 |
| Sector tags and keywords ("Perform Full Assessment") | first 5,000 characters of the case | full case | 991 |

In code, each line changed from, for example, `{case_text[:6000]}` to `{case_text}`.

## What did not change

- The rubric, criteria, scoring guides, weights and grade bands
- The wording of every prompt
- The Generator
- The Writing Assistant, which already read the full text
- Error handling: a failed AI call still scores 0 (see below)

## What to expect

- **Scores change for every document.** Do not compare reports from before this change with reports after it.
- **Most long cases should score higher**, especially on Additional Readings, Theories & Frameworks, Impact, Lessons, Delivery and Citation Quality.
- **Some criteria may score lower**, because the Analyser now also sees defects in later sections.
- **Each assessment costs more**: about 1.2× more AI input for a short caselet, and up to about 5× for a long Lesson-Drawing case.

## Known risk

Larger requests make it more likely that the AI account hits its per-minute limit. A failed criterion still scores **0**, and its reasoning starts with **"Error:"**. Check reports for this. If it happens often, add retries next.

## Results (re-run on 19 Oct 2026)

The new reports are in `benchmark_docs/benckmark_analyser_truncation_change/`. The old ones are in `benchmark_docs/benchmark_model_change/`.

### Gemini analyser: the fix works

| Case type | Gemini-written case | GPT-4o-written case |
|---|---|---|
| Caselet | 65.7% → **91.1%** | 68.3% → **79.5%** |
| Decision-Forcing | 47.3% → **86.8%** | 42.5%* → **74.5%** |
| Lesson-Drawing | 45.0% → **91.4%** | 42.2% → **70.2%** |

\*The old 42.5% report had scored the wrong files. The new run uses the correct ones.

- **No more cut-off complaints.** No report says the text "ends abruptly" or quotes a cut-off point.
- **The three criteria that scored 0 everywhere now score:**

| Criterion | Before | After |
|---|---|---|
| Additional Readings | 0 | 1 in all six reports |
| Theories & Frameworks | 0 | 2 for the three Gemini-written teaching notes; still 0 for the GPT-4o ones, which is fair because none names a framework |
| Citation Quality | 0 | 1–2 |

- **Long cases gained the most:**

| Criterion | Before | After |
|---|---|---|
| Impact Visibility | 0 | 2–3 |
| Delivery & Implementation | 1 | 3 |
| Best Practices & Lessons | 0 | 2 |
| Factual Correctness (Gemini-written cases) | 0–1 | 3 |

- **No criterion went down.** Captivating Hook and Language Quality did not change. These are Generator issues, handled in Step 2.
- **The gap between generators now shows.** Gemini-written cases score 87–91%, and GPT-4o-written cases score 70–80%. Before the fix, the two were only about 3 points apart.
- **The gains were bigger than estimated.** [PROMPT_IMPROVEMENTS.md](PROMPT_IMPROVEMENTS.md) estimated about 62–72% for the Lesson-Drawing case; the Gemini-written one scored 91.4%.

### GPT-4o analyser: these runs need repeating

The GPT-4o analyser reports cannot be used for this comparison:

- **Old code.** Five of the six reports have exactly the same composite score as before, and almost every criterion score is the same. They still quote the old cut-off points, such as "an incomplete header ('Digital India Land ')" and "Administrative attention was drawn to this". The app instance running with OpenAI was most likely still on the old code: it ran at the same time as the Gemini runs (09:21–09:42 UTC, which is 14:50–15:12 IST) but was probably started before the change.
- **Wrong files.** The Decision-Forcing, Gemini-written × GPT-4o analyser report (49.2%) scored the Lesson-Drawing GPT-4o-written case instead. It quotes that case's Hook, and its criterion scores match that report exactly.

### Defect found in a case

The Decision-Forcing GPT-4o-written case cuts off mid-sentence ("…the pace of digitisation was often dictated") just before the "Dilemma or Puzzle" heading. The Analyser reported this correctly. It is a Generator issue, so it goes to Step 2.

## Status

- `app.py` compiles.
- Tested with the Gemini analyser: works as intended.
- GPT-4o analyser: not yet tested on the new code.
- **Not committed.**

## Next

1. Restart the app instance that uses OpenAI, on the `analyser-truncation-chnage` branch.
2. Re-run the six GPT-4o analyser reports, using the correct files for the Decision-Forcing Gemini-written case.
3. Compare those reports with the Gemini analyser results above, criterion by criterion.
4. Use the full new score table as the starting point for Step 2 (Generator prompt changes).
