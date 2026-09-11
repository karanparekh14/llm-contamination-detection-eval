# LLM Knowledge Contamination Evaluation

> **Paper:** *When Auditors Fabricate: Batch-Size Degradation and Confident Hallucination in LLM Detection of Planted Document Contamination*  
> **DOI (concept, always resolves to the newest version):** [10.5281/zenodo.21939087](https://doi.org/10.5281/zenodo.21939087)  
> **arXiv:** [arXiv:2609.09696](https://arxiv.org/abs/2609.09696) (cs.CL, cross list cs.AI)  
> **Course:** ADTA 5770 · University of North Texas  
> **Contaminated Knowledge Base:** [adta-5760-group-4.vercel.app](https://adta-5760-group-4.vercel.app)

## Overview

Can Large Language Models detect deliberately planted misinformation in academic papers?

This project builds a **contamination-detection evaluation pipeline** for Google Gemini 3.0 Pro. We injected **450 contaminants** of three types (typographical corruption, semantic reversal, and absurd out-of-context insertion) into **150 academic PDFs** spanning supply chain management and medical research, then systematically prompted the model through escalating batch sizes to measure its detection accuracy. Responses are scored against a **180-contaminant answer key covering 60 documents**.

## Evaluation Pipeline

| Stage | Description | Documents prompted | Documents scored |
|-------|-------------|--------------------|------------------|
| **Part I: Single-Doc Baseline** | One supply-chain and one medical PDF, full contaminant list provided | 2 | 2 |
| **Part II: Multi-Doc Batch** | 7 SC + 3 MED papers, evaluate batch reasoning | 10 | 10 |
| **Part III: Scaled Evaluation** | Remaining corpus in batched prompts (5A/5B/6A/6B) | 138 | 48 |
| **Part IV: Response Scoring & Analysis** | Statistical evaluation of Gemini's detection performance | n/a | n/a |

## Key Results

- **Detection is unreliable even at small scale and collapses entirely at large scale.** Recovery was 50% on single documents (3/6) and 60% on small batches (18/30), a difference this sample cannot resolve (two-proportion z = 0.45, p = 0.65), against **2.8% on large batches (4/144)**, where the collapse is a measured effect (z = 8.58, p below 1e-17). Reported as 95% Wilson intervals: Part I 18.8% to 81.2%, Part II 42.3% to 75.4%, Part III 1.1% to 6.9%.
- **The failure mode at scale is fabrication, not abstention.** Rather than reporting incomplete processing, the model returned confident findings that included contaminants it invented itself, absurdities such as "telepathic squirrel" and "quantum-powered toaster" that appear in no document.
- **Plausible corruptions are missed most often.** Within completed evaluations, absurd insertions were recovered at 75% (9/12), while semantic reversals and typographical corruptions were each recovered at only 50% (6/12). The contamination types most likely to occur in the wild are the ones the model catches least.
- **One corpus property qualifies that asymmetry.** The absurd insertion was not varied uniformly: the string "dolphin choir" accounts for 37 of the 60 nonsense contaminants in the scored set, so those items are not independent within a batch. Separating the repeated string leaves the rate unchanged (6/8 against 3/4), but the varied subset is too small to be reassuring.
- **Injection versus reference is an open confound, not a result.** Every regime here referenced documents by URL, so this design cannot separate context overload from silent retrieval failure.
- **Small, bounded prompts are the only reliable regime**, and every reported finding still needs mechanical verification against the source text.

## Repository Structure

```
├── hw3_generate_prompts.py          # Generates structured prompts with PDF URLs
├── analyze_all_gemini_responses.py  # Scores all Gemini responses against answer key
├── analyze_gemini_responses.py      # Per-prompt response analysis
├── hw3_part_iv_analysis.py          # EDA, charts & statistical summary for Part IV
├── fill_hw3_templates.py            # Auto-fills evaluation score templates
├── hw3_prep.py                      # Document selection & preprocessing
├── generate_results.py              # Aggregates final results
├── gemini_analysis_results.txt      # Summary output
├── gemini_full_analysis_results.txt # Full detection log
├── prompts_list.txt                 # All generated prompts
└── HW3_prep_document_selection.txt  # Document selection rationale
```

## A note on the Likert scores in this repository

`fill_hw3_templates.py` contains `get_likert_scores`, and `hw3_part_iv_analysis.py` plots its output. These values exist because the course assignment template required a five-criteria self-evaluation. They are **assigned by rule from whether a contaminant was found**, not collected as independent human or model ratings, so any pattern in them is a property of the scoring function rather than an observation about the model.

**No result in the paper rests on them.** Every number reported in the paper comes from programmatic scoring against the answer key, with manual adjudication of page and type mismatches. A claim that did rest on these ratings was removed in version 3 of the paper for exactly this reason. The code is kept here rather than deleted so that the published record and the repository agree.

## Tech Stack

`Python` · `Google Gemini API` · `openpyxl` · `matplotlib` · `python-docx`

## How to Run

```bash
# 1. Generate prompts from the contamination CSV
python hw3_generate_prompts.py

# 2. Analyze Gemini responses against the answer key
python analyze_all_gemini_responses.py

# 3. Generate Part IV statistical analysis & charts
python hw3_part_iv_analysis.py
```

> **Note:** The contaminated PDFs and source CSV are hosted separately in the [genai-rag-qa-portfolio](https://github.com/karanparekh14/genai-rag-qa-portfolio) repository.

## Citation

Parekh, K., Pendyala Ravinder, S., Mhapsekar, S., & Maloku, M. (2026). *When Auditors Fabricate: Batch-Size Degradation and Confident Hallucination in LLM Detection of Planted Document Contamination*. Zenodo. https://doi.org/10.5281/zenodo.21939087

Also available as arXiv:2609.09696.

## Team (Group 4)

Karan Parekh · Sanjana PR · Sana Mhapsekar · Medina Maloku
