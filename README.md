# LLM Knowledge Contamination Evaluation

> **Course:** ADTA 5770 · University of North Texas  
> **Contaminated Knowledge Base:** [adta-5760-group-4.vercel.app](https://adta-5760-group-4.vercel.app)

## Overview

Can Large Language Models detect deliberately planted misinformation in academic papers?

This project builds a **contamination-detection evaluation pipeline** for Google Gemini. We injected **450 contaminants** (factual errors, fabricated citations, statistical manipulation) into **150 academic PDFs** spanning supply chain management and medical research, then systematically prompted Gemini through escalating difficulty tiers to measure its detection accuracy.

## Evaluation Pipeline

| Stage | Description | Documents |
|-------|-------------|-----------|
| **Part I: Single-Doc Baseline** | One supply-chain and one medical PDF, full contaminant list provided | 2 |
| **Part II: Multi-Doc Batch** | 7 SC + 3 MED papers, evaluate batch reasoning | 10 |
| **Part III: Scaled Evaluation** | Remaining corpus in batched prompts (5A/5B/6A/6B) | 138 |
| **Part IV: Response Scoring & Analysis** | Statistical evaluation of Gemini's detection performance | n/a |

## Key Results

- Gemini's detection performance **degrades significantly** as batch size increases  
- **Fabricated citations** are the easiest contamination type to detect  
- **Statistical manipulations** are the hardest, and are often missed entirely  
- Single-document prompts achieve the highest precision  

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

## Team (Group 4)

Karan Parekh · Sanjana PR · Sana Mhapsekar · Medina Maloku
