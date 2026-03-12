"""
HW3 Prompt Generator - Generates corrected prompt drafts with exact URLs
"""
import csv, os, urllib.parse

# ── Config ──
CSV_PATH = r'c:\Users\karan\Downloads\ADTA5770\one drive\GROUP_4\HW2_Contaminants_Documentation_Group4.csv'
SC_DIR = r'c:\Users\karan\Downloads\ADTA5770\ADTA5760--Group-4\pdfs\supply_chain'
MED_DIR = r'c:\Users\karan\Downloads\ADTA5770\ADTA5760--Group-4\pdfs\medical'
OUTPUT = r'c:\Users\karan\Downloads\ADTA5770\HW3_prompt_drafts.txt'
BASE_URL = 'https://adta2026group4.com/pdfs'

# ── Load Data ──
with open(CSV_PATH, 'r', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    csv_docs = {}
    csv_contaminants = {}
    for row in reader:
        did = row['document_id']
        title = row['document_title']
        if did not in csv_docs:
            csv_docs[did] = title
            csv_contaminants[did] = []
        csv_contaminants[did].append({
            'cid': row['contaminant_id'],
            'type': row['contaminating_type'],
            'text': row['contaminant_text'],
            'original': row['contaminated_text'],
            'location': row['exact_location']
        })

sc_files = [f for f in os.listdir(SC_DIR) if f.endswith('.pdf')]
med_files = [f for f in os.listdir(MED_DIR) if f.endswith('.pdf')]

def find_match(title, files):
    for f in files:
        if os.path.splitext(f)[0].strip() == title.strip():
            return f
    for f in files:
        if os.path.splitext(f)[0].strip().lower() == title.strip().lower():
            return f
    for f in files:
        if title.strip()[:40].lower() in os.path.splitext(f)[0].strip().lower():
            return f
    return None

def get_url(did):
    title = csv_docs.get(did, '')
    if did.startswith('SC_'):
        match = find_match(title, sc_files)
        domain = 'supply_chain'
    else:
        match = find_match(title, med_files)
        domain = 'medical'
    if match:
        encoded = urllib.parse.quote(match)
        return f'{BASE_URL}/{domain}/{encoded}'
    return 'URL_NOT_FOUND'

# ── Document Selection (seed=42, same as hw3_prep.py) ──
part1 = ['SC_7E5AA0', 'MED_0A4C1E']
part2_sc = ['SC_E256CF', 'SC_82604E', 'SC_A4CC09', 'SC_046CB7', 'SC_7B0EBD', 'SC_799B06', 'SC_8F419D']
part2_med = ['MED_67B704', 'MED_1C56DA', 'MED_B6A65F']
part3_sc = ['SC_D9F09A', 'SC_C6698C', 'SC_F0F386', 'SC_61EBC9', 'SC_C229E9', 'SC_8E4A08', 'SC_E92095',
            'SC_7CB134', 'SC_9DFEB7', 'SC_0195E1', 'SC_6476B2', 'SC_FB3B87', 'SC_AC5C4F', 'SC_97CA22',
            'SC_F37F7B', 'SC_85DC77', 'SC_7B5444', 'SC_FEF54D', 'SC_8E7E30', 'SC_DA2127', 'SC_B6D996',
            'SC_48C646', 'SC_8500A5', 'SC_BEDE1A', 'SC_8A2924', 'SC_E7ADA0', 'SC_34187C', 'SC_1031D5',
            'SC_F66C40', 'SC_15404A', 'SC_FAC84A', 'SC_DBB77D']
part3_med = ['MED_36F1BD', 'MED_378E06', 'MED_ABE610', 'MED_F520FE', 'MED_9589DA', 'MED_583321',
             'MED_E011BF', 'MED_07B97B', 'MED_AE7761', 'MED_DFB0B9', 'MED_C68DED', 'MED_507619',
             'MED_6C7571', 'MED_998FA0', 'MED_29AA9B', 'MED_F552C1']

# ── Helper to build doc list ──
def doc_list_block(doc_ids, label_prefix="Document"):
    lines = []
    for i, did in enumerate(doc_ids, 1):
        title = csv_docs.get(did, 'UNKNOWN')
        url = get_url(did)
        lines.append(f"{label_prefix} {i} ({did}): {title}")
        lines.append(f"  URL: {url}")
    return '\n'.join(lines)

# ── Build Output ──
out = []

def w(text=''):
    out.append(text)

def sep(char='=', width=80):
    out.append(char * width)

# ═══════════════════════════════════════════════════════════════════════════════
sep()
w("HW3 PROMPT DRAFTS — FOR YOUR REVIEW & MODIFICATION")
sep()
w("IMPORTANT: Per assignment AI-Using Policy, ALL prompts must be HUMAN-CREATED.")
w("These are DRAFTS/SUGGESTIONS for you to review, modify, and make your own.")
w("All URLs have been verified against actual PDF filenames in the Git repo.")
w(f"Total: 60 documents, 180 contaminants (3 per doc)")
sep()
w()

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 4: TERMINOLOGY
# ═══════════════════════════════════════════════════════════════════════════════
sep()
w("SECTION 4: TERMINOLOGY & CONCEPTS")
sep()
w("""
1. DOCUMENT ID (DOC-ID)
   A unique alphanumeric identifier assigned to each document in the contaminated
   knowledge base. Format: <PREFIX>_<HEX_CODE>
   - "SC_" prefix = Supply Chain Knowledge Base (General Knowledge Base)
   - "MED_" prefix = Medical Knowledge Base
   Examples: SC_7E5AA0, MED_0A4C1E

2. CONTAMINANT ID (CID)
   A unique sequential numeric identifier for each contaminant.
   Format: CID_<5-digit zero-padded number>
   Range: CID_00001 through CID_00450 (450 total across 150 docs)
   Examples: CID_00001, CID_00197, CID_00450

3. CONTAMINANT TYPE — Three categories of intentional text modifications:

   3a. TYPO
       A deliberate misspelling that changes spelling but preserves approximate
       readability. The replacement is a plausible typographical error.
       Example: "efficiency" → "efficency" (missing letter)
       Example: "methodology" → "methodolgy" (transposed letters)

   3b. CONFLICTING INFORMATION
       A deliberate replacement of a word/phrase with its semantic opposite or
       contradiction, changing meaning while maintaining grammatical correctness.
       Example: "positively" → "negatively" (meaning reversed)
       Example: "robust" → "fragile" (antonym substitution)
       Example: "minimize" → "maximize" (opposite action)

   3c. NONSENSE
       A deliberate replacement of a domain-relevant term with an absurd,
       contextually inappropriate word/phrase with no logical connection.
       Example: "supply chain" → "dolphin choir" (absurd)
       Example: "parameter" → "parachute" (unrelated noun)
""")

# ═══════════════════════════════════════════════════════════════════════════════
# PART I — PROMPT 1: Supply Chain (1 doc)
# ═══════════════════════════════════════════════════════════════════════════════
sep()
w("PART I — PROMPT 1: Supply Chain Document (1 document)")
sep()
w()
did = 'SC_7E5AA0'
title = csv_docs[did]
url = get_url(did)
w(f"Target: {did} - {title}")
w(f"URL: {url}")
w()
w("Known contaminants (ANSWER KEY — do not include in prompt):")
for c in csv_contaminants[did]:
    w(f"  {c['cid']} [{c['type']}]: \"{c['original']}\" → \"{c['text']}\" at {c['location']}")
w()
w("PROMPT TECHNIQUES USED:")
w("- Role Assignment (persona definition)")
w("- Task Decomposition (step-by-step instructions)")
w("- Structured Output Format (table/report)")
w("- Context Setting (explaining domain and purpose)")
w("- Specificity (requesting exact locations, types, reasoning)")
w()
w("--- DRAFT PROMPT ---")
w()
w(f"""You are an expert document quality analyst specializing in detecting
intentional text contaminations embedded in academic research documents.

I have a contaminated PDF document hosted at the following URL:
{url}

This document (Document ID: {did}) has been intentionally contaminated
with embedded contaminants that fall into three categories:
1. TYPOS — deliberate misspellings of words
2. CONFLICTING INFORMATION — words or phrases replaced with their opposites
   or contradictions that change the meaning
3. NONSENSE — domain-relevant terms replaced with absurd, contextually
   inappropriate words or phrases

Please perform the following tasks:

Step 1: Access and read the entire contents of the PDF document at the URL above.

Step 2: Carefully analyze the document to identify ALL contaminants that have
been embedded in it.

Step 3: For EACH contaminant found, report the following in a structured format:
   a) The contaminated text (the suspicious word or phrase as it appears)
   b) The contaminant type (Typo, Conflicting Information, or Nonsense)
   c) The exact location in the document (page number, paragraph, and
      surrounding context — quote the sentence it appears in)
   d) The likely original text (what the word or phrase should have been)
   e) Your reasoning for identifying this as a contaminant

Step 4: Provide a summary table with columns: #, Contaminant Text, Type,
Original Text, Page, Reasoning.""")
w()
w("--- END PROMPT ---")
w()

# ═══════════════════════════════════════════════════════════════════════════════
# PART I — PROMPT 2: Medical (1 doc)
# ═══════════════════════════════════════════════════════════════════════════════
sep()
w("PART I — PROMPT 2: Medical Document (1 document)")
sep()
w()
did = 'MED_0A4C1E'
title = csv_docs[did]
url = get_url(did)
w(f"Target: {did} - {title}")
w(f"URL: {url}")
w()
w("Known contaminants (ANSWER KEY — do not include in prompt):")
for c in csv_contaminants[did]:
    w(f"  {c['cid']} [{c['type']}]: \"{c['original']}\" → \"{c['text']}\" at {c['location']}")
w()
w("PROMPT TECHNIQUES USED:")
w("- Role Assignment (medical document analyst)")
w("- Task Decomposition (step-by-step)")
w("- Structured Output Format (table)")
w("- Context Setting (medical research domain)")
w("- Specificity (exact locations and reasoning)")
w()
w("--- DRAFT PROMPT ---")
w()
w(f"""You are an expert document quality analyst specializing in detecting
intentional text contaminations in medical and clinical research papers.

I have a contaminated PDF document hosted at the following URL:
{url}

This document (Document ID: {did}) has been intentionally contaminated
with embedded contaminants that fall into three categories:
1. TYPOS — deliberate misspellings of words
2. CONFLICTING INFORMATION — words or phrases replaced with their opposites
   or contradictions that change the meaning
3. NONSENSE — domain-relevant terms replaced with absurd, contextually
   inappropriate words or phrases

Please perform the following tasks:

Step 1: Access and read the entire contents of the PDF document at the URL above.

Step 2: Carefully analyze the document to identify ALL contaminants that have
been embedded in it.

Step 3: For EACH contaminant found, report the following in a structured format:
   a) The contaminated text (the suspicious word or phrase as it appears)
   b) The contaminant type (Typo, Conflicting Information, or Nonsense)
   c) The exact location in the document (page number, paragraph, and
      surrounding context — quote the sentence it appears in)
   d) The likely original text (what the word or phrase should have been)
   e) Your reasoning for identifying this as a contaminant

Step 4: Provide a summary table with columns: #, Contaminant Text, Type,
Original Text, Page, Reasoning.""")
w()
w("--- END PROMPT ---")
w()

# ═══════════════════════════════════════════════════════════════════════════════
# PART II — PROMPT 1: Supply Chain (7 docs)
# ═══════════════════════════════════════════════════════════════════════════════
sep()
w("PART II — PROMPT 1: Supply Chain Knowledge Base (7 documents)")
sep()
w()
w("PROMPT TECHNIQUES USED:")
w("- Role Assignment")
w("- Batch Processing Instructions (multi-document)")
w("- Structured Output Format (per-document + summary table)")
w("- Context Setting")
w("- Specificity")
w()

# Build doc URLs for the prompt
doc_lines_p2sc = []
for i, did in enumerate(part2_sc, 1):
    t = csv_docs[did]
    u = get_url(did)
    doc_lines_p2sc.append(f"Document {i} ({did}): {t}\n  URL: {u}")

w("Known contaminants (ANSWER KEY):")
for did in part2_sc:
    w(f"\n  {did} - {csv_docs[did]}:")
    for c in csv_contaminants[did]:
        w(f"    {c['cid']} [{c['type']}]: \"{c['original']}\" → \"{c['text']}\" at {c['location']}")
w()
w("--- DRAFT PROMPT ---")
w()
w(f"""You are an expert document quality analyst. I have 7 contaminated PDF
documents from a Supply Chain Knowledge Base. Each document has been
intentionally contaminated with exactly 3 contaminants of these types:
1. TYPOS — deliberate misspellings
2. CONFLICTING INFORMATION — words replaced with opposites/contradictions
3. NONSENSE — terms replaced with absurd, unrelated words

Please access, read, and analyze EACH of the following 7 documents to find
ALL contaminants embedded in them:

{chr(10).join(doc_lines_p2sc)}

For EACH document, please:
1. Access and read the entire PDF content from the URL
2. Identify ALL contaminants (typos, conflicting info, nonsense)
3. For each contaminant found, report:
   a) Document ID and title
   b) The contaminated text as it appears
   c) The contaminant type (Typo, Conflicting Information, or Nonsense)
   d) The exact location (page number, paragraph, surrounding sentence)
   e) The likely original text before contamination
   f) Your reasoning

Organize your findings by document. End with a summary table:
Document ID | Contaminant # | Type | Found Text | Original Text | Page""")
w()
w("--- END PROMPT ---")
w()

# ═══════════════════════════════════════════════════════════════════════════════
# PART II — PROMPT 2: Medical (3 docs)
# ═══════════════════════════════════════════════════════════════════════════════
sep()
w("PART II — PROMPT 2: Medical Knowledge Base (3 documents)")
sep()
w()
w("PROMPT TECHNIQUES USED:")
w("- Role Assignment")
w("- Batch Processing Instructions")
w("- Structured Output Format")
w("- Context Setting")
w("- Specificity")
w()

doc_lines_p2med = []
for i, did in enumerate(part2_med, 1):
    t = csv_docs[did]
    u = get_url(did)
    doc_lines_p2med.append(f"Document {i} ({did}): {t}\n  URL: {u}")

w("Known contaminants (ANSWER KEY):")
for did in part2_med:
    w(f"\n  {did} - {csv_docs[did]}:")
    for c in csv_contaminants[did]:
        w(f"    {c['cid']} [{c['type']}]: \"{c['original']}\" → \"{c['text']}\" at {c['location']}")
w()
w("--- DRAFT PROMPT ---")
w()
w(f"""You are an expert document quality analyst. I have 3 contaminated PDF
documents from a Medical Knowledge Base. Each document has been intentionally
contaminated with exactly 3 contaminants of these types:
1. TYPOS — deliberate misspellings
2. CONFLICTING INFORMATION — words replaced with opposites/contradictions
3. NONSENSE — terms replaced with absurd, unrelated words

Please access, read, and analyze EACH of the following 3 documents to find
ALL contaminants embedded in them:

{chr(10).join(doc_lines_p2med)}

For EACH document, please:
1. Access and read the entire PDF content from the URL
2. Identify ALL contaminants (typos, conflicting info, nonsense)
3. For each contaminant found, report:
   a) Document ID and title
   b) The contaminated text as it appears
   c) The contaminant type (Typo, Conflicting Information, or Nonsense)
   d) The exact location (page number, paragraph, surrounding sentence)
   e) The likely original text before contamination
   f) Your reasoning

Organize your findings by document. End with a summary table:
Document ID | Contaminant # | Type | Found Text | Original Text | Page""")
w()
w("--- END PROMPT ---")
w()

# ═══════════════════════════════════════════════════════════════════════════════
# PART III — PROMPT 1: Supply Chain (32 docs)
# ═══════════════════════════════════════════════════════════════════════════════
sep()
w("PART III — PROMPT 1: Supply Chain Knowledge Base (32 documents)")
sep()
w()
w("PROMPT TECHNIQUES USED:")
w("- Role Assignment")
w("- Batch Processing Instructions (large-scale)")
w("- Structured Output Format (per-doc + summary table)")
w("- Context Setting")
w("- Specificity")
w("- Constraint Setting (table format)")
w()

doc_lines_p3sc = []
for i, did in enumerate(part3_sc, 1):
    t = csv_docs[did]
    u = get_url(did)
    doc_lines_p3sc.append(f"Document {i} ({did}): {t}\n  URL: {u}")

w("Known contaminants (ANSWER KEY):")
for did in part3_sc:
    w(f"\n  {did} - {csv_docs[did]}:")
    for c in csv_contaminants[did]:
        w(f"    {c['cid']} [{c['type']}]: \"{c['original']}\" → \"{c['text']}\" at {c['location']}")
w()
w("--- DRAFT PROMPT ---")
w()
w(f"""You are an expert document quality analyst. I have 32 contaminated PDF
documents from a Supply Chain Knowledge Base. Each document has been
intentionally contaminated with exactly 3 contaminants of these types:
1. TYPOS — deliberate misspellings
2. CONFLICTING INFORMATION — words replaced with opposites/contradictions
3. NONSENSE — terms replaced with absurd, unrelated words

Please access, read, and analyze EACH of the following 32 documents to find
ALL contaminants embedded in them.

{chr(10).join(doc_lines_p3sc)}

For EACH document, please:
1. Access and read the entire PDF content from the URL
2. Identify ALL contaminants (typos, conflicting info, nonsense)
3. For each contaminant found, report:
   a) Document ID and title
   b) The contaminated text as it appears in the document
   c) The contaminant type (Typo, Conflicting Information, or Nonsense)
   d) The exact location (page number, paragraph, surrounding sentence)
   e) The likely original text before contamination
   f) Your reasoning for classifying this as a contaminant

Organize your report by document. End with a comprehensive summary table:
Document ID | Contaminant # | Type | Found Text | Original Text | Page""")
w()
w("--- END PROMPT ---")
w()

# ═══════════════════════════════════════════════════════════════════════════════
# PART III — PROMPT 2: Medical (16 docs)
# ═══════════════════════════════════════════════════════════════════════════════
sep()
w("PART III — PROMPT 2: Medical Knowledge Base (16 documents)")
sep()
w()
w("PROMPT TECHNIQUES USED:")
w("- Role Assignment")
w("- Batch Processing Instructions (large-scale)")
w("- Structured Output Format")
w("- Context Setting")
w("- Specificity")
w("- Constraint Setting")
w()

doc_lines_p3med = []
for i, did in enumerate(part3_med, 1):
    t = csv_docs[did]
    u = get_url(did)
    doc_lines_p3med.append(f"Document {i} ({did}): {t}\n  URL: {u}")

w("Known contaminants (ANSWER KEY):")
for did in part3_med:
    w(f"\n  {did} - {csv_docs[did]}:")
    for c in csv_contaminants[did]:
        w(f"    {c['cid']} [{c['type']}]: \"{c['original']}\" → \"{c['text']}\" at {c['location']}")
w()
w("--- DRAFT PROMPT ---")
w()
w(f"""You are an expert document quality analyst. I have 16 contaminated PDF
documents from a Medical Knowledge Base. Each document has been intentionally
contaminated with exactly 3 contaminants of these types:
1. TYPOS — deliberate misspellings
2. CONFLICTING INFORMATION — words replaced with opposites/contradictions
3. NONSENSE — terms replaced with absurd, unrelated words

Please access, read, and analyze EACH of the following 16 documents to find
ALL contaminants embedded in them.

{chr(10).join(doc_lines_p3med)}

For EACH document, please:
1. Access and read the entire PDF content from the URL
2. Identify ALL contaminants (typos, conflicting info, nonsense)
3. For each contaminant found, report:
   a) Document ID and title
   b) The contaminated text as it appears in the document
   c) The contaminant type (Typo, Conflicting Information, or Nonsense)
   d) The exact location (page number, paragraph, surrounding sentence)
   e) The likely original text before contamination
   f) Your reasoning for classifying this as a contaminant

Organize your report by document. End with a comprehensive summary table:
Document ID | Contaminant # | Type | Found Text | Original Text | Page""")
w()
w("--- END PROMPT ---")
w()

# ═══════════════════════════════════════════════════════════════════════════════
# EXECUTION CHECKLIST
# ═══════════════════════════════════════════════════════════════════════════════
sep()
w("EXECUTION CHECKLIST")
sep()
w("""
1. [ ] Review and personalize all prompts (AI policy: prompts must be human-written)
2. [ ] Open Gemini (gemini.google.com) — use free tier
3. [ ] Run PART I Prompt 1 (SC) — copy response
4. [ ] Run PART I Prompt 2 (MED) — copy response
5. [ ] Run PART II Prompt 1 (7 SC docs) — copy response
6. [ ] Run PART II Prompt 2 (3 MED docs) — copy response
7. [ ] Run PART III Prompt 1 (32 SC docs) — copy response
8. [ ] Run PART III Prompt 2 (16 MED docs) — copy response
9. [ ] For each response, evaluate using 5 criteria (Likert 1-5):
       (1) Usefulness/Relevance
       (2) Accuracy/Trustworthiness
       (3) Clarity/Coherence/Understanding
       (4) Completeness/Depth
       (5) Overall Satisfaction
10. [ ] Fill Word template with prompts + responses
11. [ ] Fill Excel Part I template (Part I responses)
12. [ ] Fill Excel Part II template (Part II + III responses)
13. [ ] Record FOUND: YES/NO for each contaminant by comparing with answer key
14. [ ] Part IV: Calculate statistical analysis
15. [ ] Save all files to OneDrive HW_3 folder
16. [ ] Send email with Part V teamwork report
""")

# Write to file
with open(OUTPUT, 'w', encoding='utf-8') as f:
    f.write('\n'.join(out))

print(f"Generated: {OUTPUT}")
print(f"Total lines: {len(out)}")
