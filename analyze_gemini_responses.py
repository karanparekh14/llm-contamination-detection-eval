"""
Analyze Gemini responses against the known answer key from CSV.
Produces a detailed match report for each prompt.
"""
import csv

# Load answer key from CSV
csv_path = r'c:\Users\karan\Downloads\ADTA5770\one drive\GROUP_4\HW2_Contaminants_Documentation_Group4.csv'
with open(csv_path, 'r', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    contaminants = {}
    for row in reader:
        did = row['document_id']
        if did not in contaminants:
            contaminants[did] = []
        contaminants[did].append({
            'cid': row['contaminant_id'],
            'type': row['contaminating_type'],
            'contaminated_text': row['contaminant_text'],      # what appears in doc
            'original_text': row['contaminated_text'],          # what it should have been
            'location': row['exact_location'],
            'explanation': row['explanation']
        })

# ── Define what Gemini found per prompt ──

# PROMPT 1: SC_7E5AA0
# Answer Key: efficency(Typo P7), negatively(Conflicting P15), dolphin choir(Nonsense P12)
prompt1_key = contaminants['SC_7E5AA0']
prompt1_found = [
    {'text': 'least', 'type': 'Conflicting', 'page': 1, 'original': 'most'},
    {'text': 'stuff', 'type': 'Nonsense', 'page': 2, 'original': 'materials/goods'},
    {'text': 'Object Relational Mapper', 'type': 'Nonsense', 'page': 4, 'original': 'Preferred Reporting Items...'},
    {'text': 'Chalenges', 'type': 'Typo', 'page': 7, 'original': 'Challenges'},
    {'text': 'Gupya', 'type': 'Typo', 'page': 7, 'original': 'Gupta'},
    {'text': 'does not', 'type': 'Conflicting', 'page': 8, 'original': 'does/is also associated'},
]

# PROMPT 2: MED_0A4C1E
# Answer Key: reccomend(Typo P1), fragile(Conflicting P7), parachute(Nonsense P5)
prompt2_found = [
    {'text': 'reccomend ations', 'type': 'Typo', 'page': 1, 'original': 'recommendations'},
    {'text': 'model parachute', 'type': 'Nonsense', 'page': 5, 'original': 'model parameters'},
    {'text': 'fragile ness', 'type': 'Conflicting', 'page': 6, 'original': 'robustness'},
]

# PROMPT 3 (7 SC docs)
prompt3_docs = ['SC_E256CF', 'SC_82604E', 'SC_A4CC09', 'SC_046CB7', 'SC_7B0EBD', 'SC_799B06', 'SC_8F419D']
prompt3_found = {
    'SC_E256CF': [
        {'text': 'dolphin choir', 'type': 'Nonsense', 'page': 2, 'original': 'Supply Chain'},
        {'text': 'significent', 'type': 'Typo', 'page': 14, 'original': 'significant'},
        {'text': 'criticized', 'type': 'Conflicting', 'page': 21, 'original': 'critical'},
    ],
    'SC_82604E': [
        {'text': 'dolphi hoir', 'type': 'Nonsense', 'page': 5, 'original': 'Supply Chain'},
        {'text': 'comptetitor', 'type': 'Typo', 'page': 36, 'original': 'competitor'},
        {'text': 'expansion', 'type': 'Conflicting', 'page': 151, 'original': 'reduction'},
    ],
    'SC_A4CC09': [
        {'text': 'dolphin choirresilience', 'type': 'Nonsense', 'page': 3, 'original': 'supply chain resilience'},
        {'text': 'distruption', 'type': 'Typo', 'page': 15, 'original': 'disruption'},
        {'text': 'disadvantage', 'type': 'Conflicting', 'page': 24, 'original': 'advantages'},
    ],
    'SC_046CB7': [
        {'text': 'threshhol', 'type': 'Typo', 'page': 1, 'original': 'threshold'},
        {'text': 'dolphin choir', 'type': 'Nonsense', 'page': 5, 'original': 'supply chain'},
        {'text': 'disadvantage', 'type': 'Conflicting', 'page': 6, 'original': 'advantage'},
    ],
    'SC_7B0EBD': [
        {'text': 'aquisition', 'type': 'Typo', 'page': 40, 'original': 'acquisition'},
        {'text': 'dolphin choir', 'type': 'Nonsense', 'page': 13, 'original': 'Supply Chain'},
        {'text': 'decreases', 'type': 'Conflicting', 'page': 162, 'original': 'increases'},
    ],
    'SC_799B06': [
        {'text': 'cumplier sades of condunt', 'type': 'Typo', 'page': 5, 'original': 'supplier codes of conduct'},
        {'text': 'expansion', 'type': 'Conflicting', 'page': 14, 'original': 'reduction'},
        {'text': 'dolphin choir', 'type': 'Nonsense', 'page': 11, 'original': 'Supply Chain'},
    ],
    'SC_8F419D': [
        {'text': 'dolphin choir', 'type': 'Nonsense', 'page': 243, 'original': 'supply chain'},
        {'text': 'privelege', 'type': 'Typo', 'page': 'x', 'original': 'privilege'},
        {'text': 'degradation', 'type': 'Conflicting', 'page': 9, 'original': 'enhancement'},
    ],
}

# PROMPT 4 (3 MED docs)
prompt4_docs = ['MED_67B704', 'MED_1C56DA', 'MED_B6A65F']
prompt4_found = {
    'MED_67B704': [
        {'text': 'penguinare', 'type': 'Nonsense', 'page': 5, 'original': 'patients are'},
        {'text': 'negatively/positively swap', 'type': 'Conflicting', 'page': 2, 'original': 'direction swap'},
        {'text': 'labratory', 'type': 'Typo', 'page': 7, 'original': 'laboratory'},
    ],
    'MED_1C56DA': [
        {'text': 'unfavorablutcomes', 'type': 'Conflicting', 'page': 1, 'original': 'favorable outcomes'},
        {'text': 'snorkeling', 'type': 'Nonsense', 'page': 2, 'original': 'treatments'},
        {'text': 'AC no lege MENTS', 'type': 'Typo', 'page': 3, 'original': 'ACKNOWLEDGEMENTS'},
    ],
    'MED_B6A65F': [
        {'text': 'fishbowl', 'type': 'Nonsense', 'page': 5, 'original': 'method/approach'},
        {'text': 'inter-Hemisphere', 'type': 'Conflicting', 'page': 16, 'original': 'intra-Hemisphere'},
        {'text': 'Ac ments', 'type': 'Typo', 'page': 17, 'original': 'Acknowledgments'},
    ],
}

# ── Compare function ──
def compare_doc(doc_id, gemini_found, answer_key):
    """Compare Gemini findings vs answer key for one document."""
    results = []
    for ak in answer_key:
        ak_text = ak['contaminated_text'].lower()
        ak_orig = ak['original_text'].lower()
        ak_type = ak['type'].lower()
        ak_page = ak['location'].split(',')[0].replace('Page ', '').strip()
        
        matched = False
        match_detail = ''
        
        for gf in gemini_found:
            gf_text = str(gf['text']).lower()
            gf_orig = str(gf['original']).lower()
            gf_type = gf['type'].lower()
            
            # Check text match (substring matching)
            text_match = (ak_text in gf_text or gf_text in ak_text or 
                         ak_orig in gf_orig or gf_orig in ak_orig or
                         ak_text in gf_orig or gf_text in ak_orig)
            
            # Check type match
            type_match = (ak_type[:4] in gf_type[:4] or gf_type[:4] in ak_type[:4])
            
            if text_match and type_match:
                matched = True
                match_detail = f"Gemini: '{gf['text']}' (P{gf['page']})"
                break
            elif text_match:
                matched = True
                match_detail = f"Gemini: '{gf['text']}' (P{gf['page']}) [type mismatch: Gemini={gf['type']}, Key={ak['type']}]"
                break
        
        results.append({
            'cid': ak['cid'],
            'type': ak['type'],
            'key_text': ak['contaminated_text'],
            'key_original': ak['original_text'],
            'key_page': ak_page,
            'found': 'YES' if matched else 'NO',
            'detail': match_detail if matched else 'NOT FOUND by Gemini'
        })
    
    return results

# ── Run Analysis ──
output = []
def w(s=''):
    output.append(s)

w('=' * 90)
w('GEMINI RESPONSE ANALYSIS — Comparison Against Known Answer Key')
w('=' * 90)
w()

# === PROMPT 1 ===
w('PROMPT 1: SC_7E5AA0 (Part I - Supply Chain, 1 doc)')
w('-' * 60)
w(f'Gemini found: 6 contaminants')
w(f'Answer key has: 3 contaminants')
w()
results_p1 = compare_doc('SC_7E5AA0', prompt1_found, contaminants['SC_7E5AA0'])
for r in results_p1:
    w(f"  {r['cid']} [{r['type']}]: \"{r['key_original']}\" → \"{r['key_text']}\" (Key: P{r['key_page']})")
    w(f"    FOUND: {r['found']} — {r['detail']}")
w()
matched_p1 = sum(1 for r in results_p1 if r['found'] == 'YES')
w(f'  SCORE: {matched_p1}/3 of our embedded contaminants correctly identified')
w(f'  NOTE: Gemini found 6 items but {3-matched_p1} of our 3 were missed.')
w(f'  Gemini found {6-matched_p1} items NOT in our answer key (false positives or pre-existing issues)')
w()

# === PROMPT 2 ===
w('PROMPT 2: MED_0A4C1E (Part I - Medical, 1 doc)')
w('-' * 60)
w(f'Gemini found: 3 contaminants')
w(f'Answer key has: 3 contaminants')
w()
results_p2 = compare_doc('MED_0A4C1E', prompt2_found, contaminants['MED_0A4C1E'])
for r in results_p2:
    w(f"  {r['cid']} [{r['type']}]: \"{r['key_original']}\" → \"{r['key_text']}\" (Key: P{r['key_page']})")
    w(f"    FOUND: {r['found']} — {r['detail']}")
w()
matched_p2 = sum(1 for r in results_p2 if r['found'] == 'YES')
w(f'  SCORE: {matched_p2}/3 correctly identified')
w()

# === PROMPT 3 ===
w('PROMPT 3: 7 SC Documents (Part II - Supply Chain)')
w('-' * 60)
total_p3 = 0
matched_p3 = 0
for doc_id in prompt3_docs:
    ak = contaminants[doc_id]
    gf = prompt3_found[doc_id]
    results = compare_doc(doc_id, gf, ak)
    m = sum(1 for r in results if r['found'] == 'YES')
    total_p3 += 3
    matched_p3 += m
    w(f'  {doc_id}: {m}/3 matched')
    for r in results:
        marker = '✓' if r['found'] == 'YES' else '✗'
        w(f"    {marker} {r['cid']} [{r['type']}] \"{r['key_text']}\" — {r['found']}: {r['detail']}")
    w()
w(f'  TOTAL SCORE: {matched_p3}/{total_p3}')
w()

# === PROMPT 4 ===
w('PROMPT 4: 3 MED Documents (Part II - Medical)')
w('-' * 60)
total_p4 = 0
matched_p4 = 0
for doc_id in prompt4_docs:
    ak = contaminants[doc_id]
    gf = prompt4_found[doc_id]
    results = compare_doc(doc_id, gf, ak)
    m = sum(1 for r in results if r['found'] == 'YES')
    total_p4 += 3
    matched_p4 += m
    w(f'  {doc_id}: {m}/3 matched')
    for r in results:
        marker = '✓' if r['found'] == 'YES' else '✗'
        w(f"    {marker} {r['cid']} [{r['type']}] \"{r['key_text']}\" — {r['found']}: {r['detail']}")
    w()
w(f'  TOTAL SCORE: {matched_p4}/{total_p4}')
w()

# === PROMPTS 5 & 6 ===
w('PROMPT 5: 32 SC Documents (Part III - Supply Chain)')
w('-' * 60)
w('  STATUS: FAILED — Gemini refused to access the URLs')
w('  FOUND: 0/96 contaminants')
w()
w('PROMPT 6: 16 MED Documents (Part III - Medical)')
w('-' * 60)
w('  STATUS: FAILED — Gemini refused to access the URLs')
w('  FOUND: 0/48 contaminants')
w()

# === OVERALL SUMMARY ===
w('=' * 90)
w('OVERALL SUMMARY')
w('=' * 90)
total_contaminants = 3 + 3 + 21 + 9 + 96 + 48  # 180
total_found = matched_p1 + matched_p2 + matched_p3 + matched_p4
w(f'Prompt 1 (Part I SC):    {matched_p1}/3   contaminants matched')
w(f'Prompt 2 (Part I MED):   {matched_p2}/3   contaminants matched')
w(f'Prompt 3 (Part II SC):   {matched_p3}/21  contaminants matched')
w(f'Prompt 4 (Part II MED):  {matched_p4}/9   contaminants matched')
w(f'Prompt 5 (Part III SC):  0/96  FAILED (Gemini refused)')
w(f'Prompt 6 (Part III MED): 0/48  FAILED (Gemini refused)')
w(f'')
w(f'TOTAL: {total_found}/{total_contaminants} contaminants correctly identified ({total_found/total_contaminants*100:.1f}%)')
w(f'')
w(f'CRITICAL ISSUES:')
w(f'1. Prompts 5 & 6 (80% of all contaminants) completely FAILED — Gemini refused URL access')
w(f'2. Prompt 1 found extra contaminants not in our key (may be real issues OR hallucinations)')
w(f'3. Some type/page mismatches in Prompts 3 & 4 (Gemini found the right text but wrong page/type)')

# Write output
out_path = r'c:\Users\karan\Downloads\ADTA5770\gemini_analysis_results.txt'
with open(out_path, 'w', encoding='utf-8') as f:
    f.write('\n'.join(output))

print('\n'.join(output))
