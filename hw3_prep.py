import csv, random

random.seed(42)

# Read CSV
with open(r'c:\Users\karan\Downloads\ADTA5770\one drive\GROUP_4\HW2_Contaminants_Documentation_Group4.csv', 'r', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    rows = list(reader)

# Build doc -> contaminants map
doc_map = {}
for row in rows:
    did = row['document_id']
    if did not in doc_map:
        doc_map[did] = {'title': row['document_title'], 'contaminants': []}
    doc_map[did]['contaminants'].append({
        'contaminant_id': row['contaminant_id'],
        'contaminant_text': row['contaminant_text'],
        'contaminating_type': row['contaminating_type'],
        'contaminated_text': row['contaminated_text'],
        'exact_location': row['exact_location'],
        'explanation': row['explanation']
    })

# Separate SC and MED
sc_ids = [d for d in doc_map if d.startswith('SC_')]
med_ids = [d for d in doc_map if d.startswith('MED_')]
random.shuffle(sc_ids)
random.shuffle(med_ids)

# PART I: 1 SC + 1 MED
p1_sc = [sc_ids[0]]
p1_med = [med_ids[0]]

# PART II (Word template): 7 SC + 3 MED
p2_sc = sc_ids[1:8]
p2_med = med_ids[1:4]

# PART III (Word template): 32 SC + 16 MED
p3_sc = sc_ids[8:40]
p3_med = med_ids[4:20]

# Build comprehensive prep file
output = []
output.append("=" * 80)
output.append("HW3 PREPARATION - DOCUMENT SELECTION & CONTAMINANT REFERENCE")
output.append("Group 4 | Generated with seed=42 | ADTA-DAST 5770")
output.append("=" * 80)

all_selected = p1_sc + p1_med + p2_sc + p2_med + p3_sc + p3_med
output.append(f"\nTotal documents selected: {len(all_selected)}")
output.append(f"  Part I:   {len(p1_sc)} SC + {len(p1_med)} MED = {len(p1_sc)+len(p1_med)}")
output.append(f"  Part II:  {len(p2_sc)} SC + {len(p2_med)} MED = {len(p2_sc)+len(p2_med)}")
output.append(f"  Part III: {len(p3_sc)} SC + {len(p3_med)} MED = {len(p3_sc)+len(p3_med)}")
output.append(f"  TOTAL: {len(all_selected)} unique documents")
remaining_sc = len(sc_ids) - len(p1_sc) - len(p2_sc) - len(p3_sc)
remaining_med = len(med_ids) - len(p1_med) - len(p2_med) - len(p3_med)
output.append(f"  Remaining unused: {remaining_sc} SC + {remaining_med} MED")

for section_name, sc_list, med_list in [
    ("PART I (2 documents: 1 SC + 1 MED)", p1_sc, p1_med),
    ("PART II (10 documents: 7 SC + 3 MED)", p2_sc, p2_med),
    ("PART III (48 documents: 32 SC + 16 MED)", p3_sc, p3_med),
]:
    output.append(f"\n{'='*80}")
    output.append(f"  {section_name}")
    output.append(f"{'='*80}")
    
    for kb_name, ids in [("Supply Chain", sc_list), ("Medical", med_list)]:
        output.append(f"\n  --- {kb_name} Documents ---")
        for i, did in enumerate(ids, 1):
            info = doc_map[did]
            folder = 'supply_chain' if did.startswith('SC_') else 'medical'
            output.append(f"\n  [{i}] {did}: {info['title']}")
            output.append(f"      URL: https://adta2026group4.com/pdfs/{folder}/{info['title']}.pdf")
            output.append(f"      Contaminants ({len(info['contaminants'])}):")
            for c in info['contaminants']:
                output.append(f"        - {c['contaminant_id']} [{c['contaminating_type']}]: \"{c['contaminated_text']}\" -> \"{c['contaminant_text']}\"")
                output.append(f"          Location: {c['exact_location']}")
                output.append(f"          Explanation: {c['explanation']}")

text = '\n'.join(output)
with open(r'c:\Users\karan\Downloads\ADTA5770\HW3_prep_document_selection.txt', 'w', encoding='utf-8') as f:
    f.write(text)

print(f"Prep file written: {len(text)} chars")
print(f"\nQuick summary:")
print(f"Part I:   {p1_sc[0]} + {p1_med[0]}")
print(f"Part II:  {len(p2_sc)} SC + {len(p2_med)} MED = {len(p2_sc)+len(p2_med)}")
print(f"Part III: {len(p3_sc)} SC + {len(p3_med)} MED = {len(p3_sc)+len(p3_med)}")
