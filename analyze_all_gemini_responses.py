"""
Comprehensive Analysis: ALL Gemini Responses vs Answer Key
Covers Prompts 1-4 (original) + 5A-5D + 6A-6B (split batches)
"""
import csv
import os

# ============================================================
# LOAD ANSWER KEY FROM CSV
# ============================================================
csv_path = r"one drive\GROUP_4\HW2_Contaminants_Documentation_Group4.csv"
answer_key = {}  # doc_id -> list of {cid, type, contaminant_text, original_text, page_coords}

with open(csv_path, 'r', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    for row in reader:
        doc_id = row.get('document_id', '').strip()
        if not doc_id:
            continue
        if doc_id not in answer_key:
            answer_key[doc_id] = []
        answer_key[doc_id].append({
            'cid': row.get('contaminant_id', '').strip(),
            'type': row.get('contaminating_type', '').strip(),
            'contaminant': row.get('contaminant_text', '').strip(),
            'original': row.get('contaminated_text', '').strip(),  # what was replaced
        })

# ============================================================
# DEFINE ALL DOCUMENT SETS PER PROMPT
# ============================================================

# Which docs belong to each prompt
prompt_docs = {
    'Prompt 1 (Part I SC)': ['SC_7E5AA0'],
    'Prompt 2 (Part I MED)': ['MED_0A4C1E'],
    'Prompt 3 (Part II SC - 7 docs)': ['SC_E256CF', 'SC_82604E', 'SC_A4CC09', 'SC_046CB7', 'SC_7B0EBD', 'SC_799B06', 'SC_8F419D'],
    'Prompt 4 (Part II MED - 3 docs)': ['MED_67B704', 'MED_1C56DA', 'MED_B6A65F'],
    'Prompt 5A (Part III SC batch 1)': ['SC_D9F09A', 'SC_C6698C', 'SC_F0F386', 'SC_61EBC9', 'SC_C229E9', 'SC_8E4A08', 'SC_E92095', 'SC_7CB134'],
    'Prompt 5B (Part III SC batch 2)': ['SC_9DFEB7', 'SC_0195E1', 'SC_6476B2', 'SC_FB3B87', 'SC_AC5C4F', 'SC_97CA22', 'SC_F37F7B', 'SC_85DC77'],
    'Prompt 5C (Part III SC batch 3)': ['SC_7B5444', 'SC_FEF54D', 'SC_8E7E30', 'SC_DA2127', 'SC_B6D996', 'SC_48C646', 'SC_8500A5', 'SC_BEDE1A'],
    'Prompt 5D (Part III SC batch 4)': ['SC_8A2924', 'SC_E7ADA0', 'SC_34187C', 'SC_1031D5', 'SC_F66C40', 'SC_15404A', 'SC_FAC84A', 'SC_DBB77D'],
    'Prompt 6A (Part III MED batch 1)': ['MED_36F1BD', 'MED_378E06', 'MED_ABE610', 'MED_F520FE', 'MED_9589DA', 'MED_583321', 'MED_E011BF', 'MED_07B97B'],
    'Prompt 6B (Part III MED batch 2)': ['MED_AE7761', 'MED_DFB0B9', 'MED_C68DED', 'MED_507619', 'MED_6C7571', 'MED_998FA0', 'MED_29AA9B', 'MED_F552C1'],
}

# ============================================================
# WHAT GEMINI ACTUALLY REPORTED FOR EACH DOC
# Format: doc_id -> list of (found_text, guessed_type)
# Types: 'Typo', 'Conflicting', 'Nonsense'
# ============================================================

gemini_findings = {}

# --- PROMPT 1: SC_7E5AA0 ---
gemini_findings['SC_7E5AA0'] = [
    ('least', 'Conflicting'),           # P1
    ('stuff', 'Nonsense'),              # P2
    ('Object Relational Mapper', 'Nonsense'),  # P4
    ('Chalenges', 'Typo'),              # P7
    ('Gupya', 'Typo'),                  # P7
    ('does not', 'Conflicting'),        # P8
]

# --- PROMPT 2: MED_0A4C1E ---
gemini_findings['MED_0A4C1E'] = [
    ('reccomend ations', 'Typo'),       # P1
    ('model parachute', 'Nonsense'),    # P5
    ('fragile ness', 'Conflicting'),    # P6
]

# --- PROMPT 3: 7 SC docs ---
gemini_findings['SC_E256CF'] = [
    ('dolphin choir', 'Nonsense'),      # P2
    ('significent', 'Typo'),            # P14
    ('criticized', 'Conflicting'),      # P21
]
gemini_findings['SC_82604E'] = [
    ('dolphi hoir', 'Nonsense'),        # P5
    ('comptetitor', 'Typo'),            # P36
    ('expansion', 'Conflicting'),       # P151
]
gemini_findings['SC_A4CC09'] = [
    ('dolphin choirresilience', 'Nonsense'),  # P3
    ('distruption', 'Typo'),            # P15
    ('disadvantage', 'Conflicting'),    # P24
]
gemini_findings['SC_046CB7'] = [
    ('threshhol', 'Typo'),             # P1
    ('dolphin choir', 'Nonsense'),      # P5
    ('disadvantage', 'Conflicting'),    # P6
]
gemini_findings['SC_7B0EBD'] = [
    ('aquisition', 'Typo'),            # P40
    ('dolphin choir', 'Nonsense'),      # P13
    ('decreases', 'Conflicting'),       # P162
]
gemini_findings['SC_799B06'] = [
    ('cumplier sades of condunt', 'Typo'),  # P5
    ('expansion', 'Conflicting'),       # P14
    ('dolphin choir', 'Nonsense'),      # P11
]
gemini_findings['SC_8F419D'] = [
    ('dolphin choir', 'Nonsense'),      # P243
    ('privelege', 'Typo'),             # Px
    ('degradation', 'Conflicting'),     # P9
]

# --- PROMPT 4: 3 MED docs ---
gemini_findings['MED_67B704'] = [
    ('penguinare', 'Nonsense'),         # P5
    ('negatively associated', 'Conflicting'),  # P2
    ('labratory', 'Typo'),             # P7
]
gemini_findings['MED_1C56DA'] = [
    ('unfavorablutcomes', 'Conflicting'),  # P1
    ('snorkeling', 'Nonsense'),         # P2
    ('AC no lege MENTS', 'Typo'),       # P3
]
gemini_findings['MED_B6A65F'] = [
    ('fishbowl', 'Nonsense'),          # P5
    ('inter-Hemisphere', 'Conflicting'),  # P16
    ('Ac ments', 'Typo'),              # P17
]

# --- PROMPT 5A: 8 SC docs (batch 1) ---
# Gemini gave responses but they look heavily hallucinated
gemini_findings['SC_D9F09A'] = [
    ('Al-based', 'Typo'),               # P10
    ('minimize the buzz', 'Conflicting'),  # P1
    ('embodiment in machines', 'Nonsense'),  # P2
]
gemini_findings['SC_C6698C'] = [
    ('articlelooks', 'Typo'),           # P1
    ('need help', 'Conflicting'),       # P1
    ('clients', 'Nonsense'),            # P1
]
gemini_findings['SC_F0F386'] = [
    ('L..G.', 'Typo'),                 # P1
    ('burning question', 'Conflicting'),  # P1
    ('conservative', 'Nonsense'),       # P8
]
gemini_findings['SC_61EBC9'] = [
    ('of Global Sourcing', 'Typo'),     # P1
    ('International Sourcing', 'Conflicting'),  # P1
    ('Social Sciences', 'Nonsense'),    # P1
]
gemini_findings['SC_C229E9'] = [
    ('in Procurement', 'Typo'),         # P3
    ('Restricted tender', 'Conflicting'),  # P57
    ('Best-...', 'Nonsense'),           # P176
]
gemini_findings['SC_8E4A08'] = [
    ('JCSM', 'Typo'),                  # P1
    ('mission has adapted', 'Conflicting'),  # P1
    ('economic system', 'Nonsense'),    # P1
]
gemini_findings['SC_E92095'] = [
    ('sul31810150', 'Typo'),            # P1
    ('stays neutral', 'Conflicting'),   # P1
    ('Digital Platform', 'Nonsense'),   # P1
]
gemini_findings['SC_7CB134'] = [
    ('RoopsingB', 'Typo'),             # P1
    ('After modification', 'Conflicting'),  # P1
    ('PERFORM', 'Nonsense'),            # P1
]

# --- PROMPT 5B: 8 SC docs (batch 2) ---
gemini_findings['SC_9DFEB7'] = [
    ('Center of astrology', 'Nonsense'),  # P2
    ('Confidentialy', 'Typo'),          # P2
    ('competitive disadvantage', 'Conflicting'),  # P9
]
gemini_findings['SC_0195E1'] = [
    ('sells more cars', 'Nonsense'),    # P253
    ('Expactations', 'Typo'),           # P9
    ('increase the time', 'Conflicting'),  # P138
]
gemini_findings['SC_6476B2'] = [
    ('enchanted crystal balls', 'Nonsense'),  # Abstract
    ('Enviromental', 'Typo'),           # Title
    ('Increasing carbon emissions', 'Conflicting'),  # Intro
]
gemini_findings['SC_FB3B87'] = [
    ('Perceptive Craving Game', 'Nonsense'),  # Title
    ('Confidentialy', 'Typo'),          # Security section
    ('by hiding all data', 'Conflicting'),  # Framework
]
gemini_findings['SC_AC5C4F'] = [
    ('bid only with its second paddle', 'Nonsense'),  # P128
    ('CO', 'Typo'),                     # P236
    ('most inaccessible locations', 'Conflicting'),  # Section
]
gemini_findings['SC_97CA22'] = [
    ('carrier pigeons', 'Nonsense'),    # Methodology
    ('Complience', 'Typo'),             # Title
    ('maximize environmental degradation', 'Conflicting'),  # Conclusion
]
gemini_findings['SC_F37F7B'] = [
    ('squirrels in the warehouse', 'Nonsense'),  # P45
    ('Availabilty', 'Typo'),            # P5
    ('stockouts occur daily', 'Conflicting'),  # Case Study
]
gemini_findings['SC_85DC77'] = [
    ('underwater basket weaving', 'Nonsense'),  # Criteria
    ('Assessmet', 'Typo'),              # Title
    ('low-quality suppliers', 'Conflicting'),  # Results
]

# --- PROMPT 5C: 8 SC docs (batch 3) ---
gemini_findings['SC_7B5444'] = [
    ('emergence ace', 'Typo'),          # P1
    ('ignore trade-offs', 'Conflicting'),  # P3
    ('quantum-powered toaster', 'Nonsense'),  # P12
]
gemini_findings['SC_FEF54D'] = [
    ('comon', 'Typo'),                  # P1
    ('decrease visibility', 'Conflicting'),  # P12
    ('sentient marshmallow', 'Nonsense'),  # P18
]
gemini_findings['SC_8E7E30'] = [
    ('revview', 'Typo'),               # P1
    ('maximize disruptions', 'Conflicting'),  # P4
    ('astrological jellybeans', 'Nonsense'),  # P9
]
gemini_findings['SC_DA2127'] = [
    ('Kyoenggi', 'Typo'),              # P1
    ('vulnerable', 'Conflicting'),      # P2
    ('bicycle-riding octopus', 'Nonsense'),  # P7
]
gemini_findings['SC_B6D996'] = [
    ('emergence ace', 'Typo'),          # P1
    ('exclusion', 'Conflicting'),       # P1
    ('flying pancake', 'Nonsense'),     # P4
]
gemini_findings['SC_48C646'] = [
    ('increass', 'Typo'),              # P2
    ('increases during decrease', 'Conflicting'),  # P2
    ('disco-dancing warehouse', 'Nonsense'),  # P5
]
gemini_findings['SC_8500A5'] = [
    ('term', 'Typo'),                   # P1
    ('increase costs', 'Conflicting'),  # P2
    ('telepathic squirrel', 'Nonsense'),  # P15
]
gemini_findings['SC_BEDE1A'] = [
    ('integartion', 'Typo'),           # P1
    ('surplus of stability', 'Conflicting'),  # P1
    ('invisible unicycles', 'Nonsense'),  # P8
]

# --- PROMPT 5D: 8 SC docs (batch 4) ---
gemini_findings['SC_8A2924'] = [
    ('Decreasing visibility', 'Conflicting'),  # P12
    ('performence', 'Typo'),            # P4
    ('banana peel', 'Nonsense'),        # P25
]
gemini_findings['SC_E7ADA0'] = [
    ('Alexarider', 'Typo'),            # P1
    ('reduce trust', 'Conflicting'),    # P1
    ('Uberisation', 'Nonsense'),        # P9
]
gemini_findings['SC_34187C'] = [
    ('disruptive behaviour', 'Conflicting'),  # P9
    ('Holmstróm', 'Typo'),             # P6
    ('toaster', 'Nonsense'),            # P9
]
gemini_findings['SC_1031D5'] = [
    ('uninfluenced', 'Conflicting'),    # P1
    ('Decisiun', 'Typo'),              # P1
    ('Pogo-Sticking', 'Nonsense'),      # P1
]
gemini_findings['SC_F66C40'] = [
    ('boot strapping', 'Typo'),         # P14
    ('insensitive', 'Conflicting'),     # P2
    ('hash browns', 'Nonsense'),        # P2
]
gemini_findings['SC_15404A'] = [
    ('22% off', 'Typo'),               # P5
    ('higher costs', 'Conflicting'),    # P5
    ('level in the clouds', 'Nonsense'),  # P5
]
gemini_findings['SC_FAC84A'] = [
    ('Abramova¹"', 'Typo'),            # P1
    ('increased fragmentation', 'Conflicting'),  # P3
    ('magic wand regression', 'Nonsense'),  # P5
]
gemini_findings['SC_DBB77D'] = [
    ('highlands University', 'Typo'),   # P1
    ('ability to collapse', 'Conflicting'),  # P4
    ('dancing with wolves', 'Nonsense'),  # P7
]

# --- PROMPT 6A: 8 MED docs (batch 1) ---
gemini_findings['MED_36F1BD'] = [
    ('tuberculosiz', 'Typo'),           # P1
    ('ineffective', 'Conflicting'),     # P1
    ('marshmallow', 'Nonsense'),        # P3
]
gemini_findings['MED_378E06'] = [
    ('Equati0ns', 'Typo'),             # P1
    ('decreases', 'Conflicting'),       # P2
    ('disco dance', 'Nonsense'),        # P4
]
gemini_findings['MED_ABE610'] = [
    ('Kinaze', 'Typo'),                # P1
    ('inhibits', 'Conflicting'),        # P2
    ('bubblegum', 'Nonsense'),          # P5
]
gemini_findings['MED_F520FE'] = [
    ('Lipodystr0phy', 'Typo'),         # P1
    ('excess', 'Conflicting'),          # P1
    ('pancake', 'Nonsense'),            # P3
]
gemini_findings['MED_9589DA'] = [
    ('Ultras0und', 'Typo'),            # P1
    ('increased', 'Conflicting'),       # P4
    ('banana peel', 'Nonsense'),        # P2
]
gemini_findings['MED_583321'] = [
    ('Cor0navirus', 'Typo'),           # P1
    ('known', 'Conflicting'),           # P1
    ('magic carpet', 'Nonsense'),       # P5
]
gemini_findings['MED_E011BF'] = [
    ('treatmunt', 'Typo'),             # P1
    ('Ineffective', 'Conflicting'),     # P2
    ('kazoo', 'Nonsense'),              # P3
]
gemini_findings['MED_07B97B'] = [
    ('Onc0logy', 'Typo'),             # P1
    ('Simple', 'Conflicting'),          # P2
    ('pogo stick', 'Nonsense'),         # P4
]

# --- PROMPT 6B: 8 MED docs (batch 2) ---
gemini_findings['MED_AE7761'] = [
    ('synethesis', 'Typo'),             # P1
    ('increases oxidative stress', 'Conflicting'),  # P3
    ('marshmallow fluff injections', 'Nonsense'),  # P5
]
gemini_findings['MED_DFB0B9'] = [
    ('thoraccic', 'Typo'),             # P1
    ('contraindicated', 'Conflicting'),  # P2
    ('disco glitter', 'Nonsense'),      # P4
]
gemini_findings['MED_C68DED'] = [
    ('dependance', 'Typo'),            # P2
    ('non-addictive', 'Conflicting'),   # P4
    ('blue cheese popsicles', 'Nonsense'),  # P5
]
gemini_findings['MED_507619'] = [
    ('cancr', 'Typo'),                 # P1
    ('less effective', 'Conflicting'),  # P3
    ('sourdough starter', 'Nonsense'),  # P2
]
gemini_findings['MED_6C7571'] = [
    ('electrottonic', 'Typo'),         # P1
    ('increases to 0 bpm', 'Conflicting'),  # P4
    ('magic 8-ball', 'Nonsense'),       # P3
]
gemini_findings['MED_998FA0'] = [
    ('arhythmias', 'Typo'),            # P1
    ('worsen cardiac rhythm', 'Conflicting'),  # P5
    ('pizza cutter', 'Nonsense'),       # P3
]
gemini_findings['MED_29AA9B'] = [
    ('moleclar', 'Typo'),              # P1
    ('incapable of detecting', 'Conflicting'),  # P2
    ('floating pineapples', 'Nonsense'),  # P4
]
gemini_findings['MED_F552C1'] = [
    ('inflamation', 'Typo'),           # P1
    ('promote inflammation', 'Conflicting'),  # P3
    ('drum kit in a hurricane', 'Nonsense'),  # P4
]


# ============================================================
# COMPARISON LOGIC
# ============================================================

def check_match(answer_contaminant_text, answer_type, gemini_items):
    """Check if any Gemini finding matches an answer key entry.
    Returns (matched_bool, gemini_text_if_matched)"""
    # Normalize answer
    ans_lower = answer_contaminant_text.lower().strip()
    ans_type_lower = answer_type.lower().strip()
    
    for g_text, g_type in gemini_items:
        g_lower = g_text.lower().strip()
        
        # Check if the contaminant text appears in Gemini's finding or vice versa
        text_match = (ans_lower in g_lower) or (g_lower in ans_lower)
        
        # Also check partial substring for multi-word items
        if not text_match:
            # Check individual words
            ans_words = ans_lower.split()
            for w in ans_words:
                if len(w) > 3 and w in g_lower:
                    text_match = True
                    break
        
        if text_match:
            return True, g_text
    
    return False, None


# ============================================================
# GENERATE REPORT
# ============================================================

output_lines = []
total_matched = 0
total_contaminants = 0
prompt_results = {}

for prompt_name, doc_ids in prompt_docs.items():
    output_lines.append("=" * 80)
    output_lines.append(f"  {prompt_name}")
    output_lines.append("=" * 80)
    
    prompt_matched = 0
    prompt_total = 0
    
    for doc_id in doc_ids:
        ak = answer_key.get(doc_id, [])
        gf = gemini_findings.get(doc_id, [])
        
        output_lines.append(f"\n  {doc_id}: {len(ak)} contaminants in answer key, Gemini reported {len(gf)} items")
        
        doc_matched = 0
        for entry in ak:
            prompt_total += 1
            total_contaminants += 1
            
            matched, g_text = check_match(entry['contaminant'], entry['type'], gf)
            
            status = "YES" if matched else "NO"
            if matched:
                doc_matched += 1
                prompt_matched += 1
                total_matched += 1
            
            output_lines.append(
                f"    {entry['cid']} [{entry['type']}] \"{entry['contaminant']}\" "
                f"→ FOUND: {status}"
                + (f" (Gemini: \"{g_text}\")" if matched else "")
            )
        
        output_lines.append(f"    Score: {doc_matched}/{len(ak)}")
    
    output_lines.append(f"\n  >>> {prompt_name} TOTAL: {prompt_matched}/{prompt_total}")
    prompt_results[prompt_name] = (prompt_matched, prompt_total)
    output_lines.append("")

# ============================================================
# OVERALL SUMMARY
# ============================================================

output_lines.append("\n" + "=" * 80)
output_lines.append("  OVERALL SUMMARY")
output_lines.append("=" * 80)
output_lines.append("")

for pname, (m, t) in prompt_results.items():
    pct = (m/t*100) if t > 0 else 0
    bar = "█" * int(pct/5) + "░" * (20 - int(pct/5))
    output_lines.append(f"  {pname:45s}  {m:3d}/{t:<3d}  ({pct:5.1f}%)  {bar}")

output_lines.append("")
pct_total = (total_matched/total_contaminants*100) if total_contaminants > 0 else 0
output_lines.append(f"  GRAND TOTAL: {total_matched}/{total_contaminants} contaminants matched ({pct_total:.1f}%)")

# Part breakdown
part_i_m = prompt_results.get('Prompt 1 (Part I SC)', (0,0))[0] + prompt_results.get('Prompt 2 (Part I MED)', (0,0))[0]
part_i_t = prompt_results.get('Prompt 1 (Part I SC)', (0,0))[1] + prompt_results.get('Prompt 2 (Part I MED)', (0,0))[1]
part_ii_m = prompt_results.get('Prompt 3 (Part II SC - 7 docs)', (0,0))[0] + prompt_results.get('Prompt 4 (Part II MED - 3 docs)', (0,0))[0]
part_ii_t = prompt_results.get('Prompt 3 (Part II SC - 7 docs)', (0,0))[1] + prompt_results.get('Prompt 4 (Part II MED - 3 docs)', (0,0))[1]
part_iii_m = sum(v[0] for k,v in prompt_results.items() if '5' in k or '6' in k)
part_iii_t = sum(v[1] for k,v in prompt_results.items() if '5' in k or '6' in k)

output_lines.append("")
output_lines.append(f"  Part I  (2 docs,   6 contaminants): {part_i_m}/{part_i_t}")
output_lines.append(f"  Part II (10 docs, 30 contaminants): {part_ii_m}/{part_ii_t}")
output_lines.append(f"  Part III(48 docs,144 contaminants): {part_iii_m}/{part_iii_t}")

# ============================================================
# CRITICAL OBSERVATION
# ============================================================
output_lines.append("")
output_lines.append("=" * 80)
output_lines.append("  CRITICAL OBSERVATION")
output_lines.append("=" * 80)
output_lines.append("""
  Prompts 5A-5D and 6A-6B: Gemini is HALLUCINATING contaminants.
  
  Instead of actually reading the PDFs, Gemini appears to be:
  - Inventing plausible-sounding contaminants
  - Using made-up nonsense terms like "quantum-powered toaster", "telepathic 
    squirrel", "disco-dancing warehouse", "flying pancake", "magic carpet",
    "pizza cutter", "pogo stick equation", etc.
  - These are NOT the actual contaminants embedded in our documents
  
  Our actual contaminants are things like:
  - "dolphin choir" replacing "supply chain"
  - "efficency" (typo for "efficiency")
  - "buisness" (typo for "business")
  - "deteriorations" replacing "improvements"
  
  Gemini found NONE of these in the split batch responses.
  
  However, this is FINE for the HW3 submission because:
  1. The assignment evaluates GenAI capabilities — poor results ARE valid data
  2. We score each response on 5 criteria (Likert 1-5)
  3. Low scores for bad responses are expected and acceptable 
  4. The FOUND: YES/NO comparison against the answer key is exactly what
     demonstrates the gap between GenAI capability and ground truth
  5. Part IV statistical analysis will show the performance patterns
  
  BOTTOM LINE: We have everything needed to proceed with the HW3 submission.
  All 10 prompts have responses (even if most are wrong for Part III).
""")

# Write output
output_path = "gemini_full_analysis_results.txt"
with open(output_path, 'w', encoding='utf-8') as f:
    f.write('\n'.join(output_lines))

print(f"\nResults written to {output_path}")
print(f"\nQUICK SUMMARY:")
for pname, (m, t) in prompt_results.items():
    pct = (m/t*100) if t > 0 else 0
    print(f"  {pname:45s}  {m:3d}/{t:<3d}  ({pct:5.1f}%)")
print(f"\n  GRAND TOTAL: {total_matched}/{total_contaminants} ({pct_total:.1f}%)")
