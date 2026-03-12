"""
HW3 Part IV: Response-Evaluation-Scores Analysis
ADTA-DAST 5770 — Karan Parekh (Group 4)

This script:
1. Reads evaluation scores from both Excel files (Part I + Part II)
2. Calculates summary statistics
3. Performs EDA with visualizations (unique color per criterion)
4. Generates analysis charts (saved as images)
5. Adds Part IV-A, IV-B, IV-C sections to the Word document
"""

import os
import numpy as np
import openpyxl
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# Try to import matplotlib for charts
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

# ============================================================================
# CONFIGURATION
# ============================================================================
BASE_DIR = r"C:\Users\karan\Downloads\ADTA5770"
OUT_DIR = os.path.join(BASE_DIR, r"one drive\GROUP_4\HW_3\MEMBER_1_Karan_Parekh")
WORD_FILE = os.path.join(OUT_DIR, "Karan Parekh_HW_3_prompts_responses.docx")
EXCEL_I = os.path.join(OUT_DIR, "Karan Parekh_prompts_responses_evaluation_scores_PART_I.xlsx")
EXCEL_II = os.path.join(OUT_DIR, "Karan Parekh_prompts_responses_evaluation_scores_PART_II.xlsx")
CHARTS_DIR = os.path.join(OUT_DIR, "charts")

# 5 Criteria with unique colors
CRITERIA_NAMES = [
    "Usefulness/Relevance",
    "Accuracy/Trustworthiness",
    "Clarity/Coherence",
    "Completeness/Depth",
    "Overall Satisfaction"
]
CRITERIA_COLORS = [
    '#2196F3',  # Blue
    '#F44336',  # Red
    '#4CAF50',  # Green
    '#FF9800',  # Orange
    '#9C27B0',  # Purple
]
CRITERIA_SHORT = ['C1', 'C2', 'C3', 'C4', 'C5']


# ============================================================================
# READ EVALUATION SCORES FROM EXCEL FILES
# ============================================================================
def read_scores_from_excel(filepath, label):
    """Read evaluation scores from an Excel file. Returns list of dicts."""
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    
    scores = []
    row = 31
    while True:
        cid = ws.cell(row=row, column=1).value
        if cid is None:
            break
        
        found = ws.cell(row=row, column=4).value
        c1 = ws.cell(row=row, column=7).value   # G: Criterion 1
        c2 = ws.cell(row=row, column=8).value   # H: Criterion 2
        c3 = ws.cell(row=row, column=9).value   # I: Criterion 3
        c4 = ws.cell(row=row, column=10).value  # J: Criterion 4
        c5 = ws.cell(row=row, column=11).value   # K: Criterion 5
        total = ws.cell(row=row, column=12).value  # L: Total
        doc_id = ws.cell(row=row, column=3).value  # C: Document ID
        ctype = ws.cell(row=row, column=2).value   # B: Contam type
        
        if c1 is not None:
            scores.append({
                'cid': cid,
                'doc_id': doc_id,
                'type': ctype,
                'found': found,
                'c1': c1, 'c2': c2, 'c3': c3, 'c4': c4, 'c5': c5,
                'total': total,
                'source': label
            })
        row += 1
    
    wb.close()
    return scores


def categorize_by_part(scores):
    """Add part labels based on source and position."""
    for s in scores:
        if s['source'] == 'Part I':
            s['part'] = 'Part I'
        elif s['doc_id'] and (s['doc_id'].startswith('SC_') or s['doc_id'].startswith('MED_')):
            # Part II docs: first 30 rows in Part II excel
            s['part'] = s['source']  # Will be refined below
    return scores


# ============================================================================
# CALCULATE SUMMARY STATISTICS
# ============================================================================
def calc_summary_stats(scores):
    """Calculate summary statistics for each criterion."""
    criteria_data = {f'c{i}': [] for i in range(1, 6)}
    criteria_data['total'] = []
    
    for s in scores:
        for i in range(1, 6):
            criteria_data[f'c{i}'].append(s[f'c{i}'])
        criteria_data['total'].append(s['total'])
    
    stats = {}
    for key, values in criteria_data.items():
        arr = np.array(values, dtype=float)
        stats[key] = {
            'count': len(arr),
            'mean': np.mean(arr),
            'median': np.median(arr),
            'std': np.std(arr, ddof=1) if len(arr) > 1 else 0,
            'min': np.min(arr),
            'max': np.max(arr),
            'q1': np.percentile(arr, 25),
            'q3': np.percentile(arr, 75),
            'iqr': np.percentile(arr, 75) - np.percentile(arr, 25),
        }
    
    return stats


def calc_stats_by_found(scores):
    """Calculate stats grouped by FOUND: YES vs NO."""
    yes_scores = [s for s in scores if s['found'] == 'YES']
    no_scores = [s for s in scores if s['found'] == 'NO']
    
    return {
        'YES': calc_summary_stats(yes_scores) if yes_scores else None,
        'NO': calc_summary_stats(no_scores) if no_scores else None,
        'yes_count': len(yes_scores),
        'no_count': len(no_scores)
    }


def calc_stats_by_type(scores):
    """Calculate stats grouped by contaminant type."""
    types = set(s['type'] for s in scores if s['type'])
    result = {}
    for t in types:
        t_scores = [s for s in scores if s['type'] == t]
        result[t] = calc_summary_stats(t_scores)
        result[t]['count'] = len(t_scores)
    return result


# ============================================================================
# CREATE VISUALIZATIONS
# ============================================================================
def create_charts(all_scores, part_i_scores, part_ii_scores):
    """Generate all analysis charts."""
    os.makedirs(CHARTS_DIR, exist_ok=True)
    
    # ---- Chart 1: All Data Points by Criterion (Scatter/Strip Plot) ----
    fig, ax = plt.subplots(figsize=(12, 6))
    for i in range(5):
        key = f'c{i+1}'
        values = [s[key] for s in all_scores]
        # Add jitter for visibility
        jitter = np.random.normal(0, 0.08, len(values))
        x_pos = np.full(len(values), i + 1) + jitter
        ax.scatter(x_pos, values, c=CRITERIA_COLORS[i], alpha=0.5, s=30,
                  label=f'{CRITERIA_SHORT[i]}: {CRITERIA_NAMES[i]}', edgecolors='white', linewidth=0.5)
    
    ax.set_xticks(range(1, 6))
    ax.set_xticklabels([f'{CRITERIA_SHORT[i]}\n{CRITERIA_NAMES[i]}' for i in range(5)], fontsize=8)
    ax.set_ylabel('Likert Score (1-5)')
    ax.set_title('All Evaluation Scores by Criterion — All 180 Contaminants', fontweight='bold')
    ax.set_ylim(0.5, 5.5)
    ax.yaxis.set_major_locator(ticker.MultipleLocator(1))
    ax.grid(axis='y', alpha=0.3)
    ax.legend(loc='upper right', fontsize=7)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, 'chart1_all_data_points.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Chart 1: All data points by criterion ✓")
    
    # ---- Chart 2: Box Plot by Criterion ----
    fig, ax = plt.subplots(figsize=(10, 6))
    data = [[s[f'c{i+1}'] for s in all_scores] for i in range(5)]
    bp = ax.boxplot(data, patch_artist=True, tick_labels=[CRITERIA_SHORT[i] for i in range(5)])
    for patch, color in zip(bp['boxes'], CRITERIA_COLORS):
        patch.set_facecolor(color)
        patch.set_alpha(0.6)
    ax.set_ylabel('Likert Score (1-5)')
    ax.set_title('Distribution of Evaluation Scores by Criterion', fontweight='bold')
    ax.set_ylim(0.5, 5.5)
    ax.yaxis.set_major_locator(ticker.MultipleLocator(1))
    ax.grid(axis='y', alpha=0.3)
    # Add legend
    from matplotlib.patches import Patch
    legend_elements = [Patch(facecolor=CRITERIA_COLORS[i], alpha=0.6, label=f'{CRITERIA_SHORT[i]}: {CRITERIA_NAMES[i]}') for i in range(5)]
    ax.legend(handles=legend_elements, loc='upper right', fontsize=7)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, 'chart2_boxplot.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Chart 2: Box plot by criterion ✓")
    
    # ---- Chart 3: Mean Scores by Part (Grouped Bar Chart) ----
    fig, ax = plt.subplots(figsize=(10, 6))
    parts = ['Part I', 'Part II', 'Part III']
    part_data = [part_i_scores, 
                 [s for s in part_ii_scores if s.get('_is_part_ii', False)],
                 [s for s in part_ii_scores if s.get('_is_part_iii', False)]]
    
    x = np.arange(len(parts))
    width = 0.15
    
    for i in range(5):
        means = []
        for pd in part_data:
            if pd:
                means.append(np.mean([s[f'c{i+1}'] for s in pd]))
            else:
                means.append(0)
        bars = ax.bar(x + i * width - 2*width, means, width, label=f'{CRITERIA_SHORT[i]}: {CRITERIA_NAMES[i]}',
                     color=CRITERIA_COLORS[i], alpha=0.8)
        # Add value labels
        for bar, val in zip(bars, means):
            if val > 0:
                ax.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 0.05,
                       f'{val:.1f}', ha='center', va='bottom', fontsize=7)
    
    ax.set_xticks(x)
    ax.set_xticklabels(parts)
    ax.set_ylabel('Mean Likert Score')
    ax.set_title('Mean Evaluation Scores by Part and Criterion', fontweight='bold')
    ax.set_ylim(0, 5.5)
    ax.legend(fontsize=7)
    ax.grid(axis='y', alpha=0.3)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, 'chart3_mean_by_part.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Chart 3: Mean scores by part ✓")
    
    # ---- Chart 4: FOUND vs NOT FOUND comparison ----
    fig, axes = plt.subplots(1, 2, figsize=(12, 5))
    
    yes = [s for s in all_scores if s['found'] == 'YES']
    no = [s for s in all_scores if s['found'] == 'NO']
    
    for idx, (group, title) in enumerate([(yes, f'FOUND: YES (n={len(yes)})'), (no, f'FOUND: NO (n={len(no)})')]):
        ax = axes[idx]
        if group:
            for i in range(5):
                values = [s[f'c{i+1}'] for s in group]
                jitter = np.random.normal(0, 0.08, len(values))
                x_pos = np.full(len(values), i + 1) + jitter
                ax.scatter(x_pos, values, c=CRITERIA_COLORS[i], alpha=0.5, s=25,
                          label=f'{CRITERIA_SHORT[i]}', edgecolors='white', linewidth=0.5)
        ax.set_xticks(range(1, 6))
        ax.set_xticklabels(CRITERIA_SHORT, fontsize=9)
        ax.set_ylabel('Likert Score')
        ax.set_title(title, fontweight='bold')
        ax.set_ylim(0.5, 5.5)
        ax.yaxis.set_major_locator(ticker.MultipleLocator(1))
        ax.grid(axis='y', alpha=0.3)
        ax.legend(fontsize=7)
    
    plt.suptitle('Evaluation Scores: FOUND vs NOT FOUND', fontweight='bold', y=1.02)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, 'chart4_found_vs_notfound.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Chart 4: FOUND vs NOT FOUND comparison ✓")
    
    # ---- Chart 5: Histogram of Total Scores ----
    fig, ax = plt.subplots(figsize=(10, 5))
    totals = [s['total'] for s in all_scores]
    bins = range(min(totals), max(totals) + 2)
    ax.hist(totals, bins=bins, color='#607D8B', alpha=0.7, edgecolor='white', linewidth=0.8)
    ax.axvline(np.mean(totals), color='red', linestyle='--', linewidth=2, label=f'Mean: {np.mean(totals):.1f}')
    ax.axvline(np.median(totals), color='blue', linestyle='--', linewidth=2, label=f'Median: {np.median(totals):.0f}')
    ax.set_xlabel('Total Score (Sum of 5 Criteria)')
    ax.set_ylabel('Frequency')
    ax.set_title('Distribution of Total Evaluation Scores (All 180 Contaminants)', fontweight='bold')
    ax.legend()
    ax.grid(axis='y', alpha=0.3)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, 'chart5_total_histogram.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Chart 5: Total score histogram ✓")
    
    # ---- Chart 6: Heatmap — Mean Scores by Part and Criterion ----
    fig, ax = plt.subplots(figsize=(8, 4))
    part_labels = ['Part I', 'Part II (10 docs)', 'Part III (48 docs)']
    heatmap_data = []
    for pd in part_data:
        if pd:
            row = [np.mean([s[f'c{i+1}'] for s in pd]) for i in range(5)]
        else:
            row = [0, 0, 0, 0, 0]
        heatmap_data.append(row)
    
    heatmap_data = np.array(heatmap_data)
    im = ax.imshow(heatmap_data, cmap='RdYlGn', aspect='auto', vmin=1, vmax=5)
    ax.set_xticks(range(5))
    ax.set_xticklabels(CRITERIA_SHORT)
    ax.set_yticks(range(3))
    ax.set_yticklabels(part_labels)
    
    for i in range(3):
        for j in range(5):
            text = ax.text(j, i, f'{heatmap_data[i, j]:.1f}', ha='center', va='center',
                          color='black', fontweight='bold', fontsize=11)
    
    plt.colorbar(im, label='Mean Likert Score')
    ax.set_title('Mean Evaluation Scores Heatmap', fontweight='bold')
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, 'chart6_heatmap.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Chart 6: Heatmap ✓")
    
    # ---- Chart 7: Stacked Bar by Contaminant Type ----
    fig, ax = plt.subplots(figsize=(10, 6))
    types = sorted(set(s['type'] for s in all_scores if s['type']))
    type_means = {}
    for t in types:
        t_scores = [s for s in all_scores if s['type'] == t]
        type_means[t] = [np.mean([s[f'c{i+1}'] for s in t_scores]) for i in range(5)]
    
    x = np.arange(len(types))
    width = 0.15
    for i in range(5):
        vals = [type_means[t][i] for t in types]
        ax.bar(x + i * width - 2*width, vals, width, label=f'{CRITERIA_SHORT[i]}: {CRITERIA_NAMES[i]}',
              color=CRITERIA_COLORS[i], alpha=0.8)
    
    ax.set_xticks(x)
    ax.set_xticklabels(types, fontsize=9)
    ax.set_ylabel('Mean Likert Score')
    ax.set_title('Mean Evaluation Scores by Contaminant Type', fontweight='bold')
    ax.set_ylim(0, 5.5)
    ax.legend(fontsize=7)
    ax.grid(axis='y', alpha=0.3)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, 'chart7_by_contaminant_type.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Chart 7: By contaminant type ✓")
    
    return True


# ============================================================================
# ADD PART IV TO WORD DOCUMENT
# ============================================================================
def add_part_iv_to_word(all_scores, part_i_scores, part_ii_only, part_iii_only, overall_stats, found_stats, type_stats):
    """Add Part IV sections to the existing Word document."""
    print("\nLoading Word document to add Part IV...")
    doc = Document(WORD_FILE)
    
    # Helper to add a heading
    def add_heading(text, level=1):
        h = doc.add_heading(text, level=level)
        return h
    
    def add_para(text, bold=False, italic=False, font_size=11):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.font.size = Pt(font_size)
        if bold:
            run.bold = True
        if italic:
            run.italic = True
        return p
    
    def add_chart_image(chart_filename, width_inches=6.0):
        chart_path = os.path.join(CHARTS_DIR, chart_filename)
        if os.path.exists(chart_path):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            run.add_picture(chart_path, width=Inches(width_inches))
            return True
        return False
    
    def set_table_borders(table):
        """Add borders to a table using XML manipulation."""
        from docx.oxml.ns import nsdecls
        from docx.oxml import parse_xml
        tbl = table._tbl
        tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            '  <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '  <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '  <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '</w:tblBorders>'
        )
        tblPr.append(borders)
        if tbl.tblPr is None:
            tbl.insert(0, tblPr)
    
    # Add page break before Part IV
    doc.add_page_break()
    
    # ================================================================
    # PART IV-A: METHODOLOGY DISCUSSION
    # ================================================================
    add_heading("HW 3: PART IV-A: Methodology Discussion", level=1)
    
    add_para(
        "This section discusses the methodology used to analyze the response-evaluation scores "
        "obtained from evaluating Gemini 3.0 PRO's responses across all 60 documents and 180 contaminants "
        "in HW 3 Parts I, II, and III.",
        font_size=11
    )
    
    add_heading("1. Data Collection and Structure", level=2)
    add_para(
        "The evaluation dataset consists of 180 contaminant-level observations collected from "
        "60 documents across three homework parts:\n\n"
        "• Part I: 6 contaminants from 2 documents (1 Supply Chain, 1 Medical) — individual document prompts\n"
        "• Part II: 30 contaminants from 10 documents (7 SC, 3 MED) — batch prompts\n"
        "• Part III: 144 contaminants from 48 documents (32 SC, 16 MED) — split batch prompts\n\n"
        "Each contaminant was evaluated using a 5-point Likert scale across five criteria:\n"
        "C1: Usefulness/Relevance — How useful and relevant was the LLM's identification of this contaminant?\n"
        "C2: Accuracy/Trustworthiness — How accurate and trustworthy was the identification?\n"
        "C3: Clarity/Coherence — How clear and coherent was the reported finding?\n"
        "C4: Completeness/Depth — How complete and detailed was the analysis for this contaminant?\n"
        "C5: Overall Satisfaction — Overall satisfaction with Gemini's handling of this contaminant.\n\n"
        "Likert Scale: 1 = Not at all, 2 = Some, 3 = Fair, 4 = Good, 5 = Excellent"
    )
    
    add_heading("2. Analytical Methodology", level=2)
    add_para(
        "The analysis methodology consists of three main analytic activities:\n\n"
        "A. Summary Statistics\n"
        "For each criterion (C1-C5) and the aggregate Total score, we calculate:\n"
        "• Measures of central tendency: Mean, Median\n"
        "• Measures of dispersion: Standard Deviation, Range (Min-Max), Interquartile Range (IQR)\n"
        "• Quartile values: Q1 (25th percentile), Q3 (75th percentile)\n"
        "• Count (N) for each grouping\n\n"
        "These statistics are computed at multiple levels: overall (all 180 observations), "
        "by homework part (I, II, III), by FOUND status (YES vs NO), and by contaminant type "
        "(Typo, Conflicting Information, Nonsense).\n\n"
        "B. Exploratory Data Analysis (EDA)\n"
        "EDA activities include:\n"
        "• Frequency analysis of score distributions\n"
        "• Cross-tabulation of scores by FOUND status and homework part\n"
        "• Identification of patterns and anomalies in the evaluation data\n"
        "• Comparison of criterion-level scores to identify relative strengths/weaknesses\n\n"
        "C. Visualization\n"
        "All visualizations use a unique color scheme for each criterion:\n"
        "• C1 (Usefulness/Relevance): Blue (#2196F3)\n"
        "• C2 (Accuracy/Trustworthiness): Red (#F44336)\n"
        "• C3 (Clarity/Coherence): Green (#4CAF50)\n"
        "• C4 (Completeness/Depth): Orange (#FF9800)\n"
        "• C5 (Overall Satisfaction): Purple (#9C27B0)\n\n"
        "Visualization types include:\n"
        "1. Strip/scatter plots showing all individual data points per criterion\n"
        "2. Box plots for distributional comparison\n"
        "3. Grouped bar charts for mean comparison across parts\n"
        "4. Side-by-side scatter for FOUND vs NOT FOUND comparison\n"
        "5. Histogram of total scores with mean/median markers\n"
        "6. Heatmap of mean scores by part and criterion\n"
        "7. Grouped bar chart of mean scores by contaminant type"
    )
    
    add_heading("3. Software and Tools", level=2)
    add_para(
        "• Python 3.12 with NumPy for statistical computation\n"
        "• Matplotlib for visualization generation\n"
        "• OpenPyXL for Excel data extraction\n"
        "• python-docx for Word document generation\n"
        "• Google Gemini 3.0 PRO (free tier) as the evaluated LLM"
    )
    
    # ================================================================
    # PART IV-B: ANALYSIS RESULTS
    # ================================================================
    doc.add_page_break()
    add_heading("HW 3: PART IV-B: Response-Evaluation Scores Analysis", level=1)
    
    add_heading("1. Overall Summary Statistics (N=180)", level=2)
    
    # Build summary stats table
    table = doc.add_table(rows=8, cols=7)
    set_table_borders(table)
    
    # Headers
    headers = ['Statistic', 'C1: Useful', 'C2: Accuracy', 'C3: Clarity', 'C4: Complete', 'C5: Overall', 'Total']
    for j, h in enumerate(headers):
        table.rows[0].cells[j].text = h
        for paragraph in table.rows[0].cells[j].paragraphs:
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(9)
    
    # Data rows
    stat_rows = ['count', 'mean', 'median', 'std', 'min', 'max', 'q1']
    stat_labels = ['Count (N)', 'Mean', 'Median', 'Std Dev', 'Min', 'Max', 'Q1 (25th)']
    
    keys = ['c1', 'c2', 'c3', 'c4', 'c5', 'total']
    for row_idx, (stat_key, label) in enumerate(zip(stat_rows, stat_labels), 1):
        table.rows[row_idx].cells[0].text = label
        for col_idx, key in enumerate(keys, 1):
            val = overall_stats[key][stat_key]
            if stat_key == 'count':
                table.rows[row_idx].cells[col_idx].text = str(int(val))
            else:
                table.rows[row_idx].cells[col_idx].text = f'{val:.2f}'
            for paragraph in table.rows[row_idx].cells[col_idx].paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(9)
    
    add_para("")  # spacing
    
    # Chart 1: All data points
    add_heading("2. All Data Points by Criterion", level=2)
    add_para(
        "Figure 1 shows all 180 evaluation scores plotted by criterion with unique colors. "
        "Each dot represents one contaminant's evaluation score. Jitter is applied on the x-axis "
        "for visibility. The heavy concentration at scores 1-2 reflects the predominance of Part III "
        "hallucinated responses."
    )
    add_chart_image('chart1_all_data_points.png', 6.0)
    
    # Chart 2: Box plots
    add_heading("3. Score Distributions (Box Plots)", level=2)
    add_para(
        "Figure 2 presents box plots for each criterion, showing the median (center line), "
        "interquartile range (box), and outliers (dots). The low medians across all criteria "
        "reflect the heavy weight of Part III's poor performance."
    )
    add_chart_image('chart2_boxplot.png', 5.5)
    
    # Chart 3: Mean by part
    add_heading("4. Mean Scores by Homework Part", level=2)
    add_para(
        "Figure 3 compares mean criterion scores across the three homework parts. Part I shows "
        "the highest mean scores (mixed: 100% detection on MED, 0% on SC), Part II shows moderate "
        "scores (60% overall detection), and Part III shows the lowest scores due to widespread hallucination."
    )
    add_chart_image('chart3_mean_by_part.png', 5.5)
    
    # FOUND vs NOT FOUND stats
    add_heading("5. FOUND vs NOT FOUND Comparison", level=2)
    
    yes_count = found_stats['yes_count']
    no_count = found_stats['no_count']
    add_para(
        f"Of 180 contaminants, {yes_count} ({100*yes_count/180:.1f}%) were correctly identified by Gemini (FOUND: YES) "
        f"and {no_count} ({100*no_count/180:.1f}%) were not found (FOUND: NO)."
    )
    
    # FOUND comparison table
    table2 = doc.add_table(rows=3, cols=7)
    set_table_borders(table2)
    headers2 = ['FOUND', 'C1 Mean', 'C2 Mean', 'C3 Mean', 'C4 Mean', 'C5 Mean', 'Total Mean']
    for j, h in enumerate(headers2):
        table2.rows[0].cells[j].text = h
        for paragraph in table2.rows[0].cells[j].paragraphs:
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(9)
    
    for row_idx, found_val in enumerate(['YES', 'NO'], 1):
        table2.rows[row_idx].cells[0].text = found_val
        fstats = found_stats[found_val]
        if fstats:
            for col_idx, key in enumerate(keys, 1):
                table2.rows[row_idx].cells[col_idx].text = f'{fstats[key]["mean"]:.2f}'
                for paragraph in table2.rows[row_idx].cells[col_idx].paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(9)
    
    add_para("")
    add_chart_image('chart4_found_vs_notfound.png', 6.0)
    
    # Total score distribution
    add_heading("6. Total Score Distribution", level=2)
    add_para(
        "Figure 5 shows the frequency distribution of total scores (sum of C1 through C5). "
        "The distribution is heavily right-skewed (towards low scores), with a large concentration "
        "at total score = 6 (corresponding to Part III scores of 1,1,2,1,1)."
    )
    add_chart_image('chart5_total_histogram.png', 5.5)
    
    # Heatmap
    add_heading("7. Mean Scores Heatmap by Part", level=2)
    add_para(
        "Figure 6 provides a color-coded heatmap showing mean scores for each criterion across "
        "the three homework parts. Green indicates higher scores; red indicates lower. The sharp decline "
        "from Part I/II to Part III is clearly visible."
    )
    add_chart_image('chart6_heatmap.png', 5.0)
    
    # By contaminant type
    add_heading("8. Analysis by Contaminant Type", level=2)
    
    # Type stats table
    type_list = sorted(type_stats.keys())
    table3 = doc.add_table(rows=len(type_list) + 1, cols=8)
    set_table_borders(table3)
    headers3 = ['Type', 'Count', 'C1 Mean', 'C2 Mean', 'C3 Mean', 'C4 Mean', 'C5 Mean', 'Total Mean']
    for j, h in enumerate(headers3):
        table3.rows[0].cells[j].text = h
        for paragraph in table3.rows[0].cells[j].paragraphs:
            for run in paragraph.runs:
                run.bold = True
                run.font.size = Pt(9)
    
    for row_idx, t in enumerate(type_list, 1):
        table3.rows[row_idx].cells[0].text = t
        table3.rows[row_idx].cells[1].text = str(type_stats[t]['count'])
        for col_idx, key in enumerate(keys, 2):
            table3.rows[row_idx].cells[col_idx].text = f'{type_stats[t][key]["mean"]:.2f}'
            for paragraph in table3.rows[row_idx].cells[col_idx].paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(9)
    
    add_para("")
    add_chart_image('chart7_by_contaminant_type.png', 5.5)
    
    # ================================================================
    # PART IV-C: COMPREHENSIVE REPORT
    # ================================================================
    doc.add_page_break()
    add_heading("HW 3: PART IV-C: Response-Evaluation Analysis Report", level=1)
    
    # Get key values for the report
    overall_mean_total = overall_stats['total']['mean']
    overall_median_total = overall_stats['total']['median']
    
    pi_stats = calc_summary_stats(part_i_scores)
    pii_stats = calc_summary_stats(part_ii_only) if part_ii_only else None
    piii_stats = calc_summary_stats(part_iii_only) if part_iii_only else None
    
    add_heading("Executive Summary", level=2)
    add_para(
        f"This report presents the results of the response-evaluation scores analysis for ADTA-DAST 5770 HW 3. "
        f"A total of 180 contaminant evaluations were collected across 60 documents using Google Gemini 3.0 PRO "
        f"as the LLM. The overall mean total score was {overall_mean_total:.2f} out of 25 "
        f"(median: {overall_median_total:.1f}), indicating generally poor performance by the LLM in detecting "
        f"intentionally embedded contaminants. Only {found_stats['yes_count']} of 180 contaminants "
        f"({100*found_stats['yes_count']/180:.1f}%) were correctly identified."
    )
    
    add_heading("Key Findings", level=2)
    
    add_para("1. Performance Varies Dramatically by Batch Size", bold=True)
    add_para(
        f"• Part I (2 individual documents): Mean total = {pi_stats['total']['mean']:.1f}/25. "
        f"Gemini correctly identified 3/6 contaminants (50%) when analyzing single documents with focused prompts.\n"
        f"• Part II (10 documents in batches of 7+3): Mean total = {pii_stats['total']['mean']:.1f}/25. "
        f"Gemini correctly identified 18/30 contaminants (60%) with moderately sized batch prompts.\n"
        f"• Part III (48 documents in batches of 8): Mean total = {piii_stats['total']['mean']:.1f}/25. "
        f"Gemini correctly identified only 4/144 contaminants (2.8%), hallucinating almost all findings."
    )
    
    add_para("2. FOUND Status Strongly Correlates with Evaluation Scores", bold=True)
    yes_stats = found_stats['YES']
    no_stats = found_stats['NO']
    add_para(
        f"• FOUND=YES contaminants (n={found_stats['yes_count']}): Mean C1={yes_stats['c1']['mean']:.1f}, "
        f"C2={yes_stats['c2']['mean']:.1f}, C3={yes_stats['c3']['mean']:.1f}, "
        f"C4={yes_stats['c4']['mean']:.1f}, C5={yes_stats['c5']['mean']:.1f}, "
        f"Total={yes_stats['total']['mean']:.1f}\n"
        f"• FOUND=NO contaminants (n={found_stats['no_count']}): Mean C1={no_stats['c1']['mean']:.1f}, "
        f"C2={no_stats['c2']['mean']:.1f}, C3={no_stats['c3']['mean']:.1f}, "
        f"C4={no_stats['c4']['mean']:.1f}, C5={no_stats['c5']['mean']:.1f}, "
        f"Total={no_stats['total']['mean']:.1f}\n"
        f"The difference is substantial: correctly identified contaminants scored approximately "
        f"{yes_stats['total']['mean'] - no_stats['total']['mean']:.1f} points higher on total."
    )
    
    add_para("3. Criterion C3 (Clarity) Scores Slightly Higher Than Others", bold=True)
    add_para(
        f"Across all 180 observations, C3 (Clarity/Coherence) had the highest mean score "
        f"({overall_stats['c3']['mean']:.2f}), while C2 (Accuracy) had the lowest ({overall_stats['c2']['mean']:.2f}). "
        f"This suggests that even when Gemini produced incorrect findings, the responses were relatively "
        f"well-structured and readable, but factually unreliable."
    )
    
    add_para("4. Contaminant Type Influences Detection", bold=True)
    # Find best/worst type
    type_totals = {t: type_stats[t]['total']['mean'] for t in type_stats}
    best_type = max(type_totals, key=type_totals.get)
    worst_type = min(type_totals, key=type_totals.get)
    add_para(
        f"By contaminant type, '{best_type}' had the highest mean total score ({type_totals[best_type]:.1f}) "
        f"and '{worst_type}' had the lowest ({type_totals[worst_type]:.1f}). "
        f"Nonsense contaminants (e.g., 'dolphin choir', 'parachute') tend to be more detectable because they are "
        f"contextually obvious, while typos and subtle meaning reversals are harder for the LLM to detect."
    )
    
    add_heading("Limitations", level=2)
    add_para(
        "1. Gemini's inability to directly access PDF URLs at scale severely impacted Part III results. "
        "The LLM hallucinated contaminants rather than admitting it could not read the documents.\n\n"
        "2. The evaluation scores for Part III reflect the hallucination problem — the low scores (mostly 1s) "
        "are a consequence of the LLM's failure mode, not necessarily a direct critique of its analytical capabilities "
        "when given actual document content.\n\n"
        "3. Self-evaluation scores may reflect evaluator bias. Peer evaluation would provide more balanced assessment.\n\n"
        "4. The sample size for FOUND=YES (n=25) is relatively small compared to FOUND=NO (n=155), "
        "which limits the statistical power of comparisons between these groups."
    )
    
    add_heading("Conclusions and Recommendations", level=2)
    add_para(
        f"The overall evaluation results (mean total: {overall_mean_total:.1f}/25, 13.9% detection rate) "
        f"demonstrate that current LLM capabilities are insufficient for reliable automated document quality auditing, "
        f"particularly at scale with URL-based document access.\n\n"
        f"Recommendations for improving LLM-based contaminant detection:\n"
        f"1. Provide document content directly in the prompt context rather than via URLs.\n"
        f"2. Limit batch sizes to 1-3 documents per prompt for optimal accuracy.\n"
        f"3. Use multi-turn conversations where the LLM can ask clarifying questions.\n"
        f"4. Implement verification loops where identified contaminants are confirmed against source text.\n"
        f"5. Consider fine-tuning models on domain-specific quality analysis tasks.\n\n"
        f"Despite the overall low scores, Gemini showed promise with individual document analysis (Part I medical: "
        f"100% detection) and small-batch processing (Part II: 60% detection), suggesting that with proper "
        f"prompt engineering and document access methods, LLMs can be useful tools for document quality assurance."
    )
    
    # Save
    doc.save(WORD_FILE)
    print(f"\nWord document updated with Part IV: {WORD_FILE}")


# ============================================================================
# MAIN
# ============================================================================
if __name__ == "__main__":
    print("=" * 70)
    print("  HW3 Part IV Analysis — ADTA-DAST 5770 — Group 4")
    print("=" * 70)
    
    # 1. Read scores
    print("\n1. Reading evaluation scores...")
    part_i_scores = read_scores_from_excel(EXCEL_I, 'Part I')
    part_ii_scores = read_scores_from_excel(EXCEL_II, 'Part II/III')
    print(f"   Part I: {len(part_i_scores)} scores")
    print(f"   Part II/III: {len(part_ii_scores)} scores")
    
    # Tag Part II vs Part III entries
    # Part II: first 30 (10 docs × 3 contaminants)
    PART_II_DOC_IDS = {
        'SC_E256CF', 'SC_82604E', 'SC_A4CC09', 'SC_046CB7', 'SC_7B0EBD', 'SC_799B06', 'SC_8F419D',
        'MED_67B704', 'MED_1C56DA', 'MED_B6A65F'
    }
    for s in part_ii_scores:
        if s['doc_id'] in PART_II_DOC_IDS:
            s['_is_part_ii'] = True
            s['_is_part_iii'] = False
        else:
            s['_is_part_ii'] = False
            s['_is_part_iii'] = True
    
    part_ii_only = [s for s in part_ii_scores if s['_is_part_ii']]
    part_iii_only = [s for s in part_ii_scores if s['_is_part_iii']]
    
    all_scores = part_i_scores + part_ii_scores
    print(f"   Total: {len(all_scores)} scores")
    print(f"   Part II only: {len(part_ii_only)}, Part III only: {len(part_iii_only)}")
    
    # 2. Calculate summary statistics
    print("\n2. Calculating summary statistics...")
    overall_stats = calc_summary_stats(all_scores)
    found_stats = calc_stats_by_found(all_scores)
    type_stats = calc_stats_by_type(all_scores)
    
    print(f"   Overall mean total: {overall_stats['total']['mean']:.2f}")
    print(f"   FOUND=YES: {found_stats['yes_count']}, FOUND=NO: {found_stats['no_count']}")
    for key in ['c1', 'c2', 'c3', 'c4', 'c5']:
        name = CRITERIA_NAMES[int(key[1])-1]
        print(f"   {key.upper()} ({name}): mean={overall_stats[key]['mean']:.2f}, median={overall_stats[key]['median']:.1f}")
    
    # 3. Create visualizations
    print("\n3. Creating visualizations...")
    create_charts(all_scores, part_i_scores, part_ii_scores)
    
    # 4. Add Part IV to Word document
    print("\n4. Adding Part IV to Word document...")
    add_part_iv_to_word(all_scores, part_i_scores, part_ii_only, part_iii_only, 
                        overall_stats, found_stats, type_stats)
    
    print("\n" + "=" * 70)
    print("  PART IV ANALYSIS COMPLETE!")
    print("=" * 70)
