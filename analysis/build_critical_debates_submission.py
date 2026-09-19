from __future__ import annotations

import csv
import pathlib
import re
from collections import Counter, defaultdict

from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt


ROOT = pathlib.Path(__file__).resolve().parents[1]
OUT = ROOT / "output" / "critical_debates_revision"
FIG_DIR = ROOT / "results" / "figures"

TITLE = "Hearing Genomic Order: Sonifying Mutation Patterns in Xylella fastidiosa"
FILENAME_SAFE_TITLE = "Hearing Genomic Order - Sonifying Mutation Patterns in Xylella fastidiosa"

AUTHOR_NAME = "Albert Zhang"
SCHOOL = "[School Name or District]"
EMAIL = "[Corresponding Email]"


def _font(size: int, bold: bool = False):
    name = "timesbd.ttf" if bold else "times.ttf"
    path = pathlib.Path("C:/Windows/Fonts") / name
    return ImageFont.truetype(str(path), size=size) if path.exists() else ImageFont.load_default()


def _wrapped(draw: ImageDraw.ImageDraw, text: str, font, max_width: int) -> list[str]:
    words = text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        candidate = f"{current} {word}".strip()
        if not current or draw.textbbox((0, 0), candidate, font=font)[2] <= max_width:
            current = candidate
        else:
            lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


def build_combined_figure() -> pathlib.Path:
    """Create one two-panel display so the manuscript remains within the three-item limit."""
    canvas = Image.new("RGB", (2200, 900), "white")
    draw = ImageDraw.Draw(canvas)
    title_font = _font(44, bold=True)
    box_font = _font(34, bold=True)
    detail_font = _font(28)
    label_font = _font(46, bold=True)

    draw.text((28, 18), "A", font=label_font, fill="#111111")
    draw.text((92, 25), "Reproducible sonification workflow", font=title_font, fill="#111111")
    steps = [
        ("Genome assemblies", "50 quality-filtered X. fastidiosa RefSeq records"),
        ("Windowing", "60 non-overlapping 3,000 bp windows per genome"),
        ("Deterministic mapping", "codon64 and amino-acid-class pitch sequences"),
        ("Feature extraction", "pitch distributions, entropy, spread, and intervals"),
        ("Reference comparison", "Euclidean distance from each original window"),
        ("Ranked output", "AUC, Precision@4, coordinates, and MIDI tracks"),
    ]
    left_x, right_x = 85, 765
    box_h, gap, start_y = 112, 20, 92
    fills = ["#EAF2F8", "#E8F6F3", "#FCF3CF", "#FDEBD0", "#F5EEF8", "#EAECEE"]
    for idx, ((heading, detail), fill) in enumerate(zip(steps, fills)):
        top = start_y + idx * (box_h + gap)
        draw.rounded_rectangle((left_x, top, right_x, top + box_h), radius=18,
                               fill=fill, outline="#44546A", width=3)
        draw.text((left_x + 24, top + 13), heading, font=box_font, fill="#17202A")
        lines = _wrapped(draw, detail, detail_font, right_x - left_x - 48)
        for line_no, line in enumerate(lines[:2]):
            draw.text((left_x + 24, top + 57 + 28 * line_no), line, font=detail_font, fill="#34495E")
        if idx < len(steps) - 1:
            arrow_x = (left_x + right_x) // 2
            y1 = top + box_h + 3
            y2 = top + box_h + gap - 3
            draw.line((arrow_x, y1, arrow_x, y2), fill="#44546A", width=5)
            draw.polygon([(arrow_x - 10, y2 - 12), (arrow_x + 10, y2 - 12), (arrow_x, y2)], fill="#44546A")

    chart = Image.open(FIG_DIR / "anomaly_auc.png").convert("RGB")
    chart.thumbnail((1340, 790), Image.Resampling.LANCZOS)
    chart_x = 825
    chart_y = 78 + max(0, (790 - chart.height) // 2)
    canvas.paste(chart, (chart_x, chart_y))
    draw.text((805, 18), "B", font=label_font, fill="#111111")
    draw.text((868, 25), "Anomaly-detection performance", font=title_font, fill="#111111")

    out = OUT / "Fig. 1 - Workflow and anomaly detection.png"
    canvas.save(out, dpi=(300, 300))
    return out


REFERENCES_APA = [
    "Bhattacharyya, A., Stilwagen, S., Ivanova, N., D'Souza, M., Bernal, A., Lykidis, A., et al. (2002). Whole-genome comparative analysis of three phytopathogenic Xylella fastidiosa strains. Proceedings of the National Academy of Sciences, 99, 12403-12408.",
    "Chatterjee, S., Almeida, R. P. P., & Lindow, S. (2008). Living in two worlds: The plant and insect lifestyles of Xylella fastidiosa. Annual Review of Phytopathology, 46, 243-271.",
    "Chen, J., Xie, G., Han, S., Chertkov, O., Sims, D., & Civerolo, E. L. (2010). Whole-genome sequences of two Xylella fastidiosa strains (M12 and M23) causing almond leaf scorch disease in California. Journal of Bacteriology, 192, 4534.",
    "Coletta-Filho, H. D., Castillo, A. I., Laranjeira, F. F., de Andrade, E. C., Silva, N. T., de Souza, A. A., Bossi, M. E., Lopes, J. R. S., & Almeida, R. P. P. (2020). Citrus variegated chlorosis: An overview of 30 years of research and disease management. Tropical Plant Pathology, 45, 175-191.",
    "Harris, C. R., Millman, K. J., van der Walt, S. J., Gommers, R., Virtanen, P., Cournapeau, D., et al. (2020). Array programming with NumPy. Nature, 585, 357-362.",
    "Hermann, T., Hunt, A., & Neuhoff, J. G. (Eds.). (2011). The Sonification Handbook. Logos Verlag.",
    "Hopkins, D. L., & Purcell, A. H. (2002). Xylella fastidiosa: Cause of Pierce's disease of grapevine and other emergent diseases. Plant Disease, 86, 1056-1066.",
    "Kramer, G. (1994). Auditory Display: Sonification, Audification, and Auditory Interfaces. Addison-Wesley.",
    "Martin, Z., Meagher, J., & Barker, D. (2021). Using sound to understand protein sequence data: New sonification algorithms for protein sequences and multiple sequence alignments. BMC Bioinformatics, 22, 456. https://doi.org/10.1186/s12859-021-04362-7",
    "National Center for Biotechnology Information. (2026). NCBI Datasets genome data package: Xylella fastidiosa RefSeq assemblies. https://www.ncbi.nlm.nih.gov/datasets/",
    "Nunney, L., Vickerman, D. B., Bromley, R. E., Russell, S. A., Hartman, J. R., Morano, L. D., & Stouthamer, R. (2013). Recent evolutionary radiation and host plant specialization in the Xylella fastidiosa subspecies native to the United States. Applied and Environmental Microbiology, 79, 2189-2200.",
    "Plaisier, H., Meagher, J., & Barker, D. (2021). DNA sonification for public engagement in bioinformatics. BMC Research Notes, 14, 273. https://doi.org/10.1186/s13104-021-05685-7",
    "Saponari, M., Boscia, D., Nigro, F., & Martelli, G. P. (2013). Identification of DNA sequences related to Xylella fastidiosa in oleander, almond and olive trees exhibiting leaf scorch symptoms in Apulia, southern Italy. Journal of Plant Pathology, 95, 668.",
    "Sayers, E. W., Cavanaugh, M., Clark, K., Ostell, J., Pruitt, K. D., & Karsch-Mizrachi, I. (2022). GenBank. Nucleic Acids Research, 50, D161-D164.",
    "Schaad, N. W., Postnikova, E., Lacy, G., Fatmi, M., & Chang, C. J. (2004). Xylella fastidiosa subspecies: X. fastidiosa subsp. piercei, subsp. nov., X. fastidiosa subsp. multiplex subsp. nov., and X. fastidiosa subsp. pauca subsp. nov. Systematic and Applied Microbiology, 27, 290-300.",
    "Shannon, C. E. (1948). A mathematical theory of communication. Bell System Technical Journal, 27, 379-423, 623-656.",
    "Sicard, A., Zeilinger, A. R., Vanhove, M., Schartel, T. E., Beal, D. J., Daugherty, M. P., & Almeida, R. P. P. (2018). Xylella fastidiosa: Insights into an emerging plant pathogen. Annual Review of Phytopathology, 56, 181-202.",
    "Simpson, A. J. G., Reinach, F. C., Arruda, P., Abreu, F. A., Acencio, M., Alvarenga, R., et al. (2000). The genome sequence of the plant pathogen Xylella fastidiosa. Nature, 406, 151-157.",
    "Temple, M. D. (2017). An auditory display tool for DNA sequence analysis. BMC Bioinformatics, 18, 221. https://doi.org/10.1186/s12859-017-1632-x",
    "Van Rossum, G., & Drake, F. L. (2009). Python 3 Reference Manual. Python Software Foundation.",
    "Wells, J. M., Raju, B. C., Hung, H. Y., Weisburg, W. G., Mandelco-Paul, L., & Brenner, D. J. (1987). Xylella fastidiosa gen. nov., sp. nov.: Gram-negative, xylem-limited, fastidious plant bacteria related to Xanthomonas spp. International Journal of Systematic Bacteriology, 37, 136-143.",
]


def set_run_font(run, size: int = 10, bold: bool = False, italic: bool = False) -> None:
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:ascii"), "Times New Roman")
    run._element.rPr.rFonts.set(qn("w:hAnsi"), "Times New Roman")
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic


def apply_styles(doc: Document) -> None:
    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    section.header_distance = Inches(0.5)
    section.footer_distance = Inches(0.5)
    for style_name in ["Normal", "List Bullet", "List Number"]:
        style = doc.styles[style_name]
        style.font.name = "Times New Roman"
        style._element.rPr.rFonts.set(qn("w:ascii"), "Times New Roman")
        style._element.rPr.rFonts.set(qn("w:hAnsi"), "Times New Roman")
        style.font.size = Pt(10)
        style.paragraph_format.line_spacing = 1.0
        style.paragraph_format.space_after = Pt(4)
    for style_name in ["Heading 1", "Heading 2", "Heading 3"]:
        style = doc.styles[style_name]
        style.font.name = "Times New Roman"
        style._element.rPr.rFonts.set(qn("w:ascii"), "Times New Roman")
        style._element.rPr.rFonts.set(qn("w:hAnsi"), "Times New Roman")
        style.font.size = Pt(10)
        style.font.bold = True
        style.font.color.rgb = None
        style.paragraph_format.space_before = Pt(8)
        style.paragraph_format.space_after = Pt(3)


def p(doc: Document, text: str = ""):
    para = doc.add_paragraph(text)
    return para


def h(doc: Document, text: str) -> None:
    doc.add_paragraph(text, style="Heading 1")


def shade_cell(cell, fill: str = "EDEDED") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def repeat_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def format_table(table, font_size: int = 8) -> None:
    for row in table.rows:
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for para in cell.paragraphs:
                para.paragraph_format.space_after = Pt(0)
                for run in para.runs:
                    set_run_font(run, size=font_size)


def add_dataset_table(doc: Document, genomes: list[dict]) -> None:
    counts = defaultdict(Counter)
    for row in genomes:
        counts[row["host_group"]][row["assembly_level"]] += 1
    p(doc, "Table 1. Study dataset summary by host-associated group and assembly level.")
    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    headers = ["Host group", "Complete", "Chromosome", "Scaffold", "Total"]
    for i, header in enumerate(headers):
        table.rows[0].cells[i].text = header
        shade_cell(table.rows[0].cells[i])
    repeat_header(table.rows[0])
    for group in ["grapevine", "almond", "olive", "coffee", "other_prunus"]:
        row = table.add_row().cells
        row[0].text = group
        row[1].text = str(counts[group]["Complete Genome"])
        row[2].text = str(counts[group]["Chromosome"])
        row[3].text = str(counts[group]["Scaffold"])
        row[4].text = str(sum(counts[group].values()))
    format_table(table, 10)


def add_results_table(doc: Document, rows: list[dict]) -> None:
    p(doc, "Table 2. Reference-guided anomaly detection by feature set.")
    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    headers = ["Feature set", "Anomaly", "Mean AUC", "95% CI", "Precision@4"]
    for i, header in enumerate(headers):
        table.rows[0].cells[i].text = header
        shade_cell(table.rows[0].cells[i])
    repeat_header(table.rows[0])
    for row in rows:
        cells = table.add_row().cells
        cells[0].text = row["feature_kind"]
        cells[1].text = row["anomaly_type"].replace("_", " ")
        cells[2].text = f"{float(row['mean_auc']):.3f}"
        cells[3].text = f"{float(row['auc_ci_low']):.3f}-{float(row['auc_ci_high']):.3f}"
        cells[4].text = f"{float(row['mean_precision_at_k']):.3f}"
    format_table(table, 10)


def add_figure(doc: Document, path: pathlib.Path, caption: str) -> None:
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.add_run().add_picture(str(path), width=Inches(6.2))
    cap = doc.add_paragraph(caption)
    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in cap.runs:
        set_run_font(run, size=8, italic=True)


def load_inputs():
    genomes = list(csv.DictReader((ROOT / "data" / "study_genomes.csv").open(encoding="utf-8")))
    anomaly = list(csv.DictReader((ROOT / "results" / "anomaly_detection_summary.csv").open(encoding="utf-8")))
    return genomes, anomaly


def build_docx() -> pathlib.Path:
    OUT.mkdir(parents=True, exist_ok=True)
    combined_figure = build_combined_figure()
    genomes, anomaly = load_inputs()
    doc = Document()
    apply_styles(doc)

    title = p(doc)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(TITLE)
    set_run_font(run, 10, bold=True)
    for line in [AUTHOR_NAME, SCHOOL, EMAIL]:
        para = p(doc, line)
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    h(doc, "Abstract")
    p(doc, "Background/Objective: Genomic sonification can turn DNA sequence patterns into sound, but many demonstrations do not test whether sound-based features reveal information beyond ordinary sequence summaries. This study tested whether windowed sonification can represent local mutation and anomaly patterns in Xylella fastidiosa, a bacterial plant pathogen. Methods: Fifty NCBI RefSeq assemblies were selected from five host-associated groups: grapevine, almond, olive, coffee, and other Prunus hosts. Genomes were divided into 3,000 bp windows. Two deterministic sonification maps were evaluated: a 64-codon pitch map and an amino-acid-class pitch map. Audio features were compared with standard base/codon-composition features, and MIDI anomaly tracks were exported. The study tested host-group representation with leave-one-genome-out nearest-centroid classification and tested reference-guided anomaly detection using background substitutions plus planted clustered substitutions, foreign-segment replacements, and codon-order shuffles. Results: Codon64 sonification matched composition features for host-group classification (40% accuracy, p = 0.0196; chance = 20%). For codon-order shuffles, composition features failed (AUC = 0.499), while codon64 and amino-class sonification performed strongly (AUC = 0.974 and 0.958). For foreign-segment anomalies, codon64 sonification reached AUC = 0.993. Conclusions: Sonification did not replace alignment or variant calling. Its strongest use was as an interpretable representation of genomic order anomalies, especially changes that preserve composition but alter sequential structure.")
    p(doc, "Keywords: Xylella fastidiosa; genome sonification; anomaly detection; codon order; auditory display; comparative genomics; MIDI.")
    p(doc, "Lede: Windowed genome sonification converts Xylella DNA into pitch and interval patterns, revealing codon-order anomalies that ordinary composition summaries miss while preserving a cautious role for standard bioinformatics tools and workflows.")

    h(doc, "Introduction")
    p(doc, "Xylella fastidiosa is a Gram-negative, xylem-limited bacterial plant pathogen associated with Pierce's disease of grapevine, citrus variegated chlorosis, almond leaf scorch, coffee leaf scorch, and olive quick decline syndrome (Wells et al., 1987; Hopkins & Purcell, 2002; Chatterjee et al., 2008; Sicard et al., 2018; Saponari et al., 2013; Coletta-Filho et al., 2020). Its small bacterial genome, broad host range, and large public sequence record make it a useful system for student-scale comparative genomics (Simpson et al., 2000; Bhattacharyya et al., 2002; Schaad et al., 2004; Chen et al., 2010; Nunney et al., 2013; National Center for Biotechnology Information, 2026; Sayers et al., 2022).")
    p(doc, "The central problem addressed here is not whether music can replace standard bioinformatics. It cannot. Alignment, variant calling, gene annotation, and phylogenetic analysis remain the correct tools for biological inference. The open question is narrower: can sonification provide a useful representation layer for genomic changes that are difficult to notice from composition summaries alone?")
    p(doc, "Previous DNA and protein sonification studies have shown that biological sequences can be mapped to sound for exploration, education, and pattern display (Temple, 2017; Martin et al., 2021; Plaisier et al., 2021; Hermann et al., 2011; Kramer, 1994). However, many sonification projects remain qualitative. They describe sequences as more or less musical without testing whether audio-derived features recover known sequence differences. A strong sonification study therefore needs explicit mappings, negative controls, baseline methods, and ground-truth anomalies.")
    p(doc, "This study reframes the project as a windowed anomaly-representation benchmark. The hypothesis was that deterministic sonification features would preserve some host-associated sequence signal and would be especially useful for detecting order-based anomalies, such as codon-order shuffles, that leave base and codon composition largely unchanged.")
    p(doc, "The motivation for using sonification in this setting is not that sound is inherently superior to visual sequence analysis. Instead, auditory display offers a different representation of ordered information. A genome is a long symbolic sequence, and many biologically meaningful changes are not single isolated letters but disruptions in local order, repetition, spacing, and transition patterns. Humans are often good at noticing abrupt changes in rhythm, contour, and repeated intervals. If genomic order can be encoded into pitch and interval structure in a reproducible way, then sonification may become a useful exploratory companion to conventional computational genomics.")
    p(doc, "This motivation also matters for science communication. Genomic data are difficult to explain because raw sequences are too long to inspect directly, while standard plots can compress away the sequential character of DNA. A sonified genome window is not meant to be a literal biological sound. It is a controlled translation: codons become pitches, changes in codon order become changes in melodic contour, and computational features summarize those contours. The value of the approach must therefore be judged by whether the translation preserves useful information and whether it makes a specific class of patterns easier to represent.")
    p(doc, "For a student research project, this framing creates a clearer standard of evidence. Rather than claiming that a whole genome can simply be heard, the study asks whether particular sonification maps behave differently from composition-only baselines under controlled tests. If both composition and sonification perform equally, the sound representation adds little. If sonification performs well when composition fails, then the method has identified a narrower but more defensible use case: representing order-dependent anomalies.")

    h(doc, "Materials & Methods")
    p(doc, "This was a computational methods study using public bacterial genome assemblies. The analysis had two parts. First, genome-level audio and composition features were tested for host-group representation. Second, window-level features were tested for reference-guided anomaly detection. The second task was the main experiment because it directly tested whether sonification can represent mutation or anomaly windows.")
    p(doc, "Genome metadata were retrieved from the NCBI Datasets API on July 12, 2026. Assemblies were restricted to RefSeq records for X. fastidiosa with assembly levels Complete Genome, Chromosome, or Scaffold. Records were quality-filtered using CheckM metadata when available: completeness had to be at least 92.5%, and contamination had to be no more than 3%. The final dataset contained 50 assemblies, with 10 assemblies in each host-associated group.")
    add_dataset_table(doc, genomes)
    p(doc, "FASTA files were parsed by removing headers and retaining only A, C, G, and T. Each genome was divided into non-overlapping 3,000 bp windows. To keep the analysis balanced and reproducible, the first 60 complete windows from each assembly were used, giving 3,000 genome windows across the 50-assembly dataset.")
    p(doc, "Two deterministic sequence-to-pitch maps were evaluated. In the codon64 map, the 64 possible codons were ordered lexicographically from AAA through TTT and mapped to MIDI pitches 36 through 99. In the amino-class map, each codon was translated using the standard bacterial genetic code and assigned to one of seven biochemical pitch classes: hydrophobic, polar, acidic, basic, aromatic, special, or stop.")
    p(doc, "Audio features were calculated from the pitch sequence for each 3,000 bp window. Let X = (x1, x2, ..., xm) be the MIDI pitch sequence for a window. Pitch distribution was the normalized histogram of pitch values. Pitch entropy was calculated as H = -sum_i P(x_i) log2 P(x_i), where P(x_i) is the probability of pitch class i in the window and zero-probability terms were omitted (Shannon, 1948). Mean pitch was mean(X), and pitch spread was the standard deviation of X. First-order intervals were defined as d(j) = x(j+1) - x(j). Interval size was mean(|d(j)|), interval variation was sd(d(j)), and second-order interval change was mean(|d(j+1) - d(j)|). Distances between a reference window and a modified window were Euclidean distances between their feature vectors.")
    p(doc, "The baseline was a non-audio sequence summary containing base frequencies, GC proportion, and 64 codon frequencies. This baseline was intentionally strong for composition-changing anomalies. A meaningful sonification result therefore required audio features to perform well where composition features did not.")
    p(doc, "The composition baseline was designed as a conservative comparison rather than as a weak straw-man method. It captures the overall contents of a sequence window but intentionally does not preserve the exact order of codons beyond short local counts. This makes the baseline appropriate for testing whether sonification features add information about sequence order. If a mutation changes which bases or codons are present, composition should detect it. If an anomaly rearranges the same codons without changing their frequencies, composition should struggle.")
    p(doc, "Two evaluation tasks were used because they ask different questions. The host-group classification task tested whether the features retain broad biological signal across naturally occurring assemblies. The anomaly-detection task tested a more focused display question: when a reference sequence is compared with modified versions of itself, do sonification features make abnormal windows stand out? The second task is closer to the practical use case proposed in this paper, because genome analysts often compare a sample to a reference and ask where unusual local changes occur.")
    p(doc, "For each genome, window-level features were summarized by feature means and standard deviations. A leave-one-genome-out nearest-centroid classifier predicted the host-associated group for the held-out assembly. Label permutation testing with 50 permutations estimated whether accuracy exceeded chance. Because the dataset had five balanced groups, chance accuracy was 20%.")
    p(doc, "The main anomaly experiment simulated a practical use case: a sample sequence is compared with a reference sequence, and sonification is used to make high-change windows audible or visually trackable. Each window first received low background noise through 0.1% random substitutions. Four anomaly windows per genome were then planted for each anomaly type. Clustered substitutions added 1.5% substitutions within selected windows. Foreign-segment anomalies replaced the middle fifth of a window with sequence from another host group. Codon-order shuffles randomly permuted codons across the middle 60% of a window, preserving codon composition while disrupting local order.")
    p(doc, "This design also provides internal controls. Clustered substitutions are expected to be visible to most feature sets because they alter composition and local order at the same time. Foreign-segment replacements are also expected to be detectable because they introduce sequence drawn from a different host-associated group. Codon-order shuffles are the most important stress test. They deliberately remove a major advantage of composition features by keeping codon counts fixed. A successful result under this condition would support the claim that sonification can encode sequential structure rather than merely restating nucleotide composition in musical form.")
    p(doc, "For each modified genome, the analysis ranked windows by distance from the original reference window. AUC measured whether planted anomaly windows ranked above background-mutated windows. Precision@4 measured the fraction of true anomaly windows among the four highest-scoring windows. The value 4 was chosen because exactly four anomaly windows were planted per genome for each anomaly type. Bootstrap confidence intervals used 500 resamples across genome-level AUC values.")
    p(doc, "All analysis steps were deterministic after the random seed was fixed. The pipeline recorded genome accessions, host-group labels, assembly levels, selected windows, feature matrices, classification scores, anomaly scores, summary statistics, and exported MIDI anomaly tracks. The generated MIDI files were not used as subjective evidence; they were treated as reproducible auditory displays derived from the same pitch sequences used in the quantitative feature analysis.")
    p(doc, "Scripts were written in Python with NumPy support (Harris et al., 2020; Van Rossum & Drake, 2009). The study used only public bacterial genome assemblies and did not involve human participants, animals, private data, or clinical decision-making.")

    h(doc, "Results")
    p(doc, "The final dataset contained 50 assemblies and five balanced host-associated groups: grapevine, almond, olive, coffee, and other Prunus, with 10 assemblies per group. The windowed analysis used 60 windows per assembly, producing 3,000 genome windows.")
    p(doc, "The host-group representation test was modest but above chance for codon64 sonification. Composition features and codon64 audio features each reached 40% leave-one-genome-out accuracy with permutation p = 0.0196. Amino-class sonification reached 30% accuracy with p = 0.0588. These results suggest that the codon-level audio representation retained some host-associated sequence signal, but the result should not be interpreted as a diagnostic classifier.")
    p(doc, "This classification result should be interpreted cautiously. Forty percent accuracy is not high enough for diagnostic use, and host association is only an imperfect proxy for evolutionary lineage, geography, and sampling history. However, the result is useful because it shows that the sonification features were not random decorations applied after the fact. They retained enough information from the DNA sequence to perform better than chance in a difficult five-class task. In other words, the pitch and interval features carried biological signal, even if they were not optimized as a classifier.")
    p(doc, "The anomaly benchmark showed where sonification was most useful. For clustered substitutions and foreign-segment replacements, the composition baseline was very strong because these changes altered base or codon frequencies. Codon64 and amino-class sonification also performed well, but they did not outperform composition for those composition-changing anomalies.")
    p(doc, "The clearest sonification-specific result appeared in the codon-order shuffle condition. Composition features performed at chance (mean AUC = 0.499, 95% CI 0.456-0.539) because codon counts were preserved. In contrast, codon64 sonification reached mean AUC = 0.974 (95% CI 0.967-0.980), and amino-class sonification reached mean AUC = 0.958 (95% CI 0.945-0.969). The interval structure of the audio representation captured sequence-order disruption that composition summaries missed.")
    p(doc, "Precision@4 showed the same pattern in a practical ranked-list form. For codon-order shuffles, codon64 sonification recovered true anomaly windows much more often than the composition baseline (Precision@4 = 0.670 versus 0.050), while amino-class sonification reached Precision@4 = 0.730.")
    p(doc, "The contrast between AUC and Precision@4 is useful. AUC measures whether anomaly windows tend to receive higher scores than ordinary windows across all possible thresholds. Precision@4 asks a stricter practical question: if a researcher only listens to or inspects the four most suspicious windows, how often are those windows truly anomalous? The codon-order shuffle condition therefore gives the strongest support for the paper's central argument. The sonification features did not simply score every window slightly differently; they moved many true order anomalies into the very top of the ranked list.")
    p(doc, "Foreign-segment anomalies were also strongly detected by sonification. Codon64 sonification reached mean AUC = 0.993, and amino-class sonification reached mean AUC = 0.984. These results make sense because replacing a window with sequence from another host-associated group can change both composition and order. In this condition, sonification and composition are not competing explanations as much as complementary representations. The important point is that the sound-derived features remained sensitive when the anomaly was biologically interpretable as an inserted or replaced segment.")
    p(doc, "Clustered substitutions produced near-perfect AUC for all feature sets. This result is less surprising because the anomaly is intentionally strong: many bases are changed in a compact region. It functions mainly as a positive control showing that the pipeline can detect obvious local disruption. The more informative result is that composition performed perfectly on clustered substitutions but collapsed on codon-order shuffles, whereas sonification retained strong performance across both. This difference helps separate general anomaly detection from order-specific anomaly representation.")
    add_results_table(doc, anomaly)
    add_figure(doc, combined_figure, "Fig. 1. Workflow and reference-guided anomaly performance. (A) The pipeline converts quality-filtered genome assemblies into fixed windows, deterministic pitch sequences, audio-derived features, ranked anomaly scores, and synchronized MIDI outputs. (B) Codon-order shuffles provide the main use case: composition features fail while order-sensitive sonification features remain highly informative.")

    h(doc, "Discussion")
    p(doc, "The revised study supports a focused claim: sonification can represent certain genomic anomaly patterns in a way that ordinary composition summaries do not. The strongest evidence is the codon-order shuffle experiment. Because that anomaly preserves codon composition, a base/codon frequency baseline has little information to use. Sonification features based on pitch intervals and pitch order still detected those anomaly windows with high AUC.")
    p(doc, "This does not mean that sonification discovers biological truth by itself. The analysis was reference-guided, and the anomaly labels were synthetic. The method should be understood as a display and feature-extraction layer that can sit on top of standard sequence comparison. Its value is not replacing BLAST, alignment, or variant calling, but helping represent where a genome differs from a reference and making order-based changes more perceptible.")
    p(doc, "The perceptual rationale for this result is plausible but not yet proven by listener testing. Human hearing is highly sensitive to changes in pitch contour, repeated interval patterns, and sudden melodic disruptions. Because codon-order shuffles preserve the same codon inventory while changing the order in which codons occur, they can leave composition summaries unchanged but alter the interval sequence of a sonified track.")
    p(doc, "A useful analogy is the difference between a word-frequency table and a sentence. Two sentences can contain the same words but communicate very different meanings if the order changes. Composition features are powerful because they summarize what a genome window contains, but they do not fully describe the order in which symbols appear. Sonification is valuable here because its pitch sequence changes when order changes. The listener or algorithm can then respond to contour, interval, and transition patterns rather than only to counts.")
    p(doc, "The study also suggests a realistic workflow. A researcher would not listen to an entire bacterial genome from beginning to end. Instead, the pipeline could calculate anomaly scores across windows, export short audio clips or MIDI tracks for the most unusual regions, and pair those sounds with ordinary sequence coordinates. In that workflow, sonification becomes a triage and communication tool: it helps prioritize windows for closer inspection and gives students, collaborators, or public audiences a direct way to experience sequence order as a pattern.")
    p(doc, "One practical application is recombination screening in comparative bacterial genomics. Standard alignment, phylogenetic, or recombination-detection software should first establish candidate regions. Sonification could then provide a synchronized secondary track in which neighboring windows are heard and viewed with the same mapping. A sharp change in interval behavior at a candidate breakpoint would not prove recombination, but it could help an analyst compare the extent and local structure of the event across strains. This is especially relevant for X. fastidiosa because host association, lineage, and exchange of genetic material can be difficult to separate using a single summary statistic. The proposed role is therefore confirmatory and exploratory: sonification helps inspect a region already grounded in conventional genomic evidence.")
    p(doc, "A second application is genome-assembly quality assessment. Misassembled joins, contaminant segments, duplicated regions, and unusual low-complexity stretches can create abrupt transitions between adjacent sequence windows. An order-sensitive audio track could make those transitions easy to locate and communicate, particularly when paired with read coverage, GC content, contig boundaries, and established assembly-quality metrics. Sonification alone could not determine whether a transition is a biological structural variant or a technical artifact. Its contribution would be to flag and represent the transition so that raw reads, alignments, and assembly graphs can be checked at the same coordinates.")
    p(doc, "A third application is comparative review of related bacterial genomes. Analysts could generate aligned, coordinate-linked tracks for a reference and several isolates, rank the windows with the largest feature distances, and examine whether the same region is unusual in one strain or across an entire lineage. Such a system could support the study of horizontally acquired segments, inversions, mobile elements, or unusually divergent loci. The present experiment does not validate those biological categories, but it establishes the prerequisite that the representation responds to order disruption. Any future application should measure agreement with independent annotations and report false-positive regions rather than treating an audible difference as biological evidence by itself.")
    p(doc, "The strongest future test would involve real biological variation rather than synthetic anomalies. For example, later work could compare aligned X. fastidiosa strains, known recombination regions, plasmid segments, or mobile genetic elements, then ask whether sonification highlights the same regions detected by standard comparative-genomics methods. Another useful extension would compare alternative pitch mappings, including mappings based on codon usage, amino-acid physicochemical properties, or learned embeddings. A mapping that is both biologically motivated and perceptually clear would make the method more scientifically meaningful.")
    p(doc, "Computational scalability is manageable for bacterial genomes but requires a different interface for much larger datasets. Feature extraction is linear in the number of sequence symbols processed, and windows can be streamed so that the entire genome does not need to be stored as one audio object. This study analyzed 180,000 bp per assembly (60 windows x 3,000 bp), so processing a complete approximately 2.5 Mb bacterial chromosome would examine roughly fourteen times more sequence per strain. The present study did not benchmark runtime or peak memory, and future implementations should report both. For eukaryotic genomes, one-note-per-codon playback would be impractical; a multiresolution design would first score coarse windows, then sonify only selected regions at finer resolution. Parallel feature extraction and compressed event summaries could further reduce computation and listening time.")
    p(doc, "Machine learning could extend the system without changing its cautious role. Supervised models could combine composition, alignment, and sonification-derived features to rank regions with known structural variants or assembly errors, while unsupervised models could identify unusual interval trajectories without requiring predefined anomaly types. These models would need held-out genomes, interpretable feature reporting, and comparisons against established sequence-based methods. A useful result would not be that a model can classify a sound file; it would be that order-sensitive features improve localization or explanation of a biologically verified event beyond conventional inputs alone.")
    p(doc, "An interactive platform could also link four views of the same region: genomic coordinates, nucleotide or alignment data, feature plots, and short audio playback. Clicking an anomalous window could reveal the reference and sample tracks side by side, while changing the window size or pitch map would update all views. This would make the method more testable because users could trace every audible change back to a sequence operation. The same interface could support education by letting students hear how substitutions, insertions, inversions, or shuffles affect an ordered representation, while clearly labeling which examples are synthetic and which are observed in real genomes.")
    p(doc, "The educational value should also be separated from the analytical value. Even if sonification never becomes a primary genome-analysis method, it may still help learners understand that genomes are ordered systems rather than static strings of percentages. However, this paper deliberately avoids relying on education alone as the justification. The quantitative benchmark is what makes the claim stronger: sonification was not merely engaging or musical; it performed well under a controlled anomaly condition where a standard composition baseline failed.")
    p(doc, "Several limitations remain. First, the anomalies were simulated, not verified natural mutations. Second, the first 60 windows were used for standardization and runtime control, not whole-genome alignment. Third, some host groups required chromosome or scaffold assemblies because complete genomes were not equally available across groups. Fourth, the codon64 mapping is transparent but arbitrary; future work should compare multiple biologically informed pitch maps. Finally, listener studies would be needed to test whether humans can reliably hear the anomaly tracks without seeing the computed scores.")
    p(doc, "A second limitation is that AUC and Precision@4 measure computational separability, not human perceptibility. The results show that pitch-derived features contain order information, but they do not prove that an untrained listener would identify the same windows by ear. Listener experiments could test this directly by presenting short reference and anomaly clips, measuring detection accuracy, and comparing trained and untrained listeners. Such a study would connect the computational evidence to the human-centered promise of auditory display.")
    p(doc, "A third limitation is biological interpretation. A synthetic codon-order shuffle is useful because it isolates sequence order, but natural genomes rarely change through perfectly random within-window codon shuffling. Real structural variation may involve recombination, inversion, insertion, deletion, horizontal gene transfer, or assembly artifacts. Therefore, the present result should be read as proof of principle for order-sensitive representation, not as a claim that the tested anomalies directly model every natural mutation process in X. fastidiosa.")

    h(doc, "Conclusion")
    p(doc, "This paper shows a concrete use case for genomic sonification. Across 50 X. fastidiosa RefSeq assemblies, codon-level sonification preserved modest host-associated sequence signal and strongly detected reference-guided order anomalies. The key finding is that codon-order shuffles were invisible to composition features but visible to sonification-derived pitch and interval features. Sonification is therefore best framed as an interpretable anomaly-representation layer for genomics: not a replacement for standard bioinformatics, but a novel way to encode and communicate changes in sequence order.")
    p(doc, "The most important revision from the original project is the change from asking whether genomes can sound musical to asking whether sound-based features preserve testable information. Under that standard, the study gives a cautious but positive answer. Sonification is most useful when the research question involves order, transition, and local disruption rather than simple composition alone. This narrower claim is more scientifically useful because it identifies when the method should be considered, when it should not, and how future studies can test it against stronger biological references.")
    h(doc, "Ethical Approval Statement")
    p(doc, "No human participants or identifiable data were involved. The study used publicly available bacterial genome assemblies and therefore did not require IRB review.")
    h(doc, "Conflicts of Interest")
    p(doc, "The author declares no conflicts of interest.")
    h(doc, "References")
    for ref in REFERENCES_APA:
        p(doc, ref)

    path = OUT / f"{FILENAME_SAFE_TITLE}.docx"
    doc.save(path)
    return path


def build_response_letter() -> pathlib.Path:
    doc = Document()
    apply_styles(doc)
    title = p(doc)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(title.add_run("Response to Editor and Reviewer"), 10, bold=True)
    subtitle = p(doc)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(subtitle.add_run(TITLE), 10, italic=True)
    p(doc, "Dear Dr. Verma and Reviewer,")
    p(doc, "Thank you for the encouraging evaluation and the invitation to revise and resubmit this manuscript. I retained the paper's focused conclusion—that sonification is a complementary, order-sensitive representation rather than a replacement for established bioinformatics—and expanded the manuscript in the areas identified by the review. The revised version adds practical genomic applications, a scalability analysis, broader future directions, and a conceptual workflow panel. The additions are described point by point below.")

    responses = [
        ("Comment 1: The manuscript could benefit from modest expansion and a stronger connection to practical applications.",
         "Response: The Discussion now explains a realistic use pattern in which standard genomic analyses establish candidate regions and sonification provides a coordinate-linked secondary representation. New paragraphs discuss recombination screening, assembly-quality assessment, and comparative review of related bacterial genomes. Each example states what sonification could contribute and what independent evidence would still be required."),
        ("Comment 2: Add real-world examples such as recombination detection, genome assembly validation, or comparative bacterial genomics.",
         "Response: All three examples have been added. The recombination section frames sonification as a way to inspect candidate breakpoints identified by conventional methods. The assembly-validation section discusses misassembled joins, contamination, duplication, and low-complexity transitions while requiring checks against reads, coverage, contig boundaries, and assembly graphs. The comparative-genomics section proposes coordinate-linked tracks across strains and emphasizes validation against independent annotations and false-positive reporting."),
        ("Comment 3: Briefly discuss computational efficiency and scalability to larger genomes or eukaryotic datasets.",
         "Response: A new scalability paragraph states that feature extraction is linear in the sequence length and can be streamed by window. It quantifies that extending the present 180,000 bp sample per assembly to an approximately 2.5 Mb bacterial chromosome would process roughly fourteen times more sequence. It also acknowledges that runtime and memory were not benchmarked. For eukaryotic data, the manuscript proposes coarse-to-fine multiresolution scoring, selective sonification, parallel feature extraction, and compressed event summaries."),
        ("Comment 4: Expand future directions to include machine learning, interactive visualization, or education.",
         "Response: The revised Discussion adds separate paragraphs on machine learning and interactive platforms. The machine-learning paragraph proposes testable supervised and unsupervised extensions with held-out genomes, interpretability, and conventional baselines. The interface paragraph links genomic coordinates, sequence or alignment data, feature plots, and audio playback. It also explains how the same interface could support education without confusing synthetic demonstrations with observed biological variation."),
        ("Comment 5: Consider a graphical workflow or conceptual figure.",
         "Response: Figure 1 is now a two-panel display. Panel A summarizes the complete pipeline from RefSeq assemblies through windowing, deterministic pitch mapping, feature extraction, reference comparison, and ranked/MIDI outputs. Panel B retains the principal anomaly-detection result. Combining the panels improves accessibility while keeping the manuscript within the journal's limit of three total figures and tables."),
    ]
    for comment, response in responses:
        para = p(doc)
        set_run_font(para.add_run(comment), 10, bold=True)
        p(doc, response)
    p(doc, "I appreciate the reviewer's careful reading and positive recommendation. I believe these revisions broaden the practical context while preserving the manuscript's restrained claims and reproducible experimental focus.")
    p(doc, "Sincerely,\nAlbert Zhang")
    path = OUT / "Response to Editor and Reviewer - Critical Debates.docx"
    doc.save(path)
    return path


def main() -> None:
    path = build_docx()
    response = build_response_letter()
    audit_doc = Document(path)
    audit_text = "\n".join(paragraph.text for paragraph in audit_doc.paragraphs)
    body_text = audit_text.split("\nReferences\n", 1)[0]
    print(path)
    print(response)
    print(f"Words before references: {len(re.findall(r'\b[\w@.-]+\b', body_text))}")
    print(f"Words including references: {len(re.findall(r'\b[\w@.-]+\b', audit_text))}")
    print(f"Display items: {len(audit_doc.tables) + len(audit_doc.inline_shapes)}")


if __name__ == "__main__":
    main()
