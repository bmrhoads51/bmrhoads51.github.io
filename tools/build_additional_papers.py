"""Build the welding and bioprinting article pages from supplied manuscripts.

The bioprinting XML is the structured version of the supplied final_paper.pdf
(same DOI, title, abstract, figures, tables, and article text). Figures are
extracted directly from that PDF. The welding page uses the supplied Word
submission and its separate figure files, with corrections drawn from the
published_paper.pdf supplied later.
"""
from __future__ import annotations

import re
import shutil
from html import escape
from pathlib import Path
from xml.etree import ElementTree as ET

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from PIL import Image
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT.parent


def slug(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")


def shell(title: str, description: str, journal: str, year: str,
          authors: str, citation: str, doi: str, content: str,
          toc: list[tuple[str, str]]) -> str:
    nav = "".join(f'<a href="#{ident}">{escape(label)}</a>' for label, ident in toc)
    return f'''<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
  <meta name="description" content="{escape(description, quote=True)}">
  <title>{escape(title)} | Benjamin Rhoads</title>
  <link rel="stylesheet" href="style.css?v=20260922-figure-spacing">
</head>
<body>
  <a class="skip-link" href="#main">Skip to content</a>
  <header class="site-header"><div class="wrap header-inner"><a class="brand" href="index.html">Benjamin Rhoads</a><nav aria-label="Main navigation"><a href="index.html">Home</a><a href="publications.html">Publications</a><a href="projects.html">Projects</a></nav></div></header>
  <main id="main" class="wrap paper-layout">
    <aside class="paper-sidebar" aria-label="Article navigation"><p class="eyebrow">On this page</p>{nav}</aside>
    <article class="paper">
      <div class="paper-top"><p class="eyebrow">Publication · {escape(journal)} · {year}</p><h1>{escape(title)}</h1><p class="authors">{escape(authors)}</p><p class="journal"><em>{escape(journal)}</em> {escape(citation)} · DOI: {escape(doi)}</p><div class="actions"><a class="button" href="https://doi.org/{escape(doi, quote=True)}" target="_blank" rel="noopener noreferrer">Read published article ↗</a></div></div>
      {content}
    </article>
  </main>
  <footer class="site-footer"><div class="wrap">© Benjamin Rhoads · <a href="index.html">Home</a> · <a href="https://github.com/bmrhoads51" target="_blank" rel="noopener noreferrer">GitHub</a></div></footer>
</body>
</html>
'''


# Word equations use OMML. Render the common structures as selectable HTML.
W_VAL = qn("w:val")


def math_text(node) -> str:
    name = node.tag.split("}")[-1]
    children = list(node)
    if name == "t":
        return escape(node.text or "")
    if name == "f":
        num = next((x for x in children if x.tag.endswith("}num")), None)
        den = next((x for x in children if x.tag.endswith("}den")), None)
        return f'<span class="fraction"><span>{math_text(num) if num is not None else ""}</span><span>{math_text(den) if den is not None else ""}</span></span>'
    if name in ("sSup", "sSub", "sSubSup"):
        base = next((x for x in children if x.tag.endswith("}e")), None)
        sub = next((x for x in children if x.tag.endswith("}sub")), None)
        sup = next((x for x in children if x.tag.endswith("}sup")), None)
        return (math_text(base) if base is not None else "") + (f"<sub>{math_text(sub)}</sub>" if sub is not None else "") + (f"<sup>{math_text(sup)}</sup>" if sup is not None else "")
    if name == "rad":
        value = next((x for x in children if x.tag.endswith("}e")), None)
        return f'<span class="radical">√<span>{math_text(value) if value is not None else ""}</span></span>'
    if name == "d":
        props = next((x for x in children if x.tag.endswith("}dPr")), None)
        begin, end = "(", ")"
        if props is not None:
            b = next((x for x in props if x.tag.endswith("}begChr")), None)
            e = next((x for x in props if x.tag.endswith("}endChr")), None)
            if b is not None:
                begin = b.get(W_VAL, begin)
            if e is not None:
                end = e.get(W_VAL, end)
        return escape(begin) + "".join(math_text(x) for x in children if not x.tag.endswith("}dPr")) + escape(end)
    if name == "brk" or name == "tab":
        return " "
    if name.endswith("Pr") or name in ("rPr", "ctrlPr"):
        return ""
    return "".join(math_text(x) for x in children)


def run_html(run: Run) -> str:
    pieces = []
    for child in run._r:
        tag = child.tag.split("}")[-1]
        if tag == "t":
            pieces.append(escape(child.text or ""))
        elif tag == "tab":
            pieces.append(" ")
        elif tag in ("br", "cr"):
            pieces.append("<br>")
        elif tag == "oMath":
            pieces.append(f'<span class="math">{math_text(child)}</span>')
    value = "".join(pieces)
    if run.bold:
        value = f"<strong>{value}</strong>"
    if run.italic:
        value = f"<em>{value}</em>"
    return value


def paragraph_html(p: Paragraph) -> str:
    out = []
    for child in p._p:
        tag = child.tag.split("}")[-1]
        if tag == "r":
            out.append(run_html(Run(child, p)))
        elif tag in ("oMath", "oMathPara"):
            out.append(f'<span class="math">{math_text(child)}</span>')
        elif tag == "hyperlink":
            out.extend(run_html(Run(r, p)) for r in child if r.tag.endswith("}r"))
    return "".join(out).strip()


def word_table(table: Table, caption: str) -> str:
    # Read physical cells: row.cells repeats horizontally merged cells and would
    # otherwise duplicate their text in the HTML table.
    matrix = []
    spans = []
    for row in table.rows:
        cells = [_Cell(tc, table) for tc in row._tr.tc_lst]
        matrix.append(["<br>".join(escape(p.text) for p in cell.paragraphs if p.text.strip()) for cell in cells])
        spans.append([cell._tc.tcPr.gridSpan.val if cell._tc.tcPr is not None and cell._tc.tcPr.gridSpan is not None else 1 for cell in cells])
    header_rows = 2 if caption.startswith("Table 1:") else 1
    # Word's vertically merged cells report their value only near the center.
    # Repeat those shared values so each HTML data row remains understandable.
    if matrix:
        for col in range(len(matrix[0])):
            filled = [(i, row[col]) for i, row in enumerate(matrix[header_rows:], header_rows) if col < len(row) and row[col]]
            for i in range(header_rows, len(matrix)):
                if col < len(matrix[i]) and not matrix[i][col] and filled:
                    matrix[i][col] = min(filled, key=lambda item: (abs(item[0] - i), item[0] > i))[1]
    if caption.startswith("Table 2:"):
        for row in matrix[1:]:
            sample = row[0].replace("*", "")
            if sample.isdigit() and 1 <= int(sample) <= 20:
                row[2] = "24.00" if int(sample) <= 13 else "18.00"
    rows = []
    for i, row in enumerate(matrix):
        cells = []
        for value, span in zip(row, spans[i]):
            tag = "th" if i < header_rows else "td"
            colspan = f' colspan="{span}"' if span > 1 else ""
            cells.append(f"<{tag}{colspan}>{value}</{tag}>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return f'<div class="table-scroll"><table><caption>{caption}</caption><tbody>{"".join(rows)}</tbody></table></div>'


def copy_figure(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    with Image.open(src) as im:
        im = im.convert("RGB")
        if im.width > 2000:
            im.thumbnail((2000, 2000))
        im.save(dst, "JPEG", quality=88, optimize=True)


def image_size(path: Path) -> tuple[int, int]:
    with Image.open(path) as im:
        return im.size


def welding_results_section() -> str:
    width, height = image_size(ROOT / "assets" / "welding" / "figure-10.jpg")
    return f'''<p>In addition to the common ML tools discussed in Sections 3.2 and 3.3, this work also explored XAI in studying FSW. Symbolic regression, an XAI tool, was used to extract mathematical equations that describe the comprehensive relationship between features and hardness difference in the HAZ. It aimed to uncover mathematical expressions that provide a clear and interpretable relationship between input features (processing features) and output targets (hardness difference), offering greater transparency in understanding the underlying physical mechanisms of the process. Using symbolic regression, the minimum test error was obtained with the following equation:</p>
<div class="equation" id="equation-11"><span class="math">ΔHV<sub>0.3</sub> = <span class="fraction"><span>−49.4982 − ω</span><span>p</span></span> + <span class="fraction"><span>((t/(v + 3.5939))<sup>0.7303</sup> × 3.3194 + 5.3883)</span><span>0.1282</span></span> + t − <span class="fraction"><span>d</span><span>(v − 1.66)(d − v + 2.0877)</span></span> − 2.3939</span><span class="equation-number">Equation 11</span></div>
<p>where ΔHV<sub>0.3</sub> is the hardness difference (HV 0.3), t represents the plate thickness (mm), v is the tool traverse speed (ipm), p is the pseudo heat index (rpm²/ipm), d is the tool shoulder diameter (mm), and ω is the tool rotational rate (rpm).</p>
<p>To assess the performance and feasibility of Eq. (11), the predicted hardness differences based on the given equation were compared to experimental results using two evaluation metrics: RMSE and R² (see Fig. 10). The horizontal axis represents the experimental target (actual measured hardness difference values), while the vertical axis denotes the predicted target (hardness values generated by Eq. (11)). The diagonal line, extending from the bottom left to the top right, denotes the parity line, where the experimental value is equal to the predicted value. Only test data is shown, which clusters around the parity line with no significant outliers – indicating that the equation offers reliable predictions. Eq. (11) has an RMSE of 3.96 HV 0.3, and an R² of 0.90. The RMSE of 3.96 HV 0.3 occupies only 7.94 % of the total range of hardness differences (49.90 HV 0.3) observed in the dataset. It was collected by calculating the reduction from the observed maximum hardness difference (90.09 HV 0.3, Data point 23 in Table 2) to the minimum hardness difference (41.00 HV 0.3, Data point 38 in Table 6). This finding indicates a relatively small prediction error. The R² value of 0.90 denotes that 90 % of the variation in the hardness difference can be explained by the features included in the model. These results therefore demonstrate that Eq. (11) is effective in predicting the hardness difference in the HAZ. To improve the performance of Eq. (11) in the future, additional data and nonlinear relationships would make it possible to achieve a lower error.</p>
<figure id="figure-10"><img src="assets/welding/figure-10.jpg" width="{width}" height="{height}" alt="Predicted versus measured hardness difference" loading="lazy"><figcaption>Figure 10. Hardness difference predicted by Equation 11 versus experimental hardness difference.</figcaption></figure>
<p>Other simpler equations were also extracted from this model with higher errors. One of the equations extracted is provided as below:</p>
<div class="equation" id="equation-12"><span class="math">ΔHV<sub>0.3</sub> = 45.73 + <span class="fraction"><span>30.46 × t</span><span>v + 3.98</span></span></span><span class="equation-number">Equation 12</span></div>
<p>Eq. (12) is a simpler equation compared to Eq. (11), with a test RMSE of 5.91 and a test R² value of 0.77, containing only two features: tool traverse speed (v) and plate thickness (t). This equation also supports the finding from the earlier feature importance analysis (see Fig. 7), which shows that the tool traverse speed plays a more significant role in governing the hardness difference in the HAZ compared to the other four features. Furthermore, the feature importance analysis results show that all features have a notable impact on the hardness difference, which is validated by the accuracy reduction from Eq. (11) (5 features) to Eq. (12) (2 features). Overall, the two equations generated via XAI provide humanly interpretable relationship between processing features and hardness difference.</p>'''


WELDING_FINAL_CONCLUSIONS = [
    "In this study, the relationship between processing features and hardness difference in the HAZ during FSW was examined using ML techniques. Five processing features were analyzed: plate thickness, tool traverse speed, tool rotational rate, tool shoulder diameter, and pseudo heat index. Kendall correlation coefficient analysis indicated that each of the analyzed features makes a distinct contribution to hardness difference in the HAZ. Furthermore, the Random Forest and Extra Trees models both confirmed that among all the five features the tool traverse speed is the most influential feature affecting the hardness difference. Later, only after three iterations and six ML guided data points, a hardness difference of 41.00 HV 0.3 in the HAZ was achieved. In contrast, a minimum hardness difference of 42.00 HV 0.3 was obtained from 32 data points when the experiments were performed without the guidance from ML. It can be concluded that ML can significantly reduce the requirement of high cost and time-consuming experiments to optimize mechanical properties in the HAZ during FSW. Finally, a high fidelity mathematical equation linking processing features and hardness difference in the HAZ was derived using XAI, providing a better quantitative insight into the FSW process.",
    "The results presented in this work also have a few limitations. Firstly, FSW is a highly complex welding process influenced by more than 20 features. Due to limitations of the available data in reported literature, only five key processing features based on our domain knowledge were considered. Although a smaller set of features makes the ML model simple, a limited number of features restricts the prediction accuracy to ~90 % only. Secondly, the initial dataset used in this study included 32 samples, which increased the prediction bias during the adaptive design process. When a ML model is trained on a small dataset, the model may not learn sufficient information from the data, leading to strong model performance on the training data but increasing errors in predictions of the unseen test data. Lastly, data clustering was observed in the original dataset, which prevented the Kendall correlation coefficient analysis from fully capturing the underlying physics of the FSW process. The accuracy of an ML model depends on both the features provided as input as well as distribution of data within the feature space. When data clustering occurs, the training data may not accurately represent the distribution of the entire dataset, elevating the uncertainty in predicting the test data.",
]


WELDING_FINAL_ABSTRACT = """Friction stir welding of precipitation-strengthened aluminum alloys presents a significant hardness difference between the heat-affected zone (HAZ) and the base metal (BM). A wide range of processing, materials, machine and tool parameters (such as tool rotational rate, tool traverse speed, and tool geometry, especially shoulder diameter), are known to affect the HAZ hardness, making it difficult to navigate the vast combinatorial parameter search space for optimum HAZ hardness. In this study, machine learning (ML) techniques were applied to identify the key processing parameters that affect the hardness difference between the HAZ and BM in AA7075-T6 alloy. Later, an adaptive design strategy was implemented to iteratively determine the optimal combination of processing parameters needed to obtain a targeted hardness difference in the HAZ. Initial data was collected from our random experiments as well as data within reported literature. The initial dataset was then used to train adaptive design tools, and the ML predicted processing parameters were validated via subsequent experiments. It was demonstrated that the ML-guided approach required the generation of only six experimental data points to attain a hardness difference of 41.00 HV 0.3 in the HAZ, compared to trial-and-error based conventional method in which 32 data points were needed to obtain a HAZ hardness difference of 42.00 HV 0.3, thereby significantly increasing the efficiency of process optimization for a targeted mechanical property. Finally, high fidelity mathematical equations were extracted from the data to predict hardness difference in the HAZ as a function of processing parameters using explainable artificial intelligence."""

WELDING_FINAL_INTRO = """The Welding Institute (TWI) in the United Kingdom developed a solid-state joining method known as Friction Stir Welding (FSW) in 1991, which was initially used for aluminum alloys. In its simplest form, a non-consumable rotating tool is inserted into the adjoining edges of the sheets or plates to be joined and moved along the joint line. The tool serves two main functions: (i) heating of the workpiece through the friction between the tool and the workpiece in addition to adiabatic plastic deformation, and (ii) facilitating material flow in the welded region to form the joint. This localized heating softens the material surrounding the tool, and the combination of tool rotation and movement causes the material to move from the front of the pin to the back. The entire process of joining takes place below the melting temperature of the materials used, and results in a ‘solid state’ joint [1]. Mishra et al. developed Friction Stir Processing (FSP) as a universal tool for modifying microstructures, based on the fundamental principles of FSW. In this approach, a rotating tool is inserted into a single-piece workpiece to induce localized microstructural changes to enhance specific properties [2,3]. A schematic diagram of FSW is shown in Fig. 1. FSW is an alternative solution to fusion welding which causes common defects like porosity, hot cracking, and reduced strength in welded areas. This alternative is particularly advantageous for high-strength aluminum alloys (such as the 7xxx series) [3,4]. Aluminum 7075 (AA7075) is one such alloy, which is increasingly valued in both the aerospace and automotive industries for its lightweight properties and excellent mechanical performance. In fact, AA7075 is often comparable to properties of some low-grade steels [5–8]. Its composition consists of Al, 5.1–6.1 % Mg, 2.1–2.5 % Zn, and 1.2–1.6 % Cu, alongside trace elements like Fe, Si, Mn, Ti, and Cr [9]. AA7075 achieves its high strength through the presence of η′ precipitates, which restrict dislocation movement [10–12]."""


def build_welding() -> None:
    srcdir = SRC / "Welding-SR" / "Submission_Materials_Today_Communications"
    doc = Document(srcdir / "Manuscript.docx")
    supplemental = Document(srcdir / "Supplemental material.docx")
    captions = [p.text.strip() for p in Document(srcdir / "Figures" / "Figure captions.docx").paragraphs if p.text.strip()]
    for n in range(1, 11):
        copy_figure(srcdir / "Figures" / f"Figure_{n}.jpg", ROOT / "assets" / "welding" / f"figure-{n}.jpg")

    toc: list[tuple[str, str]] = [("Abstract", "abstract")]
    parts = []
    refs = []
    pending_caption = ""
    inserted = set()
    figure_after = {1: 15, 2: 32, 3: 36, 4: 69, 5: 95, 6: 97, 7: 105, 8: 112, 9: 124, 10: 134}
    paragraph_index = -1

    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "tbl":
            parts.append(word_table(Table(child, doc), pending_caption))
            pending_caption = ""
            continue
        if tag != "p":
            continue
        paragraph_index += 1
        if paragraph_index < 10:
            continue
        p = Paragraph(child, doc)
        raw = p.text.strip()
        html = paragraph_html(p)
        if 127 <= paragraph_index <= 140 or 142 <= paragraph_index <= 152:
            continue
        if raw == "Abstract":
            parts.append('<section class="abstract"><h2 id="abstract">Abstract</h2>')
            continue
        if paragraph_index == 11:
            parts.append(f"<p>{escape(WELDING_FINAL_ABSTRACT)}</p></section>")
            continue
        if paragraph_index == 15:
            html = escape(WELDING_FINAL_INTRO)
        if raw.startswith("Keywords:"):
            parts.append(f'<p class="journal">{html}</p>')
            continue
        if raw.startswith("Table "):
            pending_caption = ("Table 6. Hardness and FSW processing features collected from adaptive design (data point number follows Table 2)."
                               if raw.startswith("Table 3. Hardness and FSW processing parameters collected from adaptive design") else html)
            continue
        if p.style.name == "EndNote Bibliography":
            refs.append(re.sub(r"^\d+\.\s*", "", html))
            continue
        if raw == "References":
            continue
        if p.style.name.startswith("Heading"):
            level = min(4, int(p.style.name.split()[-1]) + 1)
            if raw == "Friction stir welding FSW)":
                raw = "Friction Stir Welding (FSW)"
            ident = slug(raw)
            parts.append(f'<h{level} id="{ident}">{escape(raw)}</h{level}>')
            if level == 2:
                toc.append((raw, ident))
            if paragraph_index == 126:
                parts.append(welding_results_section())
                inserted.add(10)
            if paragraph_index == 141:
                parts.extend(f"<p>{escape(text)}</p>" for text in WELDING_FINAL_CONCLUSIONS)
                parts.append('<h3>CRediT authorship contribution statement</h3><p>Choudhury Samrat: Writing – review &amp; editing, Supervision, Project administration, Methodology, Investigation, Funding acquisition, Formal analysis, Conceptualization. Bhowmik Shubhrodev: Writing – review &amp; editing, Writing – original draft, Methodology, Investigation, Formal analysis, Data curation, Conceptualization. Lu Yizhou: Writing – review &amp; editing, Writing – original draft, Methodology, Investigation, Formal analysis, Data curation, Conceptualization. Kumar Nilesh: Writing – review &amp; editing, Supervision, Project administration, Methodology, Investigation, Formal analysis, Conceptualization. Rhoads Benjamin: Writing – review &amp; editing, Writing – original draft, Methodology, Investigation, Conceptualization.</p>')
                parts.append('<h3>Declaration of Competing Interest</h3><p>The authors declare that they have no known competing financial interests or personal relationships that could have appeared to influence the work reported in this paper.</p>')
                parts.append('<h3>Acknowledgement</h3><p>The author Choudhury was financially supported by the National Science Foundation Award number: 2150816. The XAI work was supported by a Department of Energy, University Nuclear Leadership Program (UNLP) Fellowship for the author Rhoads under an award of DE-NE0009347.</p>')
                parts.append('<h3>Data Availability</h3><p>We have shared the data in this article and the supplemental information.</p>')
        elif html:
            if not raw and p._p.xpath('.//m:oMath'):
                parts.append(f'<div class="equation">{html}</div>')
            else:
                parts.append(f"<p>{html}</p>")
        for n, target in figure_after.items():
            if paragraph_index == target and n not in inserted:
                width, height = image_size(ROOT / "assets" / "welding" / f"figure-{n}.jpg")
                parts.append(f'<figure id="figure-{n}"><img src="assets/welding/figure-{n}.jpg" width="{width}" height="{height}" alt="Figure {n} from the welding study" loading="lazy"><figcaption>{escape(captions[n-1])}</figcaption></figure>')
                inserted.add(n)

    parts.append('<h2 id="references">References</h2><ol class="references">' + "".join(f"<li>{r}</li>" for r in refs) + "</ol>")
    toc.append(("References", "references"))

    # The submitted supplement contains explanatory text and three data tables.
    parts.append('<h2 id="supplemental">Supplemental information</h2>')
    toc.append(("Supplemental", "supplemental"))
    sup_caption = ""
    for child in supplemental.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "tbl":
            parts.append(word_table(Table(child, supplemental), sup_caption))
            sup_caption = ""
        elif tag == "p":
            p = Paragraph(child, supplemental)
            raw = p.text.strip()
            if not raw or raw in ("Supplemental Information",) or raw.startswith(("A machine learning approach", "Yizhou Lu", "1Department", "2Department", "Journal:", "*Corresponding author")):
                continue
            if raw.startswith("Supp. Table"):
                sup_caption = escape(raw)
                continue
            if raw in ("Data", "Methods"):
                parts.append(f'<h3 id="supplemental-{slug(raw)}">{escape(raw)}</h3>')
            else:
                parts.append(f"<p>{paragraph_html(p)}</p>")

    title = "A machine learning approach to mitigating hardness degradation in the heat-affected zone of AA7075-T6 alloy"
    page = shell(title, "Full welding and machine learning research manuscript by Benjamin Rhoads and collaborators.",
                 "Materials Today Communications", "2025",
                 "Yizhou Lu · Shubhrodev Bhowmik · Benjamin Rhoads · Nilesh Kumar · Samrat Choudhury",
                 "46, 112689", "10.1016/j.mtcomm.2025.112689", "\n".join(parts), toc)
    (ROOT / "welding-paper.html").write_text(page, encoding="utf-8")
    print("Welding:", len(refs), "references,", len(inserted), "figures")


XLINK = "{http://www.w3.org/1999/xlink}href"


def xml_inline(e: ET.Element) -> str:
    from latex2mathml.converter import convert as latex_to_mathml
    tag = e.tag.split("}")[-1]
    if tag in ("fig", "table-wrap", "disp-formula"):
        return ""
    if tag == "tex-math":
        match = re.search(r"\$\$(.*?)\$\$", e.text or "", re.S)
        if not match:
            return escape(e.text or "")
        return latex_to_mathml(match.group(1))
    start = escape(e.text or "")
    children = "".join(xml_inline(c) + escape(c.tail or "") for c in e)
    value = start + children
    if tag in ("italic", "bold", "sup", "sub"):
        html_tag = {"italic": "em", "bold": "strong", "sup": "sup", "sub": "sub"}[tag]
        return f"<{html_tag}>{value}</{html_tag}>"
    if tag == "xref":
        rid = e.get("rid", "")
        if rid.startswith("MOESM"):
            return f'<a href="https://doi.org/10.1007/s13346-025-02006-4" target="_blank" rel="noopener noreferrer">{value}</a>'
        target = re.sub(r"^(?:Fig|Tab|Equ)", lambda m: {"Fig": "figure-", "Tab": "table-", "Equ": "equation-"}[m.group()], rid)
        if rid.startswith("CR") or rid.startswith("Ref"):
            target = "ref-" + re.sub(r"\D", "", rid)
        return f'<a href="#{escape(target, quote=True)}">{value}</a>' if target else value
    if tag == "ext-link":
        href = e.get(XLINK, "")
        return f'<a href="{escape(href, quote=True)}" target="_blank" rel="noopener noreferrer">{value}</a>' if href.startswith("https://") else value
    return value


def xml_paragraph(p: ET.Element) -> str:
    return f"<p>{xml_inline(p)}</p>"


def xml_table(e: ET.Element) -> str:
    ident = "table-" + re.sub(r"\D", "", e.get("id", ""))
    caption = xml_inline(e.find("caption")) if e.find("caption") is not None else ""
    table = e.find("table")
    rows = []
    if table is not None:
        for row in table.iter("tr"):
            cells = []
            for cell in row:
                tag = cell.tag.split("}")[-1]
                if tag not in ("th", "td"):
                    continue
                span = "".join(f' {name}="{cell.get(name)}"' for name in ("colspan", "rowspan") if cell.get(name) and cell.get(name) != "1")
                cells.append(f"<{tag}{span}>{xml_inline(cell)}</{tag}>")
            rows.append("<tr>" + "".join(cells) + "</tr>")
    return f'<div class="table-scroll" id="{ident}"><table><caption>{escape(e.findtext("label") or "")} — {caption}</caption><tbody>{"".join(rows)}</tbody></table></div>'


def xml_figure(e: ET.Element) -> str:
    n = re.sub(r"\D", "", e.get("id", ""))
    caption = xml_inline(e.find("caption")) if e.find("caption") is not None else ""
    width, height = image_size(ROOT / "assets" / "bioprinting" / f"figure-{n}.jpg")
    return f'<figure id="figure-{n}"><img src="assets/bioprinting/figure-{n}.jpg" width="{width}" height="{height}" alt="Figure {n} from the final published bioprinting paper" loading="lazy"><figcaption>Figure {n}. {caption}</figcaption></figure>'


def extract_bio_figures() -> None:
    pdf = PdfReader(SRC / "ML_Optimized_Bio_3D_Printing" / "final_paper.pdf")
    pages = {1: (4, 0), 2: (7, 0), 3: (8, 0), 4: (9, 0), 5: (10, 0), 6: (11, 1), 7: (11, 0), 8: (12, 0)}
    outdir = ROOT / "assets" / "bioprinting"
    outdir.mkdir(parents=True, exist_ok=True)
    for n, (page, index) in pages.items():
        image = pdf.pages[page - 1].images[index].image.convert("RGB")
        if image.width > 2055:
            image.thumbnail((2055, 2055))
        image.save(outdir / f"figure-{n}.jpg", "JPEG", quality=91, optimize=True)


def build_bio() -> None:
    root = ET.parse(ROOT / "sources" / "bio-final.xml").getroot()
    extract_bio_figures()
    toc: list[tuple[str, str]] = [("Abstract", "abstract")]
    parts = []
    used_ids = {"abstract"}
    abstract = root.find(".//abstract")
    if abstract is not None:
        parts.append('<section class="abstract"><h2 id="abstract">Abstract</h2>' + "".join(xml_paragraph(p) for p in abstract.findall("p")) + "</section>")
    keywords = ["".join(x.itertext()).strip() for x in root.findall('.//front//kwd-group/kwd')]
    if keywords:
        parts.append('<p class="journal"><strong>Keywords:</strong> ' + escape(' · '.join(keywords)) + '</p>')

    def block(e: ET.Element, depth: int = 0) -> None:
        tag = e.tag.split("}")[-1]
        if tag == "sec":
            title = e.find("title")
            label = xml_inline(title) if title is not None else "Section"
            plain = "".join(title.itertext()).strip() if title is not None else "Section"
            ident = slug(plain)
            if ident in used_ids:
                ident += "-" + (e.get("id") or str(len(used_ids))).lower()
            used_ids.add(ident)
            level = min(4, 2 + depth)
            parts.append(f'<h{level} id="{ident}">{label}</h{level}>')
            if depth == 0:
                toc.append((plain, ident))
            for child in e:
                if child.tag.split("}")[-1] != "title":
                    block(child, depth + 1 if child.tag.split("}")[-1] == "sec" else depth)
        elif tag == "p":
            fragments = [escape(e.text or "")]
            for child in e:
                child_tag = child.tag.split("}")[-1]
                if child_tag in ("fig", "table-wrap", "disp-formula"):
                    value = "".join(fragments).strip()
                    if value:
                        parts.append(f"<p>{value}</p>")
                    fragments = []
                    block(child, depth)
                else:
                    fragments.append(xml_inline(child))
                fragments.append(escape(child.tail or ""))
            value = "".join(fragments).strip()
            if value:
                parts.append(f"<p>{value}</p>")
        elif tag == "fig":
            parts.append(xml_figure(e))
        elif tag == "table-wrap":
            parts.append(xml_table(e))
        elif tag == "disp-formula":
            number = escape(e.findtext("label") or "")
            tex = e.find("tex-math")
            math = xml_inline(tex) if tex is not None else escape("".join(e.itertext()))
            parts.append(f'<div class="equation" id="equation-{number}">{math}<span class="equation-number">Equation {number}</span></div>')
        elif tag == "supplementary-material":
            parts.append('<p>Supplementary material is available with the <a href="https://doi.org/10.1007/s13346-025-02006-4" target="_blank" rel="noopener noreferrer">published article</a>.</p>')

    body = root.find("body")
    for child in body:
        block(child)
    back = root.find("back")
    for child in back:
        tag = child.tag.split("}")[-1]
        if tag == "ref-list":
            parts.append('<h2 id="references">References</h2>')
            toc.append(("References", "references"))
            entries = []
            for ref in child.findall("ref"):
                n = re.sub(r"\D", "", ref.get("id", ""))
                cit = ref.find(".//mixed-citation")
                if cit is None:
                    cit = ref.find(".//element-citation")
                if cit is not None and cit.tag == "mixed-citation":
                    citation = (cit.text or "") + "".join(
                        "".join(part.itertext()) + (part.tail or "")
                        for part in cit if part.tag != "pub-id"
                    )
                    value = escape(citation.strip().replace("�", "–"))
                    doi = next((item.text for item in cit.findall("pub-id") if item.get("pub-id-type") == "doi"), None)
                    if doi:
                        value += f' <a href="https://doi.org/{escape(doi, quote=True)}" target="_blank" rel="noopener noreferrer">doi:{escape(doi)}</a>'
                else:
                    value = xml_inline(cit) if cit is not None else escape("".join(ref.itertext()))
                entries.append(f'<li id="ref-{n}">{value}</li>')
            parts.append('<ol class="references">' + "".join(entries) + "</ol>")
        elif tag in ("ack", "notes", "fn-group"):
            title = child.findtext("title") or ("Publisher note" if tag == "fn-group" else "Additional information")
            ident = slug(title)
            parts.append(f'<h2 id="{ident}">{escape(title)}</h2>')
            for p in child.iter("p"):
                parts.append(xml_paragraph(p))

    title = "Optimization of print parameters for batch and continuous manufacturing of three-dimensional (3D) printed dosage forms using artificial intelligence and machine learning"
    page = shell(title, "Full published research article on AI and machine learning for 3D printed pharmaceutical dosage forms by Benjamin Rhoads and collaborators.",
                 "Drug Delivery and Translational Research", "2026",
                 "Kshitij Chitnis · Yizhou Lu · Benjamin Rhoads · Leela Raghava Jaidev Chakka · Samrat Choudhury · Mohammed Maniruzzaman",
                 "16, 2073–2088", "10.1007/s13346-025-02006-4", "\n".join(parts), toc)
    (ROOT / "bioprinting-paper.html").write_text(page, encoding="utf-8")
    print("Bioprinting:", len(root.findall('.//body//fig')), "figures,", len(root.findall('.//body//table-wrap')), "tables,", len(root.findall('.//ref-list/ref')), "references")


if __name__ == "__main__":
    build_welding()
    build_bio()
