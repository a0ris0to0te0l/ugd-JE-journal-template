# Journal of Economics (UGD) paper template

Template for papers to the *Journal of Economics* of Goce Delcev University, Stip, following its [author guidelines](https://js.ugd.edu.mk/index.php/JE/about/submissions#authorGuidelines). The guidelines are published as three images on that page: a sample paper with the format of each element marked on it.

| File | What it is |
|---|---|
| `js-ugd-template.docx` | Word template. The journal asks for Word files |
| `template.tex`, `js-ugd.sty`, `references.bib` | LaTeX template |
| `js-ugd-sample.pdf` | `template.tex` compiled |
| `figure-placeholder.png` | Example figure used by both |
| `tools/build_docx.py` | Rebuilds the `.docx` (`pip install python-docx`) |

## Word

Copy the file, rename it and replace the text. Every format is a style: Heading 1 and Heading 2 for numbered headings, and styles starting with "JE" for the rest. Heading, reference, table and figure numbers are automatic; after moving tables or figures, select all and press F9 to update them.

For blind review, delete the author and affiliation lines, then clear the author name with File → Info → Check for Issues → Inspect Document.

## LaTeX

Copy the files into the paper's folder, rename `template.tex` and compile with `latexmk -pdf`. This needs pdflatex, biber, biblatex-apa and the Arial package `uarial`. For blind review, use `\usepackage[anonymous]{js-ugd}`. Citations in the text that reveal the authors still have to be changed by hand.

The journal's submission checklist asks for a Word file, so a LaTeX paper still needs converting before it is submitted.

## Formats

Spacing is before/after in points, as marked in the guidelines.

| Element | Spacing | Font |
|---|---|---|
| Title | 0/15 | 20 bold |
| Authors | 0/10 | 16 italic |
| Affiliations | 0/10 | 11 italic |
| Abstract, 250-500 words, on the first page | 20/10 | 11 bold |
| Keywords heading | 20/2 | 12 bold |
| Keywords, up to 5, alphabetical | 0/2 | 11 |
| Section heading | 20/10 | 14 bold |
| Subsection heading | 20/10 | 12 bold italic |
| Body text, no indent, justified | 0/2 | 11 |
| Table caption, centred | 10/10 | 10 |
| Figure caption, centred, after the figure | 5/10 | 10 |
| References, numbered, APA, 0.5 cm hanging | 0/0 | 10 |

Font: Arial. Page: A4, margins 2.5 cm, header 1.25 cm, footer 1 cm. Length: 8-12 pages.

## Where the journal's text is unclear

| Point | Guidelines | Used here |
|---|---|---|
| Font size | Checklist: 12 pt. Sample paper: Arial 11 | Arial 11 |
| Table caption | "above" and "after the table" | Above, as in the sample table |
| Heading case | "small caps" | Title case, as in the sample |
| Bullets | Text: bullet at 0.7 cm. Sample: text at about 0.7 cm | Bullet at 0.7 cm, text at 1.2 cm |
| Table text | A special style, size not given | 10 pt, as the sample looks. Use 9 pt for wide tables |
| Table rules | "only horizontal rules" | A rule under every row, as in the sample |
| Equations | Not covered | Centred, numbered at the right |
| "et al." | "(Lazarov, et al., 2023)" | Standard APA, "(Lazarov et al., 2023)" |
| Journal article format | The example repeats the proceedings format | Standard APA |
| Format | Checklist: Word. Sample: Word and PDF | Both are available here |

The bullet, heading case, caption and font-size choices are the ones used in the Model A paper for this journal (`Master Thesis/Model_A_Macro/Publication at Journal UGD`).

The example references are real works. Five come from that paper's bibliography, whose DOIs were checked against Crossref; the book and the webpage were added for this template.
