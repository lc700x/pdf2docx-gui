import os
import struct
import tempfile
import unittest
import zipfile
import zlib

from docx import Document

from latex_to_docx import (
    _create_reference_docx,
    _normalize_minipage_tables,
    _restore_table_captions_and_missing_tables,
    _table_float_records,
    _wide_subfigure_linewidth_images,
    convert_latex_to_docx,
)


def _png_chunk(kind, data):
    body = kind + data
    return (struct.pack(">I", len(data)) + body
            + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF))


def _write_test_png(path):
    width, height = 300, 200
    header = struct.pack(">2I5B", width, height, 8, 6, 0, 0, 0)
    row = b"\x00" + b"\x33\x99\xff\xff" * width
    pixels = zlib.compress(row * height)
    with open(path, "wb") as image_file:
        image_file.write(b"\x89PNG\r\n\x1a\n")
        image_file.write(_png_chunk(b"IHDR", header))
        image_file.write(_png_chunk(b"IDAT", pixels))
        image_file.write(_png_chunk(b"IEND", b""))


class LatexConversionTests(unittest.TestCase):
    def test_wide_subfigure_widths_ignore_commented_graphics(self):
        source = r"""\begin{subfigure}[t]{0.32\textwidth}
\includegraphics[width=\linewidth]{small.png}
\end{subfigure}
% \begin{subfigure}[t]{0.96\textwidth}
% \includegraphics[width=\linewidth]{commented.png}
% \end{subfigure}
\begin{subfigure}[t]{0.96\textwidth}
\includegraphics[width=\linewidth]{wide.png}
\end{subfigure}"""
        image_count, widths = _wide_subfigure_linewidth_images(source)
        self.assertEqual(image_count, 2)
        self.assertEqual(widths, [(1, 0.96)])

    def test_minipage_normalization_ignores_commented_environments(self):
        commented = r"""\begin{table}
% \begin{minipage}{\linewidth}
% \begin{tabularx}{\textwidth}{X}
% \end{tabularx}
% \end{minipage}
\begin{tabular}{l}alpha \\
\end{tabular}
\end{table}"""
        self.assertEqual(_normalize_minipage_tables(commented), commented)

        active = r"""\begin{table}
\begin{minipage}{\linewidth}
\begin{tabularx}{\textwidth}{X}
alpha \\
\end{tabularx}
\end{minipage}
\end{table}"""
        normalized = _normalize_minipage_tables(active)
        self.assertIn(r"\begin{tabular}{l}", normalized)
        self.assertNotIn(r"\begin{minipage}", normalized)
        self.assertNotIn(r"\begin{tabularx}", normalized)

    def test_recovers_native_table_when_pandoc_kept_only_its_caption(self):
        source = r"""\documentclass{article}
\begin{document}
\begin{table}
\caption{Unique table content}
\begin{tabular}{ll}
Header Alpha & Header Beta \\
Rare Zebra Marker & 42 \\
\end{tabular}
\begin{tabular}{ll}
Second Header Alpha & Second Header Beta \\
Rare Otter Marker & 17 \\
\end{tabular}
\end{table}
\end{document}"""
        with tempfile.TemporaryDirectory() as directory:
            source_path = os.path.join(directory, "main.tex")
            reference_path = os.path.join(directory, "reference.docx")
            output_path = os.path.join(directory, "output.docx")
            with open(source_path, "w", encoding="utf-8") as source_file:
                source_file.write(source)
            _create_reference_docx(source, reference_path)

            document = Document()
            caption = document.add_paragraph("Table 1. Unique table content")
            caption.style = "Caption"
            unrelated = document.add_table(rows=2, cols=2)
            unrelated.cell(0, 0).text = "Other heading"
            unrelated.cell(0, 1).text = "Different heading"
            unrelated.cell(1, 0).text = "Unrelated value"
            unrelated.cell(1, 1).text = "Another value"
            document.save(output_path)

            _restore_table_captions_and_missing_tables(
                source, _table_float_records(source), output_path, directory,
                reference_path, directory)

            converted = Document(output_path)
            self.assertEqual(len(converted.tables), 3)
            self.assertEqual(
                sum(p.text.startswith("Table 1.") for p in converted.paragraphs), 1)
            self.assertTrue(any(
                "Rare Zebra Marker" in cell.text
                for table in converted.tables
                for row in table.rows for cell in row.cells))
            self.assertTrue(any(
                "Rare Otter Marker" in cell.text
                for table in converted.tables
                for row in table.rows for cell in row.cells))

    def test_conversion_keeps_text_tables_math_figures_and_citations(self):
        with tempfile.TemporaryDirectory() as directory:
            image_dir = os.path.join(directory, "figures")
            os.mkdir(image_dir)
            image_path = os.path.join(image_dir, "pixel.png")
            _write_test_png(image_path)

            source_path = os.path.join(directory, "main.tex")
            with open(source_path, "w", encoding="utf-8") as source_file:
                source_file.write(r"""\documentclass[12pt,a4paper]{article}
\usepackage[margin=1in]{geometry}
\usepackage{graphicx}
\usepackage{subcaption}
\usepackage{natbib}
\begin{document}
\section{Editable section}
An editable paragraph with inline math $x^2$, citation \citep{smith2020}, and 50\% coverage.
\begin{equation}x = 2\end{equation}
\begin{table}
\caption{Numeric results}
\begin{tabular}{lc}
Name & Value \\
alpha & 3 \\
\end{tabular}
\end{table}
\begin{figure}
\includegraphics[
% TeX comments inside a graphics option must not break Pandoc parsing.
width=0.5\textwidth
]{figures/pixel.png}
\caption{Pixel plot}
\end{figure}
\begin{figure}
\begin{subfigure}[t]{0.96\textwidth}
\includegraphics[width=\linewidth]{figures/pixel.png}
\caption{Wide pixel plot}
\end{subfigure}
\end{figure}
\bibliographystyle{plainnat}
\bibliography{references}
\end{document}
""")
            with open(os.path.join(directory, "references.bib"), "w",
                      encoding="utf-8") as bibliography:
                bibliography.write("""@article{smith2020,
  author = {Jane Smith},
  title = {A Useful Paper},
  journal = {Test Journal},
  year = {2020}
}
""")

            destination_path = os.path.join(directory, "main_from_tex.docx")
            statuses = []
            with open(source_path, "rb") as source_file:
                original_source = source_file.read()
            result = convert_latex_to_docx(
                source_path, destination_path, on_status=statuses.append)

            self.assertEqual(result, destination_path)
            with open(source_path, "rb") as source_file:
                self.assertEqual(source_file.read(), original_source)
            self.assertEqual(statuses, ["Converting LaTeX to Word…"])
            document = Document(destination_path)
            document_text = "\n".join(
                [paragraph.text for paragraph in document.paragraphs]
                + [cell.text for table in document.tables
                   for row in table.rows for cell in row.cells])
            self.assertIn("Editable section", document_text)
            self.assertIn("editable paragraph", document_text.lower())
            self.assertIn("50% coverage", document_text)
            self.assertIn("A Useful Paper", document_text)
            self.assertIn("Pixel plot", document_text)
            self.assertIn("Wide pixel plot", document_text)
            self.assertIn("Table 1. Numeric results", document_text)
            figure_captions = [
                paragraph for paragraph in document.paragraphs
                if paragraph.style.name == "Image Caption"]
            self.assertTrue(any(
                paragraph.text == "Figure 1. Pixel plot"
                for paragraph in figure_captions))
            self.assertTrue(any(
                paragraph.runs and paragraph.runs[0].text == "Figure 1."
                and paragraph.runs[0].bold
                for paragraph in figure_captions))
            table_caption = next(
                paragraph for paragraph in document.paragraphs
                if paragraph.text.startswith("Table 1."))
            self.assertEqual(table_caption.style.name, "Table Caption")
            heading = next(
                paragraph for paragraph in document.paragraphs
                if paragraph.text.endswith("Editable section"))
            self.assertEqual(heading.style.name, "Heading 1")
            self.assertEqual(len(document.tables), 2)
            self.assertTrue(any(
                any(cell.text == "alpha" for row in table.rows for cell in row.cells)
                for table in document.tables))
            self.assertEqual(len(document.inline_shapes), 2)
            self.assertGreater(document.inline_shapes[0].width, 0)
            self.assertGreater(document.inline_shapes[0].height, 0)
            expected_width = (210 / 25.4 - 2) / 2
            self.assertAlmostEqual(
                document.inline_shapes[0].width / 914400,
                expected_width, delta=0.1)
            self.assertAlmostEqual(
                document.inline_shapes[0].width
                / document.inline_shapes[0].height, 1.5, places=2)
            self.assertAlmostEqual(
                document.inline_shapes[1].width / 914400,
                expected_width * 2 * 0.96, delta=0.1)
            self.assertAlmostEqual(
                document.sections[0].page_width / 914400, 210 / 25.4, places=1)
            self.assertAlmostEqual(
                document.sections[0].left_margin / 914400, 1.0, places=2)
            self.assertEqual(document.styles["Normal"].font.name, "Times New Roman")
            self.assertEqual(document.styles["Normal"].font.size.pt, 12)
            self.assertEqual(
                document.styles["Image Caption"].font.size.pt, 10)
            self.assertEqual(
                document.styles["Heading 1"].paragraph_format.space_before.pt,
                12)
            with zipfile.ZipFile(destination_path) as archive:
                xml = archive.read("word/document.xml").decode("utf-8")
                self.assertIn("<m:oMath", xml)

    def test_missing_source_reports_clear_error(self):
        with tempfile.TemporaryDirectory() as directory:
            source_path = os.path.join(directory, "missing.tex")
            destination_path = os.path.join(directory, "missing.docx")
            with self.assertRaisesRegex(FileNotFoundError, "source file not found"):
                convert_latex_to_docx(source_path, destination_path)


if __name__ == "__main__":
    unittest.main()
