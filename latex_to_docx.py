"""Convert a LaTeX source document to an editable Word document with Pandoc."""

import os
import re
import subprocess
import tempfile
from copy import deepcopy
from difflib import SequenceMatcher

import pypandoc
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Mm, Pt
from docx.text.paragraph import Paragraph


_BIBLIOGRAPHY_PATTERN = re.compile(
    r"\\(?:bibliography\b|addbibresource\b)\s*"
    r"(?:\[[^\]]*\]\s*)?\{([^}]+)\}"
)
_ENVIRONMENT_PATTERN = re.compile(r"\\(begin|end)\s*\{([^}]+)\}")
_CAPTION_PATTERN = re.compile(r"\\caption(?![A-Za-z])(\*)?")
_LABEL_PATTERN = re.compile(r"\\label\s*\{([^}]+)\}")
_SECTION_COMMAND_PATTERN = re.compile(
    r"\\(subparagraph|subsubsection|subsection|paragraph|chapter|section|part)"
    r"(\*)?(?![A-Za-z])")

_PAPER_SIZES = {
    "a4paper": (Mm(210), Mm(297)),
    "a5paper": (Mm(148), Mm(210)),
    "letterpaper": (Inches(8.5), Inches(11)),
    "legalpaper": (Inches(8.5), Inches(14)),
    "executivepaper": (Mm(184.15), Mm(266.7)),
}
_LENGTH_PATTERN = re.compile(r"^\s*(-?\d+(?:\.\d+)?)\s*(in|inch|cm|mm|pt|bp)?\s*$", re.I)


def _strip_tex_comments(text):
    """Remove TeX comments while preserving escaped percent signs and lines."""
    output = []
    index = 0
    while index < len(text):
        if text[index] == "%":
            slashes = 0
            previous = index - 1
            while previous >= 0 and text[previous] == "\\":
                slashes += 1
                previous -= 1
            if slashes % 2 == 0:
                newline = text.find("\n", index)
                output.append(" ")
                index = len(text) if newline < 0 else newline
                continue
        output.append(text[index])
        index += 1
    return "".join(output)


def _length_in_inches(value):
    match = _LENGTH_PATTERN.match(value)
    if not match:
        return None
    amount = float(match.group(1))
    unit = (match.group(2) or "pt").lower()
    factors = {
        "in": 1,
        "inch": 1,
        "cm": 1 / 2.54,
        "mm": 1 / 25.4,
        "pt": 1 / 72,
        "bp": 1 / 72.27,
    }
    return amount * factors[unit]


def _layout_from_source(source_text):
    """Read common paper-size and margin settings from a LaTeX preamble."""
    active_source = _strip_tex_comments(source_text)
    document_class = re.search(
        r"\\documentclass(?:\s*\[([^\]]*)\])?\s*\{[^}]+\}",
        active_source,
    )
    options = (document_class.group(1).split(",")
               if document_class and document_class.group(1) else [])
    options = [option.strip().lower() for option in options]
    page_size = next(
        (_PAPER_SIZES[option] for option in options if option in _PAPER_SIZES),
        _PAPER_SIZES["letterpaper"],
    )
    margins = {side: 1.0 for side in ("left", "right", "top", "bottom")}

    geometry_chunks = []
    geometry_package = re.search(
        r"\\usepackage(?:\s*\[([^\]]*)\])?\s*\{[^}]*\bgeometry\b[^}]*\}",
        active_source,
    )
    if geometry_package and geometry_package.group(1):
        geometry_chunks.append(geometry_package.group(1))
    geometry_chunks.extend(re.findall(r"\\geometry\s*\{([^}]*)\}", active_source))

    for chunk in geometry_chunks:
        for option in chunk.split(","):
            key, separator, value = option.partition("=")
            key = key.strip().lower()
            value = value.strip()
            if not separator:
                if key in _PAPER_SIZES:
                    page_size = _PAPER_SIZES[key]
                continue
            inches = _length_in_inches(value)
            if inches is None:
                continue
            if key == "margin":
                margins.update({side: inches for side in margins})
            elif key in margins:
                margins[key] = inches
            elif key == "hmargin":
                margins["left"] = margins["right"] = inches
            elif key == "vmargin":
                margins["top"] = margins["bottom"] = inches

    return page_size, margins


def _set_style_font(style, size, bold=None, italic=None):
    style.font.name = "Times New Roman"
    style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    if italic is not None:
        style.font.italic = italic
    style._element.get_or_add_rPr().get_or_add_rFonts().set(
        qn("w:eastAsia"), "Times New Roman")


def _style_by_name(document, name):
    return next(style for style in document.styles if style.name == name)


def _ensure_table_caption_style(document):
    try:
        return _style_by_name(document, "Table Caption")
    except StopIteration:
        style = document.styles.add_style("Table Caption", WD_STYLE_TYPE.PARAGRAPH)
        try:
            style.base_style = _style_by_name(document, "Caption")
        except StopIteration:
            style.base_style = _style_by_name(document, "Normal")
        _set_style_font(style, 10, italic=False)
        style.paragraph_format.line_spacing = 1
        style.paragraph_format.space_after = Pt(8)
        style.paragraph_format.keep_with_next = True
        return style


def _create_reference_docx(source_text, destination_path):
    """Create a Pandoc reference file using the source paper's basic layout."""
    pandoc = pypandoc.get_pandoc_path()
    result = subprocess.run(
        [pandoc, "--print-default-data-file", "reference.docx"],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    with open(destination_path, "wb") as reference_file:
        reference_file.write(result.stdout)

    document = Document(destination_path)
    page_size, margins = _layout_from_source(source_text)
    for section in document.sections:
        section.page_width, section.page_height = page_size
        section.left_margin = Inches(margins["left"])
        section.right_margin = Inches(margins["right"])
        section.top_margin = Inches(margins["top"])
        section.bottom_margin = Inches(margins["bottom"])

    normal = _style_by_name(document, "Normal")
    _set_style_font(normal, 12)
    normal.paragraph_format.line_spacing = None

    for name in ("Body Text", "First Paragraph", "Bibliography", "Abstract"):
        try:
            style = _style_by_name(document, name)
        except StopIteration:
            continue
        _set_style_font(style, 12)
        style.paragraph_format.line_spacing = 1.5
        style.paragraph_format.space_before = None
        style.paragraph_format.space_after = Pt(6)
    try:
        _style_by_name(document, "Abstract").paragraph_format.keep_with_next = True
    except StopIteration:
        pass
    try:
        abstract_title = _style_by_name(document, "Abstract Title")
        _set_style_font(abstract_title, 12, bold=True)
        abstract_title.paragraph_format.space_before = Pt(15)
        abstract_title.paragraph_format.space_after = None
        abstract_title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        abstract_title.paragraph_format.keep_with_next = True
    except StopIteration:
        pass
    try:
        compact = _style_by_name(document, "Compact")
        _set_style_font(compact, 12)
        compact.paragraph_format.space_before = Pt(1.8)
        compact.paragraph_format.space_after = Pt(1.8)
    except StopIteration:
        pass

    title = _style_by_name(document, "Title")
    _set_style_font(title, 17, bold=True)
    title.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.line_spacing = 1
    for name, size in (("Heading 1", 14), ("Heading 2", 12), ("Heading 3", 12)):
        heading = _style_by_name(document, name)
        _set_style_font(heading, size, bold=True, italic=False)
        heading.paragraph_format.space_before = Pt(12)
        heading.paragraph_format.space_after = Pt(6)
    for name in ("Table Caption", "Image Caption"):
        try:
            caption = _style_by_name(document, name)
        except StopIteration:
            caption = document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        _set_style_font(caption, 10, italic=False)
        caption.paragraph_format.line_spacing = 1
        caption.paragraph_format.space_before = None
        caption.paragraph_format.space_after = Pt(8)
    _style_by_name(document, "Table Caption").paragraph_format.keep_with_next = True

    for name in ("Author", "Date"):
        try:
            _set_style_font(_style_by_name(document, name), 12)
        except KeyError:
            pass
    document.save(destination_path)


def _is_comment_offset(text, offset):
    line_start = text.rfind("\n", 0, offset) + 1
    index = line_start
    while index < offset:
        if text[index] == "%":
            slash_count = 0
            previous = index - 1
            while previous >= line_start and text[previous] == "\\":
                slash_count += 1
                previous -= 1
            if slash_count % 2 == 0:
                return True
        index += 1
    return False


def _wide_subfigure_ranges(text):
    """Find subfigures whose image width should follow a wide line width."""
    environments = [
        match for match in _ENVIRONMENT_PATTERN.finditer(text)
        if not _is_comment_offset(text, match.start())
    ]
    ranges = []
    for index, match in enumerate(environments):
        if match.groups() != ("begin", "subfigure"):
            continue
        cursor = match.end()
        if cursor < len(text) and text[cursor] == "[":
            _, cursor = _read_tex_group(text, cursor, "[", "]")
        width, _ = _read_tex_group(text, cursor, "{", "}")
        ratio = re.fullmatch(
            r"\s*(?:(\d+(?:\.\d+)?)\s*)?\\textwidth\s*", width or "")
        ratio_value = float(ratio.group(1) or 1) if ratio else 0
        if not ratio or ratio_value < 0.5:
            continue

        depth = 1
        for closing in environments[index + 1:]:
            if closing.group(2) != "subfigure":
                continue
            depth += 1 if closing.group(1) == "begin" else -1
            if not depth:
                ranges.append((match.start(), closing.end(), str(ratio_value)))
                break
    return ranges


def _wide_subfigure_linewidth_images(text):
    """Map source image positions to widths lost by Pandoc's subfigure reader."""
    subfigures = _wide_subfigure_ranges(text)
    widths = []
    image_index = 0
    for match in re.finditer(r"\\includegraphics\*?", text):
        if _is_comment_offset(text, match.start()):
            continue
        cursor = match.end()
        options = None
        while cursor < len(text) and text[cursor].isspace():
            cursor += 1
        if cursor < len(text) and text[cursor] == "[":
            options, _ = _read_tex_group(text, cursor, "[", "]")
        containing = [item for item in subfigures
                      if item[0] < match.start() < item[1]]
        if (options and containing
                and re.search(r"\bwidth\s*=\s*\\linewidth\b", options)):
            width = max(containing, key=lambda item: item[0])[2]
            widths.append((image_index, float(width)))
        image_index += 1
    return image_index, widths


def _normalize_includegraphics_options(text):
    """Flatten commented, multiline includegraphics options for Pandoc's reader."""
    pattern = re.compile(r"\\includegraphics\*?\s*\[")
    output = []
    cursor = 0
    for match in pattern.finditer(text):
        if match.start() < cursor or _is_comment_offset(text, match.start()):
            continue

        options = []
        square_depth = 1
        brace_depth = 0
        index = match.end()
        while index < len(text):
            character = text[index]
            if character == "%":
                slash_count = 0
                previous = index - 1
                while previous >= match.end() and text[previous] == "\\":
                    slash_count += 1
                    previous -= 1
                if slash_count % 2 == 0:
                    newline = text.find("\n", index)
                    if newline < 0:
                        index = len(text)
                        break
                    options.append(" ")
                    index = newline + 1
                    continue
            if character == "\\" and index + 1 < len(text):
                options.append(text[index:index + 2])
                index += 2
                continue
            if character == "{":
                brace_depth += 1
            elif character == "}" and brace_depth:
                brace_depth -= 1
            elif not brace_depth and character == "[":
                square_depth += 1
            elif not brace_depth and character == "]":
                square_depth -= 1
                if not square_depth:
                    break
            options.append(" " if character in "\r\n\t" else character)
            index += 1

        if square_depth:
            continue
        output.extend((text[cursor:match.start()], text[match.start():match.end()],
                       "".join(options).strip(), "]"))
        cursor = index + 1

    if not output:
        return text
    output.append(text[cursor:])
    return "".join(output)


def _read_tex_group(text, index, opening, closing):
    while index < len(text) and text[index].isspace():
        index += 1
    if index >= len(text) or text[index] != opening:
        return None, index

    depth = 1
    content = []
    index += 1
    while index < len(text) and depth:
        character = text[index]
        if character == "%":
            slashes = 0
            previous = index - 1
            while previous >= 0 and text[previous] == "\\":
                slashes += 1
                previous -= 1
            if slashes % 2 == 0:
                newline = text.find("\n", index)
                content.append(" ")
                index = len(text) if newline < 0 else newline
                continue
        if character == "\\" and index + 1 < len(text):
            content.append(text[index:index + 2])
            index += 2
            continue
        if character == opening:
            depth += 1
        elif character == closing:
            depth -= 1
            if not depth:
                break
        content.append(character)
        index += 1
    if depth:
        return None, index
    return "".join(content), index + 1


def _table_float_ranges(text):
    ranges = []
    stack = []
    for match in _ENVIRONMENT_PATTERN.finditer(text):
        if _is_comment_offset(text, match.start()):
            continue
        operation, environment = match.groups()
        if operation == "begin":
            if environment in ("table", "table*") and not any(
                    entry[0] in ("table", "table*") for entry in stack):
                stack.append((environment, match.start()))
            else:
                stack.append((environment, None))
        elif stack:
            matching = next((index for index in range(len(stack) - 1, -1, -1)
                             if stack[index][0] == environment), None)
            if matching is None:
                continue
            open_environment, start = stack[matching]
            del stack[matching:]
            if open_environment in ("table", "table*") and start is not None:
                ranges.append((start, match.end()))
    return ranges


def _named_environment_ranges(text, names):
    ranges = []
    stack = []
    for match in _ENVIRONMENT_PATTERN.finditer(text):
        if _is_comment_offset(text, match.start()):
            continue
        operation, environment = match.groups()
        if operation == "begin":
            start = match.start() if environment in names and not any(
                item[0] in names and item[1] is not None for item in stack) else None
            stack.append((environment, start))
        elif stack:
            matching = next((index for index in range(len(stack) - 1, -1, -1)
                             if stack[index][0] == environment), None)
            if matching is None:
                continue
            open_environment, start = stack[matching]
            del stack[matching:]
            if open_environment in names and start is not None:
                ranges.append((start, match.end()))
    return ranges


def _figure_caption_records(text):
    """Return numbered figure captions whose labels Word/Pandoc omits."""
    records = []
    figure_number = 0
    for start, end in _named_environment_ranges(text, {"figure", "figure*"}):
        block = text[start:end]
        subfigures = _named_environment_ranges(block, {"subfigure"})
        captions = [match for match in _CAPTION_PATTERN.finditer(block)
                    if not _is_comment_offset(block, match.start())]
        top_level = [match for match in captions
                     if not match.group(1)
                     and not any(left < match.start() < right
                                 for left, right in subfigures)]
        for match in top_level:
            figure_number += 1
            cursor = match.end()
            while cursor < len(block) and block[cursor].isspace():
                cursor += 1
            if cursor < len(block) and block[cursor] == "[":
                _, cursor = _read_tex_group(block, cursor, "[", "]")
            caption, _ = _read_tex_group(block, cursor, "{", "}")
            if caption is None:
                continue
            records.append({
                "caption": caption,
                "prefix": f"Figure {figure_number}.",
                "number": figure_number,
                "float": block,
            })
    return records


def _read_section_command(text, match):
    cursor = match.end()
    while cursor < len(text) and text[cursor].isspace():
        cursor += 1
    if cursor < len(text) and text[cursor] == "[":
        _, cursor = _read_tex_group(text, cursor, "[", "]")
    while cursor < len(text) and text[cursor].isspace():
        cursor += 1
    title, cursor = _read_tex_group(text, cursor, "{", "}")
    if title is None:
        return None

    labels = []
    while cursor < len(text):
        while cursor < len(text) and text[cursor].isspace():
            cursor += 1
        label = _LABEL_PATTERN.match(text, cursor)
        if not label:
            break
        labels.append(label.group(1).strip())
        cursor = label.end()
    return title, cursor, labels


def _section_label_aliases(source_text):
    text = _strip_tex_comments(source_text)
    aliases = {}
    for match in _SECTION_COMMAND_PATTERN.finditer(text):
        record = _read_section_command(text, match)
        if record is None:
            continue
        _, _, labels = record
        if labels:
            aliases.update({label: labels[0] for label in labels[1:]})
    return aliases


def _appendix_heading_records(source_text):
    """Build the lettered heading sequence after LaTeX's appendix switch."""
    text = _strip_tex_comments(source_text)
    appendix = re.search(r"\\appendix\b", text)
    if not appendix:
        return [], {}

    level_by_command = {
        "section": 1, "subsection": 2, "subsubsection": 3,
    }
    records = []
    labels = {}
    appendix_number = 0
    appendix_letter = None
    counters = {2: 0, 3: 0}
    for match in _SECTION_COMMAND_PATTERN.finditer(text, appendix.end()):
        command, starred = match.groups()
        level = level_by_command.get(command)
        if level is None:
            continue
        parsed = _read_section_command(text, match)
        if parsed is None:
            continue
        title, _, heading_labels = parsed

        prefix = None
        if not starred:
            if level == 1:
                appendix_number += 1
                number = appendix_number
                letter = ""
                while number:
                    number, remainder = divmod(number - 1, 26)
                    letter = chr(ord("A") + remainder) + letter
                appendix_letter = letter
                counters = {2: 0, 3: 0}
                prefix = f"Appendix {appendix_letter}"
            elif appendix_letter is not None:
                counters[level] += 1
                for lower_level in range(level + 1, 4):
                    counters[lower_level] = 0
                suffix = ".".join(
                    str(counters[item]) for item in range(2, level + 1))
                prefix = f"{appendix_letter}.{suffix}"
        if prefix and heading_labels:
            number_text = prefix.removeprefix("Appendix ")
            labels.update({label: number_text for label in heading_labels})
        records.append({
            "level": level,
            "prefix": prefix,
            "title": title,
        })
    return records, labels


def _table_reference_numbers(source_text):
    numbers = {}
    for record in _table_float_records(source_text):
        if record["starred"]:
            continue
        active_float = _strip_tex_comments(record["float"])
        for label in _LABEL_PATTERN.findall(active_float):
            numbers[label.strip()] = str(record["number"])
    return numbers


def _figure_reference_numbers(source_text):
    numbers = {}
    for record in _figure_caption_records(_strip_tex_comments(source_text)):
        block = record["float"]
        subfigures = _named_environment_ranges(block, {"subfigure"})
        for match in _LABEL_PATTERN.finditer(block):
            if _is_comment_offset(block, match.start()):
                continue
            subfigure = next((index for index, (start, end) in enumerate(
                subfigures, start=1) if start < match.start() < end), None)
            number = str(record["number"])
            if subfigure is not None:
                number += chr(ord("a") + subfigure - 1)
            numbers[match.group(1).strip()] = number
    return numbers


def _table_float_records(text):
    """Return numbered table floats and their caption source text."""
    records = []
    table_number = 0
    for start, end in _table_float_ranges(text):
        block = text[start:end]
        for match in _CAPTION_PATTERN.finditer(block):
            if _is_comment_offset(block, match.start()):
                continue
            cursor = match.end()
            while cursor < len(block) and block[cursor].isspace():
                cursor += 1
            if cursor < len(block) and block[cursor] == "[":
                _, cursor = _read_tex_group(block, cursor, "[", "]")
            caption, caption_end = _read_tex_group(block, cursor, "{", "}")
            if caption is None:
                continue
            starred = bool(match.group(1))
            if not starred:
                table_number += 1
            records.append({
                "number": table_number,
                "caption": caption,
                "starred": starred,
                "float": block,
            })
    return records


def _normalize_table_captions(text):
    """Keep table captions as editable paragraphs even for unsupported floats."""
    replacements = []
    table_number = 0
    for start, end in _table_float_ranges(text):
        block = text[start:end]
        captions = []
        removals = []
        for match in _CAPTION_PATTERN.finditer(block):
            if _is_comment_offset(block, match.start()):
                continue
            cursor = match.end()
            while cursor < len(block) and block[cursor].isspace():
                cursor += 1
            if cursor < len(block) and block[cursor] == "[":
                _, cursor = _read_tex_group(block, cursor, "[", "]")
            caption, caption_end = _read_tex_group(block, cursor, "{", "}")
            if caption is None:
                continue
            starred = bool(match.group(1))
            if not starred:
                table_number += 1
                prefix = f"\\textbf{{Table {table_number}.}} "
            else:
                prefix = ""
            captions.append(prefix + caption)
            removals.append((start + match.start(), start + caption_end))

        if captions:
            replacements.append((start, start, "\n\n" + "\n\n".join(captions) + "\n\n"))
            replacements.extend((begin, finish, "") for begin, finish in removals)

    for begin, finish, replacement in sorted(replacements, reverse=True):
        text = text[:begin] + replacement + text[finish:]
    return text


def _simplify_column_spec(spec):
    """Reduce tabularx column definitions to Word-friendly alignment columns."""
    columns = []
    index = 0
    alignment = None
    while index < len(spec):
        character = spec[index]
        if character.isspace() or character == "|":
            index += 1
            continue
        if spec.startswith(">{", index):
            directive, index = _read_tex_group(spec, index + 1, "{", "}")
            if directive and "\\centering" in directive:
                alignment = "c"
            elif directive and "\\raggedleft" in directive:
                alignment = "r"
            elif directive and "\\raggedright" in directive:
                alignment = "l"
            continue
        if character == "<":
            if index + 1 < len(spec) and spec[index + 1] == "{":
                _, index = _read_tex_group(spec, index + 1, "{", "}")
            else:
                index += 1
            continue
        if character in "pmb":
            _, index = _read_tex_group(spec, index + 1, "{", "}")
            columns.append(alignment or "l")
            alignment = None
            continue
        if character in "lcrXY":
            columns.append(alignment or {"X": "l", "Y": "l"}.get(character, character))
            alignment = None
            index += 1
            continue
        if character == "*":
            count, index = _read_tex_group(spec, index + 1, "{", "}")
            repeated_spec, index = _read_tex_group(spec, index, "{", "}")
            if count and repeated_spec:
                try:
                    columns.extend(_simplify_column_spec(repeated_spec) * int(count))
                except ValueError:
                    pass
            continue
        if character in "@!":
            _, index = _read_tex_group(spec, index + 1, "{", "}")
            continue
        if character == "\\":
            command = re.match(r"\\[A-Za-z]+", spec[index:])
            index += len(command.group(0)) if command else 2
            continue
        index += 1
    return "".join(columns) or "l"


def _normalize_minipage_tables(text):
    """Flatten tabularx tables wrapped with a note minipage for Pandoc."""
    for start, end in reversed(_table_float_ranges(text)):
        block = text[start:end]
        active_environments = [
            match for match in _ENVIRONMENT_PATTERN.finditer(block)
            if not _is_comment_offset(block, match.start())
        ]
        if not any(match.groups() == ("begin", "minipage")
                   for match in active_environments):
            continue
        match = next((item for item in active_environments
                      if item.groups() == ("begin", "tabularx")), None)
        if match is None:
            continue
        index = match.end()
        _, index = _read_tex_group(block, index, "{", "}")
        spec, index = _read_tex_group(block, index, "{", "}")
        end_match = next((item for item in _ENVIRONMENT_PATTERN.finditer(block, index)
                          if item.groups() == ("end", "tabularx")
                          and not _is_comment_offset(block, item.start())), None)
        if spec is None or end_match is None:
            continue
        end_start = end_match.start()
        end_finish = end_match.end()
        begin = "\\begin{tabular}{" + _simplify_column_spec(spec) + "}"
        block = (block[:match.start()] + begin + block[index:end_start]
                 + "\\end{tabular}" + block[end_finish:])
        removals = []
        for environment_match in _ENVIRONMENT_PATTERN.finditer(block):
            if (_is_comment_offset(block, environment_match.start())
                    or environment_match.group(2) != "minipage"):
                continue
            if environment_match.group(1) == "begin":
                _, finish = _read_tex_group(
                    block, environment_match.end(), "{", "}")
            else:
                finish = environment_match.end()
            removals.append((environment_match.start(), finish))
        for begin, finish in reversed(removals):
            block = block[:begin] + block[finish:]
        text = text[:start] + block + text[end:]
    return text


def _normalize_tabular_rows(text):
    """Join table cells split across source lines before Pandoc reads them."""
    for start, end in reversed(_named_environment_ranges(
            text, {"tabular", "tabularx"})):
        block = text[start:end]
        lines = block.splitlines()
        normalized = []
        for line in lines:
            if line.lstrip().startswith("&"):
                previous = next((index for index in range(len(normalized) - 1, -1, -1)
                                 if normalized[index].strip()
                                 and not normalized[index].lstrip().startswith("%")), None)
                if (previous is not None
                        and not re.search(r"\\\\(?:\[[^]]*\])?\s*$",
                                          normalized[previous])):
                    normalized[previous] = normalized[previous].rstrip() + " " + line.strip()
                    continue
            normalized.append(line)
        text = text[:start] + "\n".join(normalized) + text[end:]
    return text


def _normalize_subfigure_tabularx(text):
    """Convert subfigure tabularx blocks to Pandoc-friendly markup."""
    for start, end in reversed(_named_environment_ranges(text, {"subfigure"})):
        block = _normalize_tabular_rows(text[start:end])
        text = text[:start] + block + text[end:]
    subfigure_ranges = _named_environment_ranges(text, {"subfigure"})
    table_ranges = _named_environment_ranges(text, {"tabularx"})
    for start, end in reversed(table_ranges):
        if not any(left < start and end < right
                   for left, right in subfigure_ranges):
            continue
        block = text[start:end]
        begin = next((match for match in _ENVIRONMENT_PATTERN.finditer(block)
                      if match.groups() == ("begin", "tabularx")
                      and not _is_comment_offset(block, match.start())), None)
        finish = next((match for match in reversed(list(
            _ENVIRONMENT_PATTERN.finditer(block)))
                       if match.groups() == ("end", "tabularx")
                       and not _is_comment_offset(block, match.start())), None)
        if begin is None or finish is None:
            continue
        _, cursor = _read_tex_group(block, begin.end(), "{", "}")
        spec, cursor = _read_tex_group(block, cursor, "{", "}")
        if spec is None:
            continue
        content = block[cursor:finish.start()]
        content = re.sub(
            r"\\(?:toprule|midrule|bottomrule|addlinespace)\*?"
            r"(?:\s*\[[^]]*\])?", "", content)
        content = re.sub(
            r"\\(?:cmidrule|cline)\*?(?:\[[^]]*\])?"
            r"(?:\([^)]*\))?\s*\{[^}]*\}", "", content)
        replacement = ("\\begin{tabular}{" + _simplify_column_spec(spec) + "}"
                       + content + "\\end{tabular}")
        text = text[:start] + replacement + text[end:]
    return text


def _standalone_table_sources(record, source_text):
    """Build standalone sources for each native table in a LaTeX float."""
    block = record["float"]
    document_start = re.search(r"\\begin\s*\{document\}", source_text)
    if not document_start:
        return []
    preamble = source_text[:document_start.start()]
    caption = record["caption"]
    if not record["starred"]:
        caption = f"\\textbf{{Table {record['number']}.}} " + caption
    environments = [
        match for match in _ENVIRONMENT_PATTERN.finditer(block)
        if not _is_comment_offset(block, match.start())
    ]
    sources = []
    consumed_until = -1
    for index, match in enumerate(environments):
        environment = match.group(2)
        if (match.start() < consumed_until or match.group(1) != "begin"
                or environment not in ("tabularx", "tabular")):
            continue
        cursor = match.end()
        if environment == "tabularx":
            _, cursor = _read_tex_group(block, cursor, "{", "}")
        column_spec, cursor = _read_tex_group(block, cursor, "{", "}")
        if column_spec is None:
            continue

        depth = 1
        end_match = None
        for item in environments[index + 1:]:
            if item.start() < cursor or item.group(2) != environment:
                continue
            depth += 1 if item.group(1) == "begin" else -1
            if not depth:
                end_match = item
                break
        if end_match is None:
            continue

        table = ("\\begin{tabular}{" + _simplify_column_spec(column_spec) + "}"
                 + block[cursor:end_match.start()] + "\\end{tabular}")
        sources.append(
            preamble + "\n\\begin{document}\n" + caption + "\n\n"
            + table + "\n\\end{document}\n")
        consumed_until = end_match.end()
    return sources


def _table_column_specs(record):
    """Read width-bearing tabular column specifications from a float."""
    block = record["float"]
    environments = [match for match in _ENVIRONMENT_PATTERN.finditer(block)
                    if not _is_comment_offset(block, match.start())]
    specs = []
    for match in environments:
        if match.group(1) != "begin" or match.group(2) not in ("tabularx", "tabular"):
            continue
        cursor = match.end()
        table_width = None
        if match.group(2) == "tabularx":
            table_width, cursor = _read_tex_group(block, cursor, "{", "}")
        column_spec, _ = _read_tex_group(block, cursor, "{", "}")
        if column_spec is not None:
            specs.append((table_width, column_spec))
    return specs


def _column_width_expressions(spec):
    columns = []
    index = 0
    while index < len(spec):
        character = spec[index]
        if character.isspace() or character == "|":
            index += 1
            continue
        if spec.startswith(">{", index):
            _, index = _read_tex_group(spec, index + 1, "{", "}")
            continue
        if character == "<":
            if index + 1 < len(spec) and spec[index + 1] == "{":
                _, index = _read_tex_group(spec, index + 1, "{", "}")
            else:
                index += 1
            continue
        if character in "pmb":
            width, index = _read_tex_group(spec, index + 1, "{", "}")
            columns.append(width)
            continue
        if character in "lcrXY":
            columns.append(None)
            index += 1
            continue
        if character == "*":
            count, index = _read_tex_group(spec, index + 1, "{", "}")
            repeated, index = _read_tex_group(spec, index, "{", "}")
            if count and repeated:
                try:
                    columns.extend(_column_width_expressions(repeated) * int(count))
                except ValueError:
                    pass
            continue
        if character in "@!":
            _, index = _read_tex_group(spec, index + 1, "{", "}")
            continue
        if character == "\\":
            command = re.match(r"\\[A-Za-z]+", spec[index:])
            index += len(command.group(0)) if command else 2
            continue
        index += 1
    return columns


def _tex_width_inches(expression, reference_width):
    if expression is None:
        return None
    expression = expression.strip()
    relative = re.fullmatch(
        r"(-?(?:\d+(?:\.\d*)?|\.\d+))\s*\\(?:linewidth|textwidth|columnwidth)",
        expression,
    )
    if relative:
        return float(relative.group(1)) * reference_width
    if expression in (r"\linewidth", r"\textwidth", r"\columnwidth"):
        return reference_width
    return _length_in_inches(expression)


def _column_widths_inches(table_width_expression, column_spec, text_width):
    table_width = _tex_width_inches(table_width_expression, text_width)
    expressions = _column_width_expressions(column_spec)
    widths = [_tex_width_inches(value, text_width) for value in expressions]
    if not widths or not any(width is not None for width in widths):
        return None
    if table_width is None:
        return None

    fixed_total = sum(width for width in widths if width is not None)
    flexible_count = sum(width is None for width in widths)
    if fixed_total > table_width and fixed_total:
        scale = table_width / fixed_total
        widths = [width * scale if width is not None else None for width in widths]
        fixed_total = table_width
    if flexible_count:
        flexible_width = max(0.0, table_width - fixed_total) / flexible_count
        widths = [flexible_width if width is None else width for width in widths]
    elif fixed_total < table_width:
        scale = table_width / fixed_total
        widths = [width * scale for width in widths]
    return widths


def _set_table_column_widths(table_element, widths):
    grid = table_element.find(qn("w:tblGrid"))
    if grid is None:
        return False
    columns = list(grid.findall(qn("w:gridCol")))
    if len(columns) != len(widths):
        return False
    twips = [max(1, round(width * 1440)) for width in widths]
    table_properties = table_element.find(qn("w:tblPr"))
    if table_properties is None:
        table_properties = OxmlElement("w:tblPr")
        table_element.insert(0, table_properties)
    table_width = table_properties.find(qn("w:tblW"))
    if table_width is None:
        table_width = OxmlElement("w:tblW")
        table_properties.insert(0, table_width)
    table_width.set(qn("w:w"), str(sum(twips)))
    table_width.set(qn("w:type"), "dxa")
    layout = table_properties.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        table_properties.append(layout)
    layout.set(qn("w:type"), "fixed")

    for column, width in zip(columns, twips):
        column.set(qn("w:w"), str(width))
    for row in table_element.findall(qn("w:tr")):
        cell_widths = []
        column_index = 0
        for cell in row.findall(qn("w:tc")):
            cell_properties = cell.find(qn("w:tcPr"))
            span = (cell_properties.find(qn("w:gridSpan"))
                    if cell_properties is not None else None)
            span = int(span.get(qn("w:val"))) if span is not None else 1
            if column_index + span > len(twips):
                break
            cell_widths.append((cell, sum(twips[column_index:column_index + span])))
            column_index += span
        if column_index != len(twips):
            continue
        for cell, width in cell_widths:
            cell_properties = cell.find(qn("w:tcPr"))
            if cell_properties is None:
                cell_properties = OxmlElement("w:tcPr")
                cell.insert(0, cell_properties)
            cell_width = cell_properties.find(qn("w:tcW"))
            if cell_width is None:
                cell_width = OxmlElement("w:tcW")
                cell_properties.insert(0, cell_width)
            cell_width.set(qn("w:w"), str(width))
            cell_width.set(qn("w:type"), "dxa")
    return True


def _preserve_latex_table_widths(source_text, output_path):
    """Apply explicit LaTeX table widths to the corresponding editable Word tables."""
    document = Document(output_path)
    section = document.sections[0]
    text_width = (section.page_width - section.left_margin - section.right_margin) / 914400
    body = document.element.body
    children = list(body)
    changed = False

    for record in _table_float_records(source_text):
        if record["starred"]:
            continue
        source_widths = [
            widths for table_width, column_spec in _table_column_specs(record)
            if (widths := _column_widths_inches(
                table_width, column_spec, text_width)) is not None
        ]
        if not source_widths:
            continue
        marker = f"Table {record['number']}."
        caption_index = next((index for index, element in enumerate(children)
                              if element.tag == qn("w:p")
                              and Paragraph(element, document).text.strip().startswith(marker)),
                             None)
        if caption_index is None:
            continue

        tables = []
        for element in children[caption_index + 1:]:
            if element.tag == qn("w:p") and re.match(
                    r"^Table\s+\d+\.", Paragraph(element, document).text.strip()):
                break
            if (element.tag == qn("w:tbl")
                    and not element.findall(".//" + qn("w:drawing"))):
                tables.append(element)
        for table_element, widths in zip(tables, source_widths):
            changed = _set_table_column_widths(table_element, widths) or changed

    if changed:
        document.save(output_path)


def _subfigure_blocks(figure_record):
    blocks = []
    figure = figure_record["float"]
    for start, end in _named_environment_ranges(figure, {"subfigure"}):
        block = figure[start:end]
        begin = next((match for match in _ENVIRONMENT_PATTERN.finditer(block)
                      if match.groups() == ("begin", "subfigure")
                      and not _is_comment_offset(block, match.start())), None)
        if begin is None:
            continue
        cursor = begin.end()
        if cursor < len(block) and block[cursor] == "[":
            _, cursor = _read_tex_group(block, cursor, "[", "]")
        width, _ = _read_tex_group(block, cursor, "{", "}")
        match = re.fullmatch(
            r"\s*(?:(\d+(?:\.\d+)?)\s*)?\\(?:textwidth|linewidth|columnwidth)\s*",
            width or "",
        )
        blocks.append({
            "block": block,
            "ratio": float(match.group(1) or 1) if match else None,
        })
    return blocks


def _preserve_wide_subfigure_tables(source_text, output_path):
    """Keep wide LaTeX subfigures stacked and their nested tables readable."""
    document = Document(output_path)
    figure_tables = [table._tbl for table in document.tables
                     if (style := table._tbl.find(
                         qn("w:tblPr")).find(qn("w:tblStyle"))) is not None
                     and style.get(qn("w:val")) == "FigureTable"]
    source_figures = [record for record in _figure_caption_records(source_text)
                      if _subfigure_blocks(record)]
    if not figure_tables or not source_figures:
        return

    section = document.sections[0]
    text_width = (section.page_width - section.left_margin - section.right_margin) / 914400
    changed = False
    for record, table_element in zip(source_figures, figure_tables):
        subfigures = _subfigure_blocks(record)
        rows = table_element.findall(qn("w:tr"))
        if (len(subfigures) < 2
                or any(item["ratio"] is None or item["ratio"] < 0.7
                       for item in subfigures)
                or len(rows) != 1):
            continue
        cells = rows[0].findall(qn("w:tc"))
        if len(cells) != len(subfigures):
            continue

        grid = table_element.find(qn("w:tblGrid"))
        if grid is None:
            continue
        for column in list(grid):
            grid.remove(column)
        column = OxmlElement("w:gridCol")
        column.set(qn("w:w"), str(round(text_width * 1440)))
        grid.append(column)
        table_element.remove(rows[0])
        for cell in cells:
            row = OxmlElement("w:tr")
            row.append(deepcopy(cell))
            table_element.append(row)
        _set_table_column_widths(table_element, [text_width])
        changed = True

        for subfigure, row in zip(subfigures, table_element.findall(qn("w:tr"))):
            cell = row.find(qn("w:tc"))
            nested_tables = list(cell.iter(qn("w:tbl")))
            used = set()
            local_width = text_width * subfigure["ratio"]
            for table_width, column_spec in _table_column_specs(
                    {"float": subfigure["block"]}):
                if table_width is None:
                    continue
                widths = _column_widths_inches(
                    table_width, column_spec, local_width)
                if not widths:
                    continue
                target = next((item for item in nested_tables
                               if id(item) not in used
                               and len(item.findall(
                                   f"{qn('w:tblGrid')}/{qn('w:gridCol')}")) == len(widths)),
                              None)
                if target is None:
                    continue
                used.add(id(target))
                changed = _set_table_column_widths(target, widths) or changed

    if changed:
        document.save(output_path)


def _plain_caption_text(caption):
    try:
        text = pypandoc.convert_text(
            caption, to="plain", format="latex", extra_args=["--wrap=none"])
    except RuntimeError:
        text = re.sub(r"\\[A-Za-z]+\*?", "", caption)
    return " ".join(text.split())


def _normalized_text(text):
    return " ".join(re.findall(r"[a-z0-9]+", text.lower()))


def _prepend_caption_label(paragraph, label):
    """Insert a bold figure label without flattening existing caption runs."""
    if re.match(r"^\s*Figure\s+\d+\.", paragraph.text, re.I):
        return False

    for run in paragraph.runs:
        if not run.text:
            continue
        trimmed = run.text.lstrip()
        if trimmed:
            if trimmed != run.text:
                run.text = trimmed
            break
        run._element.getparent().remove(run._element)

    runs = []
    label_run = OxmlElement("w:r")
    properties = OxmlElement("w:rPr")
    properties.append(OxmlElement("w:b"))
    label_run.append(properties)
    label_text = OxmlElement("w:t")
    label_text.text = label
    label_run.append(label_text)
    runs.append(label_run)

    space_run = OxmlElement("w:r")
    space_text = OxmlElement("w:t")
    space_text.set(qn("xml:space"), "preserve")
    space_text.text = " "
    space_run.append(space_text)
    runs.append(space_run)

    first_content_index = 1 if paragraph._p.pPr is not None else 0
    for offset, run in enumerate(runs):
        paragraph._p.insert(first_content_index + offset, run)
    return True


def _restore_figure_caption_labels(source_text, output_path):
    """Restore numbered figure labels while retaining Pandoc's caption text."""
    records = _figure_caption_records(source_text)
    if not records:
        return

    document = Document(output_path)
    paragraphs = [paragraph for paragraph in document.paragraphs
                  if paragraph.style.name == "Image Caption"]
    if not paragraphs:
        return

    if len(records) == len(paragraphs):
        matches = list(zip(records, paragraphs))
    else:
        matches = []
        next_paragraph = 0
        for record in records:
            expected = _normalized_text(_plain_caption_text(record["caption"]))
            if not expected:
                continue
            candidates = [
                (SequenceMatcher(
                    None, expected, _normalized_text(paragraph.text)).ratio(),
                 index, paragraph)
                for index, paragraph in enumerate(
                    paragraphs[next_paragraph:], next_paragraph)
            ]
            if not candidates:
                break
            similarity, index, paragraph = max(candidates, key=lambda item: item[0])
            if similarity < 0.45:
                continue
            matches.append((record, paragraph))
            next_paragraph = index + 1

    changed = False
    for record, paragraph in matches:
        changed = _prepend_caption_label(paragraph, record["prefix"]) or changed
    if changed:
        document.save(output_path)


def _restore_appendix_headings(document, source_text):
    records, _ = _appendix_heading_records(source_text)
    if not records:
        return False

    headings = [paragraph for paragraph in document.paragraphs
                if paragraph.style.name in ("Heading 1", "Heading 2", "Heading 3")]
    next_heading = 0
    changed = False
    for record in records:
        expected = _normalized_text(_plain_caption_text(record["title"]))
        if not expected:
            continue
        style_name = f"Heading {record['level']}"
        candidates = [
            (SequenceMatcher(
                None, expected, _normalized_text(paragraph.text)).ratio(),
             index, paragraph)
            for index, paragraph in enumerate(headings[next_heading:], next_heading)
            if paragraph.style.name == style_name
        ]
        if not candidates:
            continue
        similarity, index, paragraph = max(candidates, key=lambda item: item[0])
        if similarity < 0.45:
            continue
        next_heading = index + 1
        prefix = record["prefix"]
        if not prefix:
            continue
        for run_index, run in enumerate(paragraph.runs):
            if not re.fullmatch(r"\d+(?:\.\d+)*", run.text.strip()):
                continue
            if run.text != prefix:
                run.text = prefix
                changed = True
            for following in paragraph.runs[run_index + 1:]:
                if following.text == "\t":
                    following.text = " "
                    changed = True
                    break
                if following.text.strip():
                    break
            break
    return changed


def _hyperlink_text(hyperlink):
    return "".join(node.text or "" for node in hyperlink.iter(qn("w:t")))


def _set_hyperlink_text(hyperlink, text):
    nodes = list(hyperlink.iter(qn("w:t")))
    if not nodes:
        return False
    if _hyperlink_text(hyperlink) == text:
        return False
    nodes[0].text = text
    for node in nodes[1:]:
        node.text = ""
    return True


def _restore_cross_reference_numbers(source_text, output_path):
    """Correct Pandoc reference numbers for floats and LaTeX appendix labels."""
    document = Document(output_path)
    hyperlinks = list(document.element.iter(qn("w:hyperlink")))
    if not hyperlinks:
        return

    resolved = {}
    for hyperlink in hyperlinks:
        anchor = hyperlink.get(qn("w:anchor"))
        value = _hyperlink_text(hyperlink)
        if anchor and value and not re.fullmatch(r"\[[^\]]+\]", value):
            resolved.setdefault(anchor, value)

    _, appendix_numbers = _appendix_heading_records(source_text)
    label_numbers = dict(appendix_numbers)
    label_numbers.update(_figure_reference_numbers(source_text))
    label_numbers.update(_table_reference_numbers(source_text))
    for alias, target in _section_label_aliases(source_text).items():
        if alias in label_numbers:
            continue
        value = label_numbers.get(target, resolved.get(target))
        if value and not re.fullmatch(r"\[[^\]]+\]", value):
            label_numbers[alias] = value

    changed = False
    for hyperlink in hyperlinks:
        anchor = hyperlink.get(qn("w:anchor"))
        value = label_numbers.get(anchor)
        if value is not None:
            changed = _set_hyperlink_text(hyperlink, value) or changed
    changed = _restore_appendix_headings(document, source_text) or changed
    if changed:
        document.save(output_path)


def _table_cells(table):
    return {
        _normalized_text(cell.text)
        for row in table.rows for cell in row.cells
        if len(_normalized_text(cell.text)) >= 4
    }


def _table_matches(candidate, existing):
    if (len(candidate.rows), len(candidate.columns)) != (
            len(existing.rows), len(existing.columns)):
        return False
    if existing._tbl.findall(".//" + qn("w:drawing")):
        return False
    candidate_cells = _table_cells(candidate)
    existing_cells = _table_cells(existing)
    if not candidate_cells or not existing_cells:
        return False
    shared = len(candidate_cells & existing_cells)
    # ponytail: overlap tolerates merged-cell differences; source-position mapping
    # would be needed if Pandoc starts changing table dimensions between passes.
    return (shared / len(candidate_cells) >= 0.35
            and shared / len(existing_cells) >= 0.55)


def _insert_caption_and_table(document, number, caption_text, table_element=None):
    body = document.element.body
    next_caption = None
    pattern = re.compile(r"^Table\s+(\d+)\.")
    for paragraph in document.paragraphs:
        match = pattern.match(paragraph.text.strip())
        if match and int(match.group(1)) > number:
            next_caption = paragraph._p
            break

    if next_caption is None:
        section_properties = body.sectPr
        index = (list(body).index(section_properties)
                 if section_properties is not None else len(body))
    else:
        index = list(body).index(next_caption)

    paragraph_element = OxmlElement("w:p")
    paragraph = Paragraph(paragraph_element, document._body)
    paragraph.style = _ensure_table_caption_style(document)
    label = paragraph.add_run(f"Table {number}.")
    label.bold = True
    if caption_text:
        paragraph.add_run(" " + caption_text)
    body.insert(index, paragraph_element)
    inserted_table = None
    if table_element is not None:
        inserted_table = deepcopy(table_element)
        body.insert(index + 1, inserted_table)
    return paragraph_element, inserted_table


def _insert_table_after_element(document, anchor, table_element):
    body = document.element.body
    index = list(body).index(anchor) + 1
    inserted = deepcopy(table_element)
    body.insert(index, inserted)
    return inserted


def _restore_table_captions_and_missing_tables(
        source_text, records, output_path, source_dir,
        reference_docx_path, temp_dir):
    """Restore captions and recover native tables Pandoc silently skipped."""
    document = Document(output_path)
    changed = False
    table_caption_style = _ensure_table_caption_style(document)
    for record in records:
        if record["starred"]:
            continue
        number = record["number"]
        marker = f"Table {number}."
        matching = next((paragraph for paragraph in document.paragraphs
                         if paragraph.text.strip().startswith(marker)), None)
        if matching is not None:
            if matching.style.name != "Table Caption":
                changed = True
            matching.style = table_caption_style
            if not matching.text[len(marker):].strip():
                caption_text = _plain_caption_text(record["caption"])
                for run in list(matching.runs):
                    matching._p.remove(run._r)
                label = matching.add_run(marker)
                label.bold = True
                if caption_text:
                    matching.add_run(" " + caption_text)
                changed = True

        missing_tables = []
        for table_index, table_source in enumerate(
                _standalone_table_sources(record, source_text), start=1):
            candidate_source = os.path.join(
                temp_dir, f"fallback-table-{number}-{table_index}.tex")
            candidate_docx = os.path.join(
                temp_dir, f"fallback-table-{number}-{table_index}.docx")
            with open(candidate_source, "w", encoding="utf-8", newline="") as source_file:
                source_file.write(table_source)
            pypandoc.convert_file(
                candidate_source,
                to="docx",
                format="latex",
                outputfile=candidate_docx,
                extra_args=[
                    f"--resource-path={source_dir}",
                    f"--reference-doc={reference_docx_path}",
                    "--number-sections",
                ],
                cworkdir=source_dir,
            )
            candidate = Document(candidate_docx)
            if candidate.tables:
                table = max(
                    candidate.tables,
                    key=lambda item: len(item.rows) * len(item.columns))
                if not any(_table_matches(table, existing)
                           for existing in document.tables):
                    missing_tables.append(table._tbl)

        if matching is None:
            caption_text = _plain_caption_text(record["caption"])
            paragraph_element, anchor = _insert_caption_and_table(
                document, number, caption_text,
                missing_tables.pop(0) if missing_tables else None)
            if anchor is None:
                anchor = paragraph_element
            for table_element in missing_tables:
                anchor = _insert_table_after_element(
                    document, anchor, table_element)
            changed = True
        elif missing_tables:
            anchor = matching._p
            for table_element in missing_tables:
                anchor = _insert_table_after_element(
                    document, anchor, table_element)
            changed = True

    if changed:
        document.save(output_path)


def _preserve_wide_subfigure_image_widths(source_text, output_path):
    image_count, width_overrides = _wide_subfigure_linewidth_images(source_text)
    if not width_overrides:
        return

    document = Document(output_path)
    if len(document.inline_shapes) != image_count:
        return
    section = document.sections[0]
    text_width = section.page_width - section.left_margin - section.right_margin
    for image_index, ratio in width_overrides:
        image = document.inline_shapes[image_index]
        old_width, old_height = int(image.width), int(image.height)
        new_width = Inches(text_width / 914400 * ratio)
        image.width = new_width
        image.height = round(old_height * int(new_width) / old_width)
    document.save(output_path)


def _bibliography_files(source_path):
    """Return bibliography files declared by the source that exist beside it."""
    source_dir = os.path.dirname(os.path.abspath(source_path))
    with open(source_path, encoding="utf-8", errors="replace") as source_file:
        text = _strip_tex_comments(source_file.read())

    files = []
    for value in _BIBLIOGRAPHY_PATTERN.findall(text):
        for entry in value.split(","):
            entry = entry.strip()
            if not entry:
                continue
            if not os.path.splitext(entry)[1]:
                entry += ".bib"
            path = os.path.abspath(os.path.join(source_dir, entry))
            if os.path.isfile(path) and path not in files:
                files.append(path)
    return files


def convert_latex_to_docx(source_path, destination_path, on_status=None):
    """Convert one .tex source and its local project resources to .docx."""
    source_path = os.path.abspath(source_path)
    destination_path = os.path.abspath(destination_path)
    if not os.path.isfile(source_path):
        raise FileNotFoundError(f"LaTeX source file not found: {source_path}")
    if not source_path.lower().endswith(".tex"):
        raise ValueError("Choose a LaTeX source file with the .tex extension.")

    if on_status:
        on_status("Converting LaTeX to Word…")

    source_dir = os.path.dirname(source_path)
    extra_args = [f"--resource-path={source_dir}"]
    bibliography_files = _bibliography_files(source_path)
    for bibliography_path in bibliography_files:
        extra_args.append(f"--bibliography={bibliography_path}")
    if bibliography_files:
        extra_args.append("--citeproc")

    with open(source_path, encoding="utf-8", errors="replace") as source_file:
        source_text = source_file.read()
        table_records = _table_float_records(source_text)
        cleaned_source = _normalize_table_captions(source_text)
        cleaned_source = _normalize_minipage_tables(cleaned_source)
        cleaned_source = _normalize_subfigure_tabularx(cleaned_source)
        cleaned_source = _normalize_includegraphics_options(cleaned_source)
    with tempfile.TemporaryDirectory(prefix="pdf2docx-latex-") as temp_dir:
        temp_source_path = os.path.join(temp_dir, os.path.basename(source_path))
        reference_docx_path = os.path.join(temp_dir, "reference.docx")
        with open(temp_source_path, "w", encoding="utf-8", newline="") as temp_source:
            temp_source.write(cleaned_source)
        _create_reference_docx(source_text, reference_docx_path)
        try:
            pypandoc.convert_file(
                temp_source_path,
                to="docx",
                format="latex",
                outputfile=destination_path,
                extra_args=extra_args + [
                    f"--reference-doc={reference_docx_path}",
                    "--number-sections",
                ],
                cworkdir=source_dir,
            )
        except RuntimeError as exc:
            raise RuntimeError(str(exc).replace(temp_source_path, source_path)) from exc
        _restore_table_captions_and_missing_tables(
            source_text, table_records, destination_path, source_dir,
            reference_docx_path, temp_dir)
        _restore_figure_caption_labels(source_text, destination_path)
        _restore_cross_reference_numbers(source_text, destination_path)
        _preserve_latex_table_widths(source_text, destination_path)
        _preserve_wide_subfigure_tables(source_text, destination_path)
        _preserve_wide_subfigure_image_widths(source_text, destination_path)
    return destination_path
