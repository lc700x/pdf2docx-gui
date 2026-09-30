"""Repair what a page-faithful conversion leaves behind.

pdf2docx reproduces a PDF page by page, which is the right instinct for layout
and the wrong one for a document someone will edit. Three things follow from
it, and this module undoes all three:

* **Fonts.** A PDF embeds its own; the .docx only names them. A TeX document
  names fonts almost nobody has installed, so Word substitutes something wider
  and the text looks a size too big.
* **Pages.** Every PDF page becomes its own Word section, ending with the page
  number as body text. Reflowed text is never exactly as tall as the page it
  came from, so a section that runs two lines long spills them onto a page of
  their own and the break leaves the rest of it blank.
* **Spaces.** Justified lines are set with squeezed inter-word space. Below a
  threshold the extractor reads the gap as no gap: each word still starts its
  own run, but with no space before it, so the line renders as one long word.

None of this changes a single character of the text.
"""

import re
import zipfile
from collections import Counter

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

CAPTION = re.compile(
    r'^\s*('
    r'(Table|Figure|Fig\.?)\s*\d+\s*[:.]'
    r'|\([a-z]\)\s'
    r')', re.I)

SUBCAPTION = re.compile(r'^\s*\([a-z]\)\s')

# The first line of a page that was started for a section: "1 Introduction",
# "A Dataset definitions", "3.2 Overall accuracy".
HEADING = re.compile(r'^(\d+(\.\d+)*|[A-H])\s+[A-Z]')

W_P = qn('w:p')
W_TBL = qn('w:tbl')

# --------------------------------------------------------------------------
# fonts
# --------------------------------------------------------------------------

# TeX names its fonts after the packages that ship them. None of these are
# installed on a typical machine, and all of them are Times or Courier clones.
SERIF = re.compile(r'(termes|newtx|nimbus|times|cm(r|ss|mi|sy|ex)|'
                   r'lmroman|t1x(?!.*tt)|txsy|txmi|txex|txsym|stix)', re.I)
MONO = re.compile(r'(t1x[a-z]*tt|cmtt|lmmono|courier|mono)', re.I)

FONT_ATTRS = ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia')


def _substitute(name):
    if MONO.search(name):
        return 'Courier New'
    if SERIF.search(name):
        return 'Times New Roman'
    return None


def substitute_fonts(path):
    """Point every unavailable TeX font at one the reader will actually have."""
    seen = Counter()
    temp = path + '.fonts'
    with zipfile.ZipFile(path) as zin, \
            zipfile.ZipFile(temp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith('word/') and item.filename.endswith('.xml'):
                text = data.decode('utf8')

                def swap(match):
                    better = _substitute(match.group(2))
                    if better:
                        seen[match.group(2)] += 1
                    return match.group(1) + '"' + (better or match.group(2)) + '"'

                text = re.sub(r'(w:(?:ascii|hAnsi|cs|eastAsia)=)"([^"]+)"',
                              swap, text)
                data = text.encode('utf8')
            zout.writestr(item, data)
    import os
    os.replace(temp, path)
    return sum(seen.values())


# --------------------------------------------------------------------------
# pages
# --------------------------------------------------------------------------

def _text_of(element):
    return ''.join(node.text or '' for node in element.iter(qn('w:t'))).strip()


def _sect_of(paragraph):
    ppr = paragraph.find(qn('w:pPr'))
    return None if ppr is None else ppr.find(qn('w:sectPr'))


def _add_page_field(paragraph):
    for kind, value in (('begin', None), ('instr', ' PAGE '),
                        ('separate', None), ('text', '1'), ('end', None)):
        run = paragraph.add_run()
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        if kind == 'instr':
            node = OxmlElement('w:instrText')
            node.set(qn('xml:space'), 'preserve')
            node.text = value
        elif kind == 'text':
            node = OxmlElement('w:t')
            node.text = value
        else:
            node = OxmlElement('w:fldChar')
            node.set(qn('w:fldCharType'), kind)
        run._r.append(node)


def deliberate_breaks(pdf_path, slack=90):
    """Pages the PDF ended on purpose, rather than because they filled up.

    Every page break has to go, or content stays stranded on half-empty pages;
    but not every page break was forced by the text running out of room. A
    title page, and a page before an appendix, stop well short of the bottom,
    and that space is the author's intent. Those breaks are kept.

    `slack` is how far above the lowest anything reaches on a full page a page
    must stop before it counts as deliberate, in points.
    """
    import fitz

    bottoms = []
    with fitz.open(pdf_path) as pdf:
        for page in pdf:
            folio = str(page.number + 1)
            boxes = [b[:4] for b in page.get_text('blocks')
                     if b[4].strip() and b[4].strip() != folio]
            boxes += [tuple(page.get_image_bbox(img))
                      for img in page.get_images(full=True)
                      if page.get_image_bbox(img)]
            bottoms.append(max((b[3] for b in boxes), default=0.0))

    if not bottoms:
        return set()
    full = max(bottoms)
    keep = {i for i, bottom in enumerate(bottoms) if bottom < full - slack}

    # A page that opens a numbered section or an appendix was started for that
    # section, whether or not the page before it happened to fill up. The title
    # page is the clearest case: nothing marks it in the source, but the paper
    # begins on the page after it.
    with fitz.open(pdf_path) as pdf:
        for page in pdf:
            if page.number == 0:
                continue
            lines = [l.strip() for l in page.get_text().splitlines() if l.strip()]
            # A heading sets its number apart from its title, so the extractor
            # reads "1" and "Introduction" as two lines.
            opening = ' '.join(lines[:2])
            if HEADING.match(opening):
                keep.add(page.number - 1)
    return keep


def _page_break_paragraph(template):
    paragraph = template.makeelement(W_P, {})
    run = OxmlElement('w:r')
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    run.append(br)
    paragraph.append(run)
    return paragraph


def flow_pages(document, keep_breaks=(), bottom_margin=Pt(48),
               footer_distance=Pt(26)):
    """Let the text run on, with the page number in a real footer.

    Returns (folios moved, breaks dropped, blank paragraphs removed).
    """
    body = document.element.body
    first = document.sections[0]
    setup = dict(page_width=first.page_width, page_height=first.page_height,
                 left_margin=first.left_margin, right_margin=first.right_margin,
                 top_margin=first.top_margin)

    # 1. the page number is body text at the foot of each section; the footer
    #    is where it belongs, and where it costs the page no height
    folios = 0
    for paragraph in list(body.iterchildren(W_P)):
        if not re.fullmatch(r'\d{1,4}', _text_of(paragraph)):
            continue
        following = paragraph.getnext()
        while (following is not None and following.tag == W_P
               and not _text_of(following) and _sect_of(following) is None):
            following = following.getnext()
        ends_section = (_sect_of(paragraph) is not None
                        or following is None
                        or following.tag == qn('w:sectPr')
                        or (following.tag == W_P and not _text_of(following)
                            and _sect_of(following) is not None))
        if not ends_section:
            continue
        section = _sect_of(paragraph)
        if section is not None:          # keep the break, drop only the number
            keeper = paragraph.makeelement(W_P, {})
            ppr = OxmlElement('w:pPr')
            ppr.append(section)
            keeper.append(ppr)
            paragraph.addnext(keeper)
        body.remove(paragraph)
        folios += 1

    # 2. one section per PDF page is what strands content on half-empty pages,
    #    so the breaks go — except where the PDF page stopped early because the
    #    author wanted a fresh page there
    breaks = kept = 0
    carriers = [p for p in list(body.iterchildren(W_P)) if _sect_of(p) is not None]
    for page, paragraph in enumerate(carriers):
        paragraph.find(qn('w:pPr')).remove(_sect_of(paragraph))
        if page in keep_breaks:
            paragraph.addprevious(_page_break_paragraph(paragraph))
            kept += 1
        breaks += 1

    # 3. an empty paragraph that only ever carried a break, or padded the foot
    #    of a page, now shows as a gap in the middle of a sentence
    blanks = 0
    for paragraph in list(body.iterchildren(W_P)):
        if _text_of(paragraph) or paragraph.findall('.//' + qn('w:drawing')):
            continue
        if any(br.get(qn('w:type')) == 'page'          # a page break is empty
               for br in paragraph.findall('.//' + qn('w:br'))):
            continue
        previous, following = paragraph.getprevious(), paragraph.getnext()
        if previous is None or following is None:
            continue
        if previous.tag not in (W_P, W_TBL) or following.tag not in (W_P, W_TBL):
            continue
        body.remove(paragraph)
        blanks += 1

    # 4. spacing that was holding a page number clear of the text
    for paragraph in body.iterchildren(W_P):
        ppr = paragraph.find(qn('w:pPr'))
        spacing = None if ppr is None else ppr.find(qn('w:spacing'))
        if spacing is None:
            continue
        for edge in ('w:before', 'w:after'):
            if int(spacing.get(qn(edge)) or 0) > 400:
                spacing.set(qn(edge), '240')

    # 5. a figure and the caption under it belong on the same page, and the
    #    panels of one figure belong together: without this the text simply
    #    flows and a break can fall between them
    held = 0
    for paragraph in body.iterchildren(W_P):
        text = _text_of(paragraph)
        has_picture = bool(paragraph.findall('.//' + qn('w:drawing')))
        # A panel's own caption sits between two panels, so it has to hold on
        # to what follows as well; only the figure's caption ends the group.
        is_panel = bool(SUBCAPTION.match(text))
        is_caption = bool(CAPTION.match(text))
        if not (has_picture or is_panel or is_caption):
            continue
        ppr = paragraph.find(qn('w:pPr'))
        if ppr is None:
            ppr = OxmlElement('w:pPr')
            paragraph.insert(0, ppr)
        tags = ['w:keepLines']
        if has_picture or is_panel:
            tags.append('w:keepNext')
        for tag in tags:
            if ppr.find(qn(tag)) is None:
                ppr.insert(0, OxmlElement(tag))
        held += 1

    for section in document.sections:
        for name, value in setup.items():
            setattr(section, name, value)
        section.bottom_margin = bottom_margin
        section.footer_distance = footer_distance

    footer = document.sections[0].footer
    footer.is_linked_to_previous = False
    folio = footer.paragraphs[0]
    folio.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_page_field(folio)
    for section in document.sections[1:]:
        section.footer.is_linked_to_previous = True

    return folios, breaks, blanks, kept, held


# --------------------------------------------------------------------------
# spaces
# --------------------------------------------------------------------------

LETTER = r'[^\W\d_]'                     # any letter, including math italics
TAIL = re.compile(LETTER + r"(?:" + LETTER + r"|['-])*$", re.UNICODE)
HEAD = re.compile(r"^" + LETTER + r"(?:" + LETTER + r"|['-])*", re.UNICODE)
WORD = re.compile(LETTER + r"(?:" + LETTER + r"|['-])+", re.UNICODE)

LABELS = {'figure', 'table', 'section', 'appendix', 'equation', 'fig', 'eq',
          'chapter', 'january', 'february', 'march', 'april', 'may', 'june',
          'july', 'august', 'september', 'october', 'november', 'december'}

COMMON = {'et', 'al', 'the', 'and', 'for', 'was', 'are', 'per', 'via', 'not',
          'but', 'its', 'our', 'their', 'with', 'from', 'that', 'this'}


def _read_vocab(path):
    try:
        text = open(path, encoding='utf8', errors='replace').read()
    except OSError:
        return set()
    text = re.sub(r'(?<!\\)%.*', ' ', text)      # TeX comments
    text = re.sub(r'\\[a-zA-Z@]+', ' ', text)    # TeX commands
    return {w.lower() for w in WORD.findall(text)}


def build_vocabulary(document, extra_paths=()):
    """Words this document is known to use, to judge where a space belongs.

    Most of the text survives conversion intact, so the document is its own
    dictionary: a word seen more than once, with space around it, is a word.
    A source file alongside the PDF, such as the LaTeX it was built from, adds
    the names and terms that appear only once.
    """
    counts = Counter()
    for paragraph in document.paragraphs:
        counts.update(w.lower() for w in WORD.findall(paragraph.text))
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                counts.update(w.lower() for w in WORD.findall(cell.text))

    vocab = {word for word, n in counts.items() if n >= 2}
    vocab |= COMMON | LABELS
    for path in extra_paths:
        vocab |= _read_vocab(path)
    return vocab


def _is_shifted(run):
    rpr = run._r.find(qn('w:rPr'))
    if rpr is None:
        return False
    shift = rpr.find(qn('w:vertAlign'))
    return shift is not None and shift.get(qn('w:val')) in ('superscript',
                                                            'subscript')


def _needs_space(before, after, run_after, vocab):
    if not before or not after:
        return False
    if before[-1].isspace() or after[0].isspace():
        return False
    if before[-1] == '-' or after[0] == '-':          # hyphenated or compound
        return False
    if _is_shifted(run_after):                        # a marker, not a word
        return False
    if after[0] == '(':                               # "ratio(SNR)"
        return before[-1].isalnum() or before[-1] in '.,;:)'
    if after[0].isdigit():
        if before[-1] in ',;:':                       # "al.,2009"
            return not (len(before) > 1 and before[-2].isdigit())
        tail = TAIL.search(before)                    # "Figure4"
        return bool(tail) and tail.group(0).lower() in LABELS
    if not after[0].isalpha():
        return False
    if before[-1] in '.,;:)':                         # end of sentence or clause
        return after[0].isupper() or before[-1] != '.'
    if not before[-1].isalpha():
        return False
    tail, head = TAIL.search(before), HEAD.match(after)
    if not tail:
        return False
    if not head:                                      # a symbol, such as math
        return tail.group(0).lower() in vocab
    joined = (tail.group(0) + head.group(0)).lower()
    if joined in vocab:                               # one word, split for style
        return False
    return tail.group(0).lower() in vocab or head.group(0).lower() in vocab


def restore_spaces(document, vocab):
    """Put back the space between runs that the extractor read as no space."""
    restored = 0

    def walk(paragraphs):
        nonlocal restored
        for paragraph in paragraphs:
            runs = paragraph.runs
            for i in range(len(runs) - 1):
                before, after = runs[i].text, runs[i + 1].text
                if _needs_space(before, after, runs[i + 1], vocab):
                    runs[i].text = before + ' '
                    restored += 1

    walk(document.paragraphs)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                walk(cell.paragraphs)
    return restored


# --------------------------------------------------------------------------
# alignment
# --------------------------------------------------------------------------

JUSTIFY_FROM = 90       # characters: shorter than this and it cannot wrap


def justify_body(document):
    """Set the body justified, as the PDF had it, without stretching a heading.

    pdf2docx reads alignment off where the ink falls, so body text arrives
    left-aligned and a line the PDF had stretched arrives centred. Length is
    the reliable signal: only a paragraph long enough to wrap is body text.
    """
    justified = relaxed = 0
    for paragraph in document.paragraphs:
        ppr = paragraph._p.find(qn('w:pPr'))
        if ppr is None:
            continue
        jc = ppr.find(qn('w:jc'))
        current = None if jc is None else jc.get(qn('w:val'))

        # A paragraph holding a line break has that line justified too, so a
        # heading kept together with the text under it gets pulled apart.
        short = len(paragraph.text) < JUSTIFY_FROM
        broken = bool(paragraph._p.findall('.//' + qn('w:br')))
        if short or broken:
            if current == 'both' and paragraph.text.strip():
                jc.set(qn('w:val'), 'left')
                relaxed += 1
            continue

        if current == 'both':
            continue
        if jc is None:
            jc = OxmlElement('w:jc')
            ppr.append(jc)
        jc.set(qn('w:val'), 'both')
        justified += 1
    return justified, relaxed


# --------------------------------------------------------------------------

def polish(docx_path, pdf_path=None, vocab_paths=(), justify=True):
    """Run every repair over a converted document. Returns a summary dict."""
    keep_breaks = deliberate_breaks(pdf_path) if pdf_path else set()
    swapped = substitute_fonts(docx_path)

    document = Document(docx_path)
    folios, breaks, blanks, kept, held = flow_pages(
        document, keep_breaks=keep_breaks)
    vocab = build_vocabulary(document, vocab_paths)
    spaces = restore_spaces(document, vocab)
    justified, relaxed = justify_body(document) if justify else (0, 0)
    document.save(docx_path)

    return {'fonts': swapped, 'page numbers': folios,
            'page breaks dropped': breaks, 'page breaks kept': kept,
            'blank paragraphs': blanks, 'figures held together': held,
            'spaces': spaces, 'justified': justified,
            'vocabulary': len(vocab)}
