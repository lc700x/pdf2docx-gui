"""Compare a converted document against the PDF it came from, page by page.

A conversion can fail quietly. Text is dropped, a caption disappears into a
cropped table, spaces vanish from a justified line — and none of it raises an
error. This reads the PDF a page at a time and asks what became of that page's
words, then reports what is unaccounted for.

It compares content, not layout, so it needs nothing but the two files: where
the page breaks fall afterwards is a separate question, and one the reader can
judge by eye.

    python check.py source.pdf converted.docx

Exit status is 1 when anything is missing, so it can gate a batch.
"""

import re
import sys
import unicodedata
from difflib import SequenceMatcher

import fitz
from docx import Document
from docx.oxml.ns import qn

# A line that is mostly symbols — a formula, a row of figures — says little
# about whether text survived, and a crop is expected to absorb it.
WORD = re.compile(r"[^\W\d_](?:[^\W\d_]|['-])+", re.UNICODE)
MIN_RUN = 6            # words: shorter than this and a match means nothing
REPORT_LIMIT = 40


def normalise(text):
    """Fold the text to what both sides can be expected to agree on."""
    text = unicodedata.normalize('NFKC', text)
    return [w.lower() for w in WORD.findall(text)]


def docx_words(path):
    document = Document(path)
    parts = []
    body = document.element.body
    math_t = '{http://schemas.openxmlformats.org/officeDocument/2006/math}t'
    for node in body.iter():
        if node.tag in (qn('w:t'), math_t):
            parts.append(node.text or '')
        elif node.tag == qn('wp:docPr'):
            parts.append(node.get('descr') or '')
    return normalise(' '.join(parts))


def missing_runs(page_words, doc_words):
    """Stretches of a page's words that do not appear in the document."""
    matcher = SequenceMatcher(None, page_words, doc_words, autojunk=False)
    gaps = []
    for tag, i1, i2, _, _ in matcher.get_opcodes():
        if tag in ('delete', 'replace') and i2 - i1 >= MIN_RUN:
            gaps.append((i1, i2))
    return gaps


def check_text(pdf_path, docx_path):
    """Report every run of words the PDF has and the document does not."""
    doc_words = docx_words(docx_path)
    findings = []
    with fitz.open(pdf_path) as pdf:
        for page in pdf:
            folio = str(page.number + 1)
            text = '\n'.join(
                line for line in page.get_text().splitlines()
                if line.strip() != folio)
            page_words = normalise(text)
            if len(page_words) < MIN_RUN:
                continue
            for start, end in missing_runs(page_words, doc_words):
                findings.append((page.number + 1,
                                 ' '.join(page_words[start:end])))
    return findings


def check_spacing(docx_path):
    """Words that ran together, and sentences that lost the space after a stop."""
    document = Document(docx_path)
    text = ' '.join(p.text for p in document.paragraphs)
    vocab = {w.lower() for w in WORD.findall(text)}
    glued = [w for w in text.split()
             if len(w) > 24 and w.isalpha() and w.lower() not in vocab]
    stops = re.findall(r'[a-z]{2}\.[A-Z][a-z]', text)
    return glued, stops


def check_images(docx_path):
    document = Document(docx_path)
    body = document.element.body
    pictures = len(body.findall('.//' + qn('w:drawing')))
    tables = len(document.tables)
    return pictures, tables


def main(argv):
    # Some Windows consoles cannot encode names from the source document.
    # Escape only unrepresentable characters so diagnostics still complete.
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(errors='backslashreplace')
    if len(argv) < 2:
        print(__doc__.strip())
        return 2
    pdf_path, docx_path = argv[0], argv[1]

    pictures, tables = check_images(docx_path)
    print(f'{docx_path}: {pictures} images, {tables} tables')

    glued, stops = check_spacing(docx_path)
    if glued:
        print(f'\n{len(glued)} words with the spaces missing between them:')
        for word in glued[:10]:
            print(f'    {word}')
    if stops:
        print(f'\n{len(stops)} sentences with no space after the full stop, '
              f'for example: {", ".join(stops[:5])}')

    findings = check_text(pdf_path, docx_path)
    if not findings:
        print('\nEvery word on every page of the PDF is in the document.')
    else:
        print(f'\n{len(findings)} passages are in the PDF but not the document:')
        for page, words in findings[:REPORT_LIMIT]:
            shown = words if len(words) < 90 else words[:87] + '…'
            print(f'    p{page:<4} {shown}')
        if len(findings) > REPORT_LIMIT:
            print(f'    … and {len(findings) - REPORT_LIMIT} more')
        print('\nText inside a cropped table or equation is expected here: the '
              'crop is a picture, so its words are no longer text. Anything '
              'else is content the conversion lost.')

    return 1 if findings or glued or stops else 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
