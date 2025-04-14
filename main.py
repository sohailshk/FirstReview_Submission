import re
import sys
import time
import string
import win32com.client as win32
import pythoncom

# Regex to match text in quoted segments (both curly and straight quotes)
QUOTE_PATTERN = r'(“[^”]*”|"[^"]*")'


def full_normalize_text(text):
    """
    Thoroughly normalize text to avoid hidden formatting issues.
    1. Convert curly quotes to straight quotes.
    2. Remove or replace non-printable characters.
    3. Collapse extra whitespace.
    """
    # Convert curly quotes to straight
    text = text.replace('“', '"').replace('”', '"')

    # Remove non-printable chars using translation
    printable = set(string.printable)
    text = ''.join(ch for ch in text if ch in printable)

    # Collapse multiple spaces
    text = re.sub(r'[^\S\r\n]+', ' ', text)
    return text


def correct_language_spelling(text):
    """
    Applies UK spelling corrections in unquoted segments only.
    """
    segments = re.split(QUOTE_PATTERN, text, flags=re.DOTALL)
    for i in range(0, len(segments), 2):
        seg = segments[i]
        seg = re.sub(r'\borganize\b', 'organise', seg, flags=re.IGNORECASE)
        seg = re.sub(r'\beg\b', 'for example', seg, flags=re.IGNORECASE)
        segments[i] = seg
    return "".join(segments)


def _extract_names(segment):
    """
    Finds all occurrences of (Title First [MiddleInitial?] Last) with standard capitalization.
    Returns list of (start_index, end_index, title, first, middle, last).
    """
    pattern = r'\b(Dr|Mr|Mrs|Ms)\s+([A-Z][a-z]+)(?:\s+([A-Z]\.?))?\s+([A-Z][a-z]+)(?=[\s\.,;!?:]|$)'
    return [
        (m.start(), m.end(), m.group(1), m.group(2), m.group(3) or '', m.group(4))
        for m in re.finditer(pattern, segment)
    ]


def _correct_names_segment(segment):
    """
    Multi-pass approach: fix missing periods in middle initials, then track name appearances.
    """
    # First fix patterns like "Franklin D Roosevelt" -> "Franklin D. Roosevelt"
    segment = re.sub(
        r'\b([A-Z][a-z]+)\s+([A-Z])(?!\.)\s+([A-Z][a-z]+)\b',
        lambda m: f"{m.group(1)} {m.group(2)}. {m.group(3)}",
        segment
    )

    # Extract name occurrences
    matches = _extract_names(segment)
    if not matches:
        return segment

    # Identify ambiguous last names (shared by different first names)
    names_by_last = {}
    for _, _, title, first, middle, last in matches:
        full = f"{title} {first}" + (f" {middle}" if middle else "") + f" {last}"
        names_by_last.setdefault(last, set()).add(full)
    ambiguous_lasts = {ln for ln, fulls in names_by_last.items() if len(fulls) > 1}

    # Replace occurrences
    result = []
    prev_end = 0
    name_counts = {}
    for start, end, title, first, middle, last in matches:
        result.append(segment[prev_end:start])

        full_name = f"{title} {first}" + (f" {middle}" if middle else "") + f" {last}"
        current_count = name_counts.get(full_name, 0)
        name_counts[full_name] = current_count + 1

        # If the last name is ambiguous, always use full name
        if last in ambiguous_lasts:
            replacement = full_name
        else:
            # First mention = full name; subsequent mentions = "Title Last"
            replacement = full_name if current_count == 0 else f"{title} {last}"

        result.append(replacement)
        prev_end = end

    result.append(segment[prev_end:])
    return "".join(result)


def correct_names(text):
    """
    Correct names in unquoted segments only.
    """
    segments = re.split(QUOTE_PATTERN, text, flags=re.DOTALL)
    for i in range(0, len(segments), 2):
        segments[i] = _correct_names_segment(segments[i])
    return "".join(segments)


def process_text(text):
    """
    Overall pipeline.
    1. Thorough normalization.
    2. Spelling corrections in unquoted segments.
    3. Name replacements in unquoted segments.
    4. Prefix output with "Corrected:".
    """
    text = full_normalize_text(text)
    text = correct_language_spelling(text)
    text = correct_names(text)
    return "Corrected:\n" + text


def get_doc_text(doc, retries=5, delay=1):
    """
    Safely retrieve text from a Word Document, retrying if Word is busy.
    """
    for _ in range(retries):
        try:
            return doc.Content.Text
        except:
            time.sleep(delay)
    raise Exception("Could not retrieve document text after multiple attempts.")


def process_docx(input_path, output_path):
    """
    Opens DOCX with Word COM, processes text, and saves a new DOCX.
    """
    pythoncom.CoInitialize()
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(input_path)
        original_text = get_doc_text(doc)
        doc.Close(False)

        new_text = process_text(original_text)

        new_doc = word.Documents.Add()
        new_doc.Content.Text = new_text
        new_doc.SaveAs(output_path)
        new_doc.Close()
    finally:
        word.Quit()


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python main.py <input.docx> <output.docx>")
        sys.exit(1)
    input_path = sys.argv[1]
    output_path = sys.argv[2]
    process_docx(input_path, output_path)
    print("Done")
