import pymupdf
import json

doc = pymupdf.open('TESIS DE CHANDA SOSA.pdf')

def extract_highlighted_spans(page):
    drawings = page.get_drawings()
    # Filter fill drawings that look like highlights
    # A highlight is typically a rectangle with a fill color and no stroke or stroke None
    highlight_rects = []
    for d in drawings:
        fill = d.get('fill')
        # Check if fill color is present and d has rect
        if fill is not None:
            r = pymupdf.Rect(d['rect'])
            # Exclude full page rects or header/footer bars
            if r.height < 60 and r.width > 10:
                highlight_rects.append((r, fill))
                
    if not highlight_rects:
        return []

    # Get text words or spans
    words = page.get_text('words') # (x0, y0, x1, y1, "word", block_no, line_no, word_no)
    highlighted_words = []
    for w in words:
        w_rect = pymupdf.Rect(w[0], w[1], w[2], w[3])
        # Skip header and footer of Turnitin
        if w[1] < 45 or w[3] > 750:
            continue
        for hr, col in highlight_rects:
            if hr.intersects(w_rect) and (hr & w_rect).get_area() > 0.4 * w_rect.get_area():
                highlighted_words.append((w[4], col, w[1], w[2], w[5], w[6]))
                break
                
    return highlighted_words

all_highlights = {}
for p_num in range(5, len(doc)):
    page = doc[p_num]
    hw = extract_highlighted_spans(page)
    if hw:
        # reconstruct lines/text
        text_str = " ".join([w[0] for w in hw])
        all_highlights[p_num + 1] = {
            'count': len(hw),
            'text': text_str
        }

print(f"Total pages with highlighted text: {len(all_highlights)}")
with open('turnitin_extracted_matches.json', 'w', encoding='utf-8') as f:
    json.dump(all_highlights, f, ensure_ascii=False, indent=2)

for p, data in list(all_highlights.items())[:25]:
    print(f"Page {p} ({data['count']} words): {data['text'][:120]}...")
