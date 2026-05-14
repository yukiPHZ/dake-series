from __future__ import annotations


def split_text(text: str, chunk_size: int = 2400, overlap: int = 240) -> list[str]:
    clean = text.replace("\r\n", "\n").replace("\r", "\n").strip()
    if not clean:
        return []
    if len(clean) <= chunk_size:
        return [clean]

    chunks: list[str] = []
    start = 0
    text_length = len(clean)

    while start < text_length:
        hard_end = min(start + chunk_size, text_length)
        end = hard_end
        if hard_end < text_length:
            paragraph_break = clean.rfind("\n\n", start, hard_end)
            line_break = clean.rfind("\n", start, hard_end)
            sentence_break = max(
                clean.rfind("。", start, hard_end),
                clean.rfind(".", start, hard_end),
            )
            best_break = max(paragraph_break, line_break, sentence_break)
            if best_break > start + max(300, chunk_size // 3):
                end = best_break + 1

        chunk = clean[start:end].strip()
        if chunk:
            chunks.append(chunk)

        if end >= text_length:
            break
        start = max(end - overlap, start + 1)

    return chunks
