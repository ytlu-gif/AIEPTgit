"""Sentence read-along metadata; does not infer exact times for legacy recordings."""
import hashlib
import json
from pathlib import Path
import re


def sentences(text):
    # Keep common titles, initials, and decimal numbers inside the sentence.
    result, start = [], 0
    for match in re.finditer(r'[.!?]+[\"\u201d\u2019\')]*\s+|\n+', text):
        candidate = text[start:match.end()].strip()
        prefix = text[:match.start() + 1]
        if re.search(r'\b(?:Mr|Mrs|Ms|Dr|Prof|Sr|Jr|vs|etc)\.$', prefix, re.I):
            continue
        if re.search(r'\b[A-Z]\.$', prefix) and candidate.count('.') <= 1:
            continue
        if candidate:
            result.append(candidate)
        start = match.end()
    tail = text[start:].strip()
    if tail:
        result.append(tail)
    return result


def sidecar_path(audio_path):
    return Path(str(audio_path) + '.timings.json')


def save_timings(audio_path, segments):
    source = Path(audio_path)
    payload = {'version': 1, 'audio_sha256': hashlib.sha256(source.read_bytes()).hexdigest(),
               'segments': segments}
    target = sidecar_path(source)
    partial = target.with_suffix(target.suffix + '.partial')
    try:
        partial.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
        partial.replace(target)
    finally:
        partial.unlink(missing_ok=True)


def load_timings(audio_path, duration_ms):
    path = sidecar_path(audio_path)
    if not path.is_file() or path.stat().st_size > 2_000_000:
        return None
    try:
        data = json.loads(path.read_text(encoding='utf-8'))
        if data.get('version') != 1 or data.get('audio_sha256') != hashlib.sha256(Path(audio_path).read_bytes()).hexdigest():
            return None
        segments = data['segments']
        if not isinstance(segments, list) or not 0 < len(segments) <= 1000:
            return None
        previous = 0
        for item in segments:
            if (not isinstance(item, dict) or not isinstance(item.get('text'), str)
                    or not item['text'].strip()
                    or not isinstance(item.get('start_ms'), (int, float))
                    or not isinstance(item.get('end_ms'), (int, float))
                    or not previous <= item['start_ms'] < item['end_ms'] <= duration_ms + 250):
                return None
            previous = item['end_ms']
        return segments
    except (OSError, ValueError, TypeError, KeyError):
        return None


def estimate_timings(text, duration_ms):
    parts = sentences(text)
    weights = [max(1, len(re.findall(r"[A-Za-z]+(?:['’-][A-Za-z]+)*", part))) + 2 for part in parts]
    total = sum(weights)
    elapsed = 0
    result = []
    for part, weight in zip(parts, weights):
        start = round(duration_ms * elapsed / total)
        elapsed += weight
        result.append({'text': part, 'start_ms': start, 'end_ms': round(duration_ms * elapsed / total)})
    return result
