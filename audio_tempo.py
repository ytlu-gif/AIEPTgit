"""Pitch-preserving playback copies. Original recordings are never modified."""
import hashlib
import os
from pathlib import Path
import subprocess
import tempfile
import threading

import imageio_ffmpeg


class AudioTempoCache:
    def __init__(self):
        self._temp = tempfile.TemporaryDirectory(prefix='aiept-tempo-')
        self._lock = threading.Lock()

    def render(self, source, factor):
        if factor not in (0.5, 0.75, 1.0, 1.25, 1.5):
            raise ValueError('Unsupported playback speed')
        source = Path(source).resolve()
        if factor == 1:
            return str(source)
        stat = source.stat()
        key = hashlib.sha256(f'{source}:{stat.st_mtime_ns}:{stat.st_size}:{factor}'.encode()).hexdigest()
        target = Path(self._temp.name) / (key + '.wav')
        with self._lock:
            if target.is_file():
                return str(target)
            result = subprocess.run(
                [imageio_ffmpeg.get_ffmpeg_exe(), '-hide_banner', '-loglevel', 'error',
                 '-nostdin', '-y', '-i', str(source), '-vn', '-af', f'atempo={factor}',
                 '-c:a', 'pcm_s16le', str(target)],
                capture_output=True, timeout=180,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0,
            )
            if result.returncode or not target.is_file():
                target.unlink(missing_ok=True)
                raise RuntimeError('無法產生保留音調的變速音訊，請確認音檔可讀取。')
        return str(target)
