"""Optional local Whisper backend: explicit download, fixed sources, SHA256 checks."""
import hashlib
import os
from pathlib import Path
import subprocess
import tempfile
import zipfile

import requests

TOOL_URL = 'https://github.com/ggml-org/whisper.cpp/releases/download/b5454/whisper-bin-x64.zip'
TOOL_HASH = '6ba69e3482d7826214f90a6a9c84ca07782aec1e1d0c6a7c30c994fd5d816ccb'
MODEL_URL = 'https://huggingface.co/ggerganov/whisper.cpp/resolve/5359861c739e955e79d9a303bcbc70fb988958b1/ggml-base.bin'
MODEL_HASH = '60ed5bc3dd14eea856493d334349b405782ddcaf0028d4b5df4088345fba2efe'


def data_dir():
    return Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'AIEPT' / 'whisper'


def executable(folder=None):
    folder = Path(folder) if folder else data_dir()
    found = sorted((folder / 'runtime').rglob('whisper-cli.exe'))
    return found[0] if found else None


def ready(folder=None):
    folder = Path(folder) if folder else data_dir()
    marker = folder / 'verified.txt'
    return bool(executable(folder) and (folder / 'ggml-base.bin').is_file()
                and marker.is_file() and marker.read_text() == TOOL_HASH + MODEL_HASH)


def _download(url, destination, digest, status):
    partial = destination.with_suffix(destination.suffix + '.partial')
    try:
        sha = hashlib.sha256()
        with requests.get(url, stream=True, timeout=(15, 60)) as response:
            response.raise_for_status()
            total = int(response.headers.get('Content-Length', 0))
            count = 0
            with partial.open('wb') as out:
                for block in response.iter_content(1024 * 1024):
                    if block:
                        count += len(block)
                        out.write(block)
                        sha.update(block)
                        status(f'下載 Whisper：{count // 1048576} MB' +
                               (f' / {total // 1048576} MB' if total else ''))
        if sha.hexdigest() != digest:
            raise RuntimeError('Whisper 下載校驗失敗，請重新下載。')
        partial.replace(destination)
    finally:
        partial.unlink(missing_ok=True)


def install(folder=None, status=lambda text: None):
    folder = Path(folder) if folder else data_dir()
    folder.mkdir(parents=True, exist_ok=True)
    archive = folder / 'runtime.zip'
    _download(TOOL_URL, archive, TOOL_HASH, status)
    runtime = folder / 'runtime'
    runtime.mkdir(exist_ok=True)
    with zipfile.ZipFile(archive) as bundle:
        root = runtime.resolve()
        for item in bundle.infolist():
            target = (root / item.filename).resolve()
            if not target.is_relative_to(root):
                raise RuntimeError('Invalid Whisper archive path')
        bundle.extractall(runtime)
    archive.unlink()
    _download(MODEL_URL, folder / 'ggml-base.bin', MODEL_HASH, status)
    if not executable(folder):
        raise RuntimeError('下載包沒有 whisper-cli.exe。')
    (folder / 'verified.txt').write_text(TOOL_HASH + MODEL_HASH)


def transcribe(audio_path, language, folder=None):
    folder = Path(folder) if folder else data_dir()
    if not ready(folder):
        raise RuntimeError('請先在偏好設定下載 Whisper 本機模型。')
    with tempfile.TemporaryDirectory(prefix='aiept-stt-') as temp:
        output = Path(temp) / 'transcript'
        result = subprocess.run(
            [str(executable(folder)), '-m', str(folder / 'ggml-base.bin'),
             '-f', str(audio_path), '-l', 'en' if language.startswith('en') else 'zh',
             '-otxt', '-of', str(output), '-nt', '-np', '-ng', '-bs', '5'],
            capture_output=True, timeout=300,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0,
        )
        text_file = output.with_suffix('.txt')
        if result.returncode or not text_file.is_file():
            raise RuntimeError('Whisper 辨識失敗，請確認模型與錄音裝置，或改用 Windows 本機辨識。')
        text = text_file.read_text(encoding='utf-8-sig').strip()
        if not text:
            raise RuntimeError('未辨識到語音；請靠近麥克風，確認收音音量後重試。')
        return text
