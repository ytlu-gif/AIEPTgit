# Third-party components

AIEPT source code remains under its MIT LICENSE. The following independent components retain their own licenses.

## FFmpeg / imageio-ffmpeg

Pitch-preserving playback invokes FFmpeg as a separate process; it is not linked into AIEPT. The installer includes the unmodified FFmpeg 9.0.2 essentials build supplied by imageio-ffmpeg 0.8.1 (Gyan build), licensed GPL-3.0-or-later, without warranty. Full GPL text: licenses/FFmpeg-GPL-3.0.txt.

- Binary supplier, source reference, build information and libraries: https://www.gyan.dev/ffmpeg/builds/
- Exact FFmpeg source reference supplied by the builder: https://github.com/FFmpeg/FFmpeg/tree/946fcce07b
- FFmpeg source release: https://ffmpeg.org/releases/ffmpeg-9.0.2.tar.xz
- imageio-ffmpeg wrapper source: https://github.com/imageio/imageio-ffmpeg/tree/v0.8.1
- Wrapper license: licenses/imageio-ffmpeg-LICENSE.txt

When redistributing the installer, retain these notices and provide corresponding source code and build information for the GPL components under their license. Do not label third-party binaries as MIT-only. The upstream build source references above are provided for traceability; a downstream distributor remains responsible for satisfying the source-delivery requirements of the GPL.

## Optional Whisper local speech recognition

Whisper runtime and model are downloaded separately, only after the user requests download. No user recording is uploaded for local transcription.

- whisper.cpp b5454 (CPU x64): https://github.com/ggml-org/whisper.cpp/releases/tag/b5454
- Runtime license: licenses/whisper-cpp-LICENSE.txt (MIT)
- Model: https://huggingface.co/ggerganov/whisper.cpp (ggml-base.bin, OpenAI Whisper)
- Model license: licenses/Whisper-model-LICENSE.txt (MIT)

Downloads are checked against fixed SHA256 values recorded in local_whisper.py. The model is stored in the user's local application data directory, separate from source code and recordings.
