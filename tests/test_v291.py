import json
import math
from pathlib import Path
import runpy
import struct
import tempfile
import time
import unittest
from unittest.mock import Mock, patch
import wave

from audio_tempo import AudioTempoCache
import local_whisper

m = runpy.run_path(str(Path(__file__).resolve().parents[1] / 'app.py'), run_name='v291_tests')
C = m['FlashcardApp']


def make_audio(path, seconds=3, tone=False):
    samples = [int(12000 * math.sin(2 * math.pi * 440 * n / 16000)) if tone else 0
               for n in range(int(seconds * 16000))]
    with wave.open(str(path), 'wb') as out:
        out.setparams((1, 2, 16000, 0, 'NONE', 'not compressed'))
        out.writeframes(struct.pack(f'<{len(samples)}h', *samples))


class ContentTests(unittest.TestCase):
    def test_three_lengths_have_distinct_prompts(self):
        for label, (minimum, maximum, paragraphs, tokens) in m['ARTICLE_LENGTHS'].items():
            prompt = m['build_learning_prompt']('Travel', ['station'], '旅遊', label)
            self.assertIn(f'{minimum}～{maximum}', prompt)
            self.assertIn(f'{paragraphs} 段', prompt)

    def test_long_response_budget_and_local_only(self):
        app = C.__new__(C)
        response = Mock(status_code=200)
        response.json.return_value = {'message': {'content': 'hello'}}
        session = Mock()
        session.__enter__ = Mock(return_value=session)
        session.__exit__ = Mock(return_value=False)
        session.post.return_value = response
        with patch.object(m['requests'], 'Session', return_value=session):
            app.call_ollama([], structured=True, output_tokens=10240)
        kwargs = session.post.call_args.kwargs
        self.assertEqual(kwargs['json']['options']['num_predict'], 10240)
        self.assertGreaterEqual(kwargs['json']['options']['num_ctx'], 10240 + 4096)
        self.assertFalse(session.trust_env)
        self.assertFalse(kwargs['allow_redirects'])


class TempoTests(unittest.TestCase):
    def test_duration_changes_without_pitch_change(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / 'tone.wav'
            make_audio(source, tone=True)
            original = source.read_bytes()
            cache = AudioTempoCache()
            try:
                for factor in (0.5, 0.75, 1.25, 1.5):
                    rendered = cache.render(source, factor)
                    with wave.open(rendered, 'rb') as audio:
                        rate = audio.getframerate()
                        length = audio.getnframes() / rate
                        samples = struct.unpack(f'<{audio.getnframes()}h', audio.readframes(audio.getnframes()))
                    self.assertAlmostEqual(length, 3 / factor, delta=0.15)
                    window = samples[rate // 2: min(len(samples), rate * 2)]
                    crossings = sum(a <= 0 < b for a, b in zip(window, window[1:]))
                    pitch = crossings * rate / len(window)
                    self.assertAlmostEqual(pitch, 440, delta=3)
                    self.assertEqual(cache.render(source, factor), rendered)
                self.assertEqual(source.read_bytes(), original)
            finally:
                cache._temp.cleanup()


class UITests(unittest.TestCase):
    def setUp(self):
        self.folder = tempfile.TemporaryDirectory()
        C.__init__.__globals__['CONFIG_FILE'] = str(Path(self.folder.name) / 'config.json')
        C.__init__.__globals__['CHAT_HISTORY_FILE'] = str(Path(self.folder.name) / 'history.json')
        self.root = m['tk'].Tk()
        self.root.withdraw()
        self.app = C(self.root)

    def tearDown(self):
        self.app.stop_generated_audio(update_status=False)
        if self.app.audio_tempo_cache:
            self.app.audio_tempo_cache._temp.cleanup()
        self.root.destroy()
        self.folder.cleanup()

    def wait_for_rate(self):
        deadline = time.monotonic() + 8
        def check():
            if not self.app.audio_rate_pending or time.monotonic() >= deadline:
                self.root.quit()
            else:
                self.root.after(20, check)
        self.root.after(20, check)
        self.root.mainloop()
        self.assertFalse(self.app.audio_rate_pending)

    @unittest.skipUnless(m['sys'].platform == 'win32', 'Windows playback integration')
    def test_pitch_safe_playback_pause_position_and_resume(self):
        path = Path(self.folder.name) / 'silent.wav'
        make_audio(path, seconds=12)
        app = self.app
        with patch.object(m['messagebox'], 'showerror') as errors:
            app.set_generated_audio(str(path))
            app.play_generated_audio()
            app.send_audio_command(f'play {app.audio_device_alias} from 3000')
            app.pause_generated_audio()
            original_position = int(app.send_audio_command(f'status {app.audio_device_alias} position', True))
            app.playback_rate_var.set('0.75 倍（慢速）')
            app.change_playback_rate()
            self.wait_for_rate()
            self.assertEqual(app.audio_speed_factor, 0.75)
            self.assertEqual(app.send_audio_command(f'status {app.audio_device_alias} speed', True), '1000')
            self.assertEqual(app.send_audio_command(f'status {app.audio_device_alias} mode', True), 'paused')
            position = int(app.send_audio_command(f'status {app.audio_device_alias} position', True))
            # MCI seek/pause positions are quantized by its audio buffer; reject restarts,
            # allow under one second of device alignment latency.
            self.assertAlmostEqual(position * 0.75, original_position, delta=750)
            self.assertEqual(app.audio_duration_ms, 12000)
            app.resume_generated_audio()
            app.playback_rate_var.set('1.25 倍（快速）')
            app.change_playback_rate()
            self.wait_for_rate()
            self.assertEqual(app.audio_speed_factor, 1.25)
            self.assertEqual(app.send_audio_command(f'status {app.audio_device_alias} speed', True), '1000')
            self.assertEqual(app.send_audio_command(f'status {app.audio_device_alias} mode', True), 'playing')
            errors.assert_not_called()

    def test_local_mode_does_not_use_cloud_stt(self):
        self.app.speech_provider_var.set('OpenAI 雲端')
        self.app.microphone_var.set('2: Test microphone')
        with patch.object(m['local_whisper'], 'ready', return_value=True), patch.object(m['threading'], 'Thread') as thread:
            self.app.start_speech_input(language='en-US')
        self.assertEqual(thread.call_args.kwargs['args'][1:], ('Whisper', '', 'en-US'))
        self.assertEqual(self.app.active_microphone, '2: Test microphone')


class WhisperTests(unittest.TestCase):
    def test_corrupt_download_is_rejected_and_cleaned(self):
        with tempfile.TemporaryDirectory() as folder:
            destination = Path(folder) / 'model.bin'
            response = Mock(headers={})
            response.__enter__ = Mock(return_value=response)
            response.__exit__ = Mock(return_value=False)
            response.iter_content.return_value = [b'bad model']
            with patch.object(local_whisper.requests, 'get', return_value=response):
                with self.assertRaisesRegex(RuntimeError, '校驗'):
                    local_whisper._download('https://example.invalid/model', destination, '0' * 64, lambda text: None)
            self.assertFalse(destination.exists())
            self.assertFalse(destination.with_suffix('.bin.partial').exists())

    def test_transcription_uses_local_process_and_explicit_language(self):
        with tempfile.TemporaryDirectory() as folder:
            def run(args, **kwargs):
                output = Path(args[args.index('-of') + 1]).with_suffix('.txt')
                output.write_text('A practice sentence.', encoding='utf-8')
                self.assertEqual(args[args.index('-l') + 1], 'en')
                return Mock(returncode=0)
            with patch.object(local_whisper, 'ready', return_value=True), patch.object(local_whisper, 'executable', return_value=Path(folder) / 'whisper-cli.exe'), patch.object(local_whisper.subprocess, 'run', side_effect=run), patch.object(local_whisper.requests, 'get') as network:
                text = local_whisper.transcribe('input.wav', 'en-US', folder)
            self.assertEqual(text, 'A practice sentence.')
            network.assert_not_called()


if __name__ == '__main__':
    unittest.main()
