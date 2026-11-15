from pathlib import Path
import runpy
import tempfile
import unittest
from unittest.mock import Mock, patch
import wave

import read_along

m = runpy.run_path(str(Path(__file__).resolve().parents[1] / 'app.py'), run_name='read_along_tests')
C = m['FlashcardApp']


class TimingTests(unittest.TestCase):
    def test_sentences_keep_titles_decimals_and_quotes(self):
        text = 'Dr. Lee paid 3.50 dollars. "Is that correct?" Yes, it is.'
        self.assertEqual(read_along.sentences(text), ['Dr. Lee paid 3.50 dollars.', '"Is that correct?"', 'Yes, it is.'])

    def test_estimated_boundaries_cover_duration(self):
        items = read_along.estimate_timings('A short sentence. This sentence has a few more words.', 10000)
        self.assertEqual(items[0]['start_ms'], 0)
        self.assertEqual(items[-1]['end_ms'], 10000)
        self.assertEqual(items[0]['end_ms'], items[1]['start_ms'])

    def test_metadata_rejects_changed_audio(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'audio.wav'
            path.write_bytes(b'original')
            segments = [{'text': 'A sentence.', 'start_ms': 0, 'end_ms': 1000}]
            read_along.save_timings(path, segments)
            self.assertEqual(read_along.load_timings(path, 1000), segments)
            path.write_bytes(b'changed')
            self.assertIsNone(read_along.load_timings(path, 1000))

    @unittest.skipUnless(m['sys'].platform == 'win32', 'Windows SAPI integration')
    def test_real_sapi_sentence_offsets_match_wave_duration(self):
        with tempfile.TemporaryDirectory() as folder:
            app = C.__new__(C)
            app.logger = Mock()
            path = app.generate_windows_audio('This is the first sentence. Here is another sentence.',
                folder, 'timing_test', 'en-US-JennyNeural', RuntimeError('local test'))
            with wave.open(path) as audio:
                duration = round(audio.getnframes() * 1000 / audio.getframerate())
            items = read_along.load_timings(path, duration)
            self.assertIsNotNone(items)
            self.assertEqual(len(items), 2)
            self.assertEqual(items[0]['start_ms'], 0)
            self.assertEqual(items[0]['end_ms'], items[1]['start_ms'])
            self.assertAlmostEqual(items[-1]['end_ms'], duration, delta=1)


class HighlightTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        folder = Path(self.temp.name)
        C.__init__.__globals__['CONFIG_FILE'] = str(folder / 'config.json')
        C.__init__.__globals__['CHAT_HISTORY_FILE'] = str(folder / 'history.json')
        self.root = m['tk'].Tk()
        self.root.withdraw()
        self.app = C(self.root)
        self.first, self.second = 'The first English sentence.', 'The second English sentence.'
        self.app.render_learning_card_preview({'title': 'Test', 'article': {
            'en': [self.first, self.second], 'zh': ['第一句中文。', '第二句中文。']}, 'words': []})
        self.path = folder / 'article.wav'
        with wave.open(str(self.path), 'wb') as audio:
            audio.setparams((1, 2, 16000, 0, 'NONE', 'not compressed'))
            audio.writeframes(b'\x00\x00' * 16000 * 6)
        self.app.set_generated_audio(str(self.path))
        self.app.audio_duration_ms = 6000

    def tearDown(self):
        self.app.stop_generated_audio(update_status=False)
        self.root.destroy()
        self.temp.cleanup()

    def highlighted(self):
        ranges = self.app.card_preview_text.tag_ranges('read-along')
        return self.app.card_preview_text.get(*ranges) if ranges else ''

    def test_sentence_highlight_toggle_stop_and_finish(self):
        read_along.save_timings(self.path, [
            {'text': self.first, 'start_ms': 0, 'end_ms': 3000},
            {'text': self.second, 'start_ms': 3000, 'end_ms': 6000}])
        app = self.app
        app.prepare_read_along()
        self.assertEqual(len(app.read_along_ranges), 2)
        self.assertIn('句子時間點', app.read_along_mode_var.get())
        app.update_audio_progress(500)
        self.assertEqual(self.highlighted(), self.first)
        app.update_audio_progress(3200)
        self.assertEqual(self.highlighted(), self.second)
        app.read_along_var.set(False)
        app.toggle_read_along()
        self.assertEqual(self.highlighted(), '')
        app.read_along_var.set(True)
        app.update_audio_progress(3200)
        self.assertEqual(self.highlighted(), self.second)
        app.stop_generated_audio()
        self.assertEqual(self.highlighted(), '')
        app.update_audio_progress(6000)
        self.assertEqual(self.highlighted(), '')

    def test_legacy_recording_is_labeled_estimated(self):
        self.app.prepare_read_along()
        self.assertIn('估算', self.app.read_along_mode_var.get())
        self.app.update_audio_progress(100)
        self.assertEqual(self.highlighted(), self.first)

    def test_speed_adjusted_poll_uses_original_timeline(self):
        read_along.save_timings(self.path, [
            {'text': self.first, 'start_ms': 0, 'end_ms': 3000},
            {'text': self.second, 'start_ms': 3000, 'end_ms': 6000}])
        app = self.app
        app.prepare_read_along()
        app.audio_device_alias = 'fake_device'
        app.audio_speed_factor = 0.75
        def command(text, return_output=False):
            return 'paused' if text.endswith('mode') else '4400'
        with patch.object(app, 'send_audio_command', side_effect=command):
            app.poll_generated_audio()
        self.assertEqual(app.audio_progress_var.get(), 3300)
        self.assertEqual(self.highlighted(), self.second)
        app.audio_device_alias = None


if __name__ == '__main__':
    unittest.main()
