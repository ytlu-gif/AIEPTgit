import runpy
import tempfile
import time
import unittest
import wave
from pathlib import Path
from unittest.mock import patch

m = runpy.run_path(str(Path(__file__).resolve().parents[1] / 'app.py'), run_name='progress_tests')
C = m['FlashcardApp']


class ProgressAndImportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)
        C.__init__.__globals__['CONFIG_FILE'] = str(self.folder / 'config.json')
        C.__init__.__globals__['CHAT_HISTORY_FILE'] = str(self.folder / 'history.json')
        self.root = m['tk'].Tk()
        self.root.withdraw()
        self.app = C(self.root)

    def tearDown(self):
        self.app.stop_generated_audio(update_status=False)
        self.root.destroy()
        self.temp.cleanup()

    def test_imported_card_enables_correct_conversation_without_audio(self):
        document = m['Document']()
        document.add_paragraph('Blood sugar monitoring helps manage diabetes.')
        path = self.folder / 'bloodsugar_Flashcard.docx'
        document.save(path)
        self.app.topic_entry.insert(0, 'Unrelated workplace topic')
        self.app.practice_context = {'content': {'words': [{'word': 'old'}]}}
        with patch.object(m['filedialog'], 'askopenfilename', return_value=str(path)):
            self.app.import_historical_card()
        self.assertEqual(str(self.app.practice_btn['state']), 'normal')
        self.assertIsNone(self.app.generated_audio_path)
        self.assertEqual(self.app.practice_context['provider'], 'Ollama')
        self.assertEqual(self.app.practice_context['api_key'], '')
        self.assertEqual(self.app.practice_context['keywords'], [])
        self.assertIn('bloodsugar', self.app.practice_context['topic'])
        self.assertIn('Blood sugar monitoring', self.app.build_practice_system_prompt())
        record = self.app.create_chat_record()
        self.assertTrue(self.app.restore_practice_context(record))
        self.assertIn('Blood sugar monitoring', self.app.build_practice_system_prompt())

    @unittest.skipUnless(m['sys'].platform == 'win32', 'Windows playback integration')
    def test_progress_pause_rate_stop_and_end(self):
        path = self.folder / 'silent.wav'
        with wave.open(str(path), 'wb') as audio:
            audio.setparams((1, 2, 16000, 0, 'NONE', 'not compressed'))
            audio.writeframes(b'\x00\x00' * 16000 * 5)
        app = self.app
        with patch.object(m['messagebox'], 'showerror') as errors:
            app.set_generated_audio(str(path))
            app.play_generated_audio()
            self.assertEqual(app.audio_duration_ms, 5000)
            time.sleep(0.3)
            app.poll_generated_audio()
            self.assertGreater(app.audio_progress_var.get(), 0)
            self.assertTrue(app.audio_time_var.get().endswith(' / 00:05'))
            app.pause_generated_audio()
            app.poll_generated_audio()
            before = app.audio_progress_var.get()
            app.playback_rate_var.set('0.5 倍（很慢）')
            app.change_playback_rate()
            time.sleep(0.2)
            app.poll_generated_audio()
            self.assertEqual(before, app.audio_progress_var.get())
            app.resume_generated_audio()
            app.stop_generated_audio()
            self.assertEqual(app.audio_progress_var.get(), 0)
            self.assertEqual(app.audio_time_var.get(), '00:00 / 00:05')
            app.play_generated_audio()
            app.send_audio_command(f'play {app.audio_device_alias} from 4950')
            deadline = time.monotonic() + 3
            while app.audio_device_alias and time.monotonic() < deadline:
                self.root.update()
                time.sleep(0.05)
            self.assertIsNone(app.audio_device_alias)
            self.assertEqual(app.audio_progress_var.get(), 5000)
            self.assertEqual(app.audio_time_var.get(), '00:05 / 00:05')
            errors.assert_not_called()


if __name__ == '__main__':
    unittest.main()
