from pathlib import Path
import runpy
import unittest
from unittest.mock import Mock, patch

import update_checker

m = runpy.run_path(str(Path(__file__).resolve().parents[1] / 'app.py'), run_name='update_tests')
C = m['FlashcardApp']


class UpdateChecks(unittest.TestCase):
    def request(self, code=200, data=None):
        session = Mock()
        session.__enter__ = Mock(return_value=session)
        session.__exit__ = Mock(return_value=False)
        response = Mock(status_code=code)
        response.json.return_value = data
        session.get.return_value = response
        return session

    def test_versions_compare_numerically(self):
        self.assertGreater(update_checker.version_tuple('v2.10.0'), update_checker.version_tuple('2.9.3'))
        self.assertEqual(update_checker.version_tuple('v3.0'), (3, 0, 0))
        for invalid in ('v2.9.4-beta', 'https://example.invalid', None, 'latest'):
            self.assertIsNone(update_checker.version_tuple(invalid))

    def test_daily_schedule_and_opt_out(self):
        self.assertFalse(update_checker.should_check(False, None, 100000))
        self.assertTrue(update_checker.should_check(True, None, 100000))
        self.assertFalse(update_checker.should_check(True, 100000, 100500))
        self.assertTrue(update_checker.should_check(True, 100000, 186400))
        self.assertTrue(update_checker.should_check(True, 200000, 100000))

    def test_new_release_is_public_read_only_and_url_is_canonical(self):
        session = self.request(data={'tag_name': 'v2.10.0', 'html_url': 'https://example.invalid/malicious'})
        with patch.object(update_checker.requests, 'Session', return_value=session):
            result = update_checker.check_latest('2.9.3')
        self.assertEqual(result['status'], 'available')
        self.assertEqual(result['url'], 'https://github.com/ytlu-gif/AIEPTgit/releases/tag/v2.10.0')
        self.assertEqual(session.get.call_args.args, (update_checker.LATEST_URL,))
        kwargs = session.get.call_args.kwargs
        self.assertFalse(kwargs['allow_redirects'])
        self.assertFalse(session.trust_env)
        self.assertNotIn('Authorization', kwargs['headers'])
        self.assertNotIn('data', kwargs)
        self.assertNotIn('json', kwargs)

    def test_same_or_older_release_does_not_notify(self):
        for tag in ('v2.9.3', 'v2.7.0'):
            session = self.request(data={'tag_name': tag})
            with patch.object(update_checker.requests, 'Session', return_value=session):
                self.assertEqual(update_checker.check_latest('2.9.3')['status'], 'current')

    def test_missing_draft_and_prerelease_are_skipped(self):
        for code, data in ((404, None), (200, {'draft': True}), (200, {'prerelease': True})):
            session = self.request(code, data)
            with patch.object(update_checker.requests, 'Session', return_value=session):
                self.assertEqual(update_checker.check_latest('2.9.3')['status'], 'no_release')

    def test_rate_limit_network_and_bad_tags(self):
        for code, data in ((403, None), (429, None), (200, {'tag_name': 'latest'})):
            session = self.request(code, data)
            with patch.object(update_checker.requests, 'Session', return_value=session):
                with self.assertRaises(RuntimeError):
                    update_checker.check_latest('2.9.3')
        with patch.object(update_checker.requests, 'Session', side_effect=update_checker.requests.ConnectionError()):
            with self.assertRaisesRegex(RuntimeError, '網路'):
                update_checker.check_latest('2.9.3')


class NotificationTests(unittest.TestCase):
    def setUp(self):
        self.app = C.__new__(C)
        self.app.config = {}
        self.app.persist_update_state = Mock(side_effect=lambda **x: self.app.config.update(x))

    def test_automatic_notice_deduplicates_but_manual_check_can_repeat(self):
        result = {'status': 'available', 'tag': 'v2.10.0',
                  'url': 'https://github.com/ytlu-gif/AIEPTgit/releases/tag/v2.10.0'}
        with patch.object(m['messagebox'], 'askyesno', return_value=False) as ask, patch.object(m['webbrowser'], 'open') as browser:
            self.app.finish_update_check(result, None, False)
            self.app.finish_update_check(result, None, False)
            self.assertEqual(ask.call_count, 1)
            browser.assert_not_called()
            self.app.finish_update_check(result, None, True)
            self.assertEqual(ask.call_count, 2)

    def test_browser_opens_only_after_user_accepts(self):
        result = {'status': 'available', 'tag': 'v2.10.0',
                  'url': 'https://github.com/ytlu-gif/AIEPTgit/releases/tag/v2.10.0'}
        with patch.object(m['messagebox'], 'askyesno', return_value=True), patch.object(m['webbrowser'], 'open') as browser:
            self.app.finish_update_check(result, None, True)
            browser.assert_called_once_with(result['url'])

    def test_offline_background_check_is_quiet(self):
        with patch.object(m['messagebox'], 'showinfo') as info:
            self.app.finish_update_check(None, 'offline', False)
            info.assert_not_called()
            self.app.finish_update_check(None, 'offline', True)
            info.assert_called_once()


if __name__ == '__main__':
    unittest.main()
