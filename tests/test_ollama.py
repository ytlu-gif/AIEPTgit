import runpy, json, unittest, types
from pathlib import Path
from unittest.mock import patch, Mock
m=runpy.run_path(str(Path(__file__).resolve().parents[1] / 'app.py'),run_name='ollama_tests')
C=m['FlashcardApp']
C.__init__.__globals__['CONFIG_FILE'] = str(Path('work/ollama-test-config-missing.json').resolve())
C.__init__.__globals__['CHAT_HISTORY_FILE'] = str(Path('work/ollama-test-history-missing.json').resolve())
def sample(schema):
 if schema['type']=='object': return {k:sample(v) for k,v in schema['properties'].items()}
 if schema['type']=='array': return [sample(schema['items'])]
 return 'test'
card=sample(m['LOCAL_CARD_SCHEMA'])
card['article']={'en':['A sample article.'],'zh':['測試文章。']}
class RoutingTests(unittest.TestCase):
 def setUp(self):
  self.app=C.__new__(C)
  self.session=Mock(); self.session.__enter__=Mock(return_value=self.session); self.session.__exit__=Mock(return_value=False)
  self.response=Mock(status_code=200); self.response.json.return_value={'message':{'content':json.dumps(card)},'done_reason':'stop'}
  self.session.post.return_value=self.response
  self.patcher=patch.object(m['requests'],'Session',return_value=self.session); self.patcher.start()
 def tearDown(self): self.patcher.stop()
 def test_local_card_no_key(self):
  result=self.app.call_ai_api('Ollama','','Test','word','日常生活')
  self.assertEqual(result,card)
  args,kw=self.session.post.call_args
  self.assertEqual(args[0],'http://127.0.0.1:11434/api/chat')
  self.assertEqual(kw['json']['model'],'qwen3:4b')
  self.assertFalse(kw['json']['stream']); self.assertFalse(kw['json']['think'])
  self.assertIn('format',kw['json']); self.assertNotIn('headers',kw)
  self.assertFalse(kw['allow_redirects']); self.assertFalse(self.session.trust_env)
 def test_local_chat_no_key(self):
  self.response.json.return_value={'message':{'content':'Hello!'}}
  self.assertEqual(self.app.call_chat_api('Ollama','','Tutor',[{'role':'user','content':'Hi'}]),'Hello!')
  self.assertNotIn('format',self.session.post.call_args.kwargs['json'])
 def test_errors_never_fallback(self):
  for status in (302,404,500):
   with self.subTest(status=status):
    self.response.status_code=status
    with self.assertRaises(RuntimeError): self.app.call_chat_api('Ollama','','Tutor',[])
  self.assertEqual(self.session.post.call_count,3)
 def test_timeout_help(self):
  self.session.post.side_effect=m['requests'].Timeout()
  with self.assertRaisesRegex(RuntimeError,'超時'): self.app.call_ollama([])
 def test_incomplete_card(self):
  self.response.json.return_value={'message':{'content':'{"words":[{"word":"test","kk":"test"}]}'}}
  with self.assertRaises(ValueError): self.app.call_ai_api('Ollama','','Test','word','日常生活')
 def test_cloud_requires_key(self):
  with self.assertRaises(ValueError): self.app.call_ai_api('OpenRouter','','Test','word','日常生活')
  with self.assertRaises(ValueError): self.app.call_chat_api('OpenAI','','Tutor',[])
  self.session.post.assert_not_called()
 def test_no_cloud_fallback(self):
  with patch.object(C,'_call_openai_style_api',side_effect=RuntimeError('quota')) as call:
   with self.assertRaisesRegex(RuntimeError,'quota'): self.app.call_ai_api('OpenRouter','fake','Test','word','日常生活')
   self.assertEqual(call.call_count,1)
 def test_history_without_key(self):
  self.app.config={}; self.app.unsaved_api_keys={}
  self.app.provider_combo=types.SimpleNamespace(get=lambda:'Ollama')
  self.app.api_key_entry=types.SimpleNamespace(get=lambda:'')
  self.assertTrue(self.app.restore_practice_context({'context':{'provider':'Ollama'}}))
  self.assertEqual(self.app.practice_context['api_key'],'')
 def test_free_model(self): self.assertEqual(m['PROVIDER_MODELS']['OpenRouter'],'openrouter/free')
class UITests(unittest.TestCase):
 def test_groups_and_key_state(self):
  root=m['tk'].Tk(); root.withdraw()
  try:
   app=C(root)
   self.assertEqual(app.provider_combo.get(),'Ollama')
   self.assertEqual(str(app.api_key_entry['state']),'disabled')
   app.service_mode_var.set(list(m['SERVICE_GROUPS'])[1]); app.on_service_mode_change()
   self.assertEqual(tuple(app.provider_combo['values']),('OpenRouter',))
   self.assertEqual(str(app.api_key_entry['state']),'normal')
   app.service_mode_var.set(list(m['SERVICE_GROUPS'])[0]); app.on_service_mode_change()
   self.assertEqual(app.provider_combo.get(),'Ollama')
   self.assertEqual(app.api_key_entry.get(),'')
   app.open_settings_window(); root.update_idletasks()
   self.assertNotIn('Ollama',app.settings_provider_combo['values'])
   self.assertEqual(app.settings_provider_combo.get(),'OpenRouter')
  finally: root.destroy()
 def test_card_survives_missing_local_voice(self):
  root=m['tk'].Tk(); root.withdraw()
  try:
   app=C(root)
   folder=Path('work/ollama-card-check').resolve(); folder.mkdir(parents=True,exist_ok=True)
   with patch.object(app,'call_ai_api',return_value=card), patch.object(app,'generate_windows_audio',side_effect=RuntimeError('No English voice')), patch.object(m['messagebox'],'showinfo'), patch.object(m['messagebox'],'showerror') as err:
    app.run_process('Ollama','','Local Test','test',str(folder),'日常生活','en-US-JennyNeural',-25)
    root.update()
    err.assert_not_called()
    self.assertEqual(app.practice_context['provider'],'Ollama')
    self.assertEqual(app.practice_context['api_key'],'')
    self.assertTrue(Path(app.generated_docx_path).is_file())
    self.assertEqual(app.generated_audio_path,'')
  finally: root.destroy()
if __name__=='__main__': unittest.main(verbosity=2)
