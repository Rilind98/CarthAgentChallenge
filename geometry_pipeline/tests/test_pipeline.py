"""Offline regressions for the OpenAI switch; run with unittest discover."""
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace as NS
from unittest.mock import Mock, patch

import cv2
import numpy as np

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'agent'))
import layout
import ocr
import openai_backend as backend
import run


class ScriptedConversation:
    steps = []
    instances = []

    def __init__(self):
        self.steps = iter(type(self).steps)
        self.usage = {'in': 0, 'out': 0, 'cost_usd': 0.0, 'calls': 0}
        self.inputs, self.results = [], []
        type(self).instances.append(self)

    def user(self, text, images=()):
        self.inputs.append((text, list(images)))

    def request(self, **kwargs):
        return next(self.steps)

    def result(self, call_id, output):
        self.results.append((call_id, output))


class OpenAIProtocolTests(unittest.TestCase):
    def test_responses_keeps_reasoning_and_pairs_function_results(self):
        reasoning = Mock()
        reasoning.type = 'reasoning'
        reasoning.model_dump.return_value = {'type': 'reasoning', 'id': 'r1', 'encrypted_content': 'opaque'}
        call = Mock()
        call.type, call.call_id, call.name, call.arguments = 'function_call', 'c1', 'view', '{}'
        call.model_dump.return_value = {'type': 'function_call', 'call_id': 'c1', 'name': 'view', 'arguments': '{}'}
        response = NS(output=[reasoning, call], output_text='', usage=NS(input_tokens=10, output_tokens=5))
        client = NS(base_url='https://api.openai.com/v1/', responses=NS(create=Mock(return_value=response)))
        with patch.object(backend, 'OpenAI', return_value=client), patch.dict(os.environ, {}, clear=True):
            session = backend.Conversation()
            session.user('Read this', [b'png'])
            _, calls = session.request(tools=backend.TOOLS)
            self.assertEqual(calls, [('c1', 'view', '{}')])
            session.result('c1', 'image: view_0.png')
        kwargs = client.responses.create.call_args.kwargs
        self.assertFalse(kwargs['store'])
        self.assertIn('reasoning.encrypted_content', kwargs['include'])
        self.assertEqual(kwargs['tools'][0]['name'], 'view')
        self.assertEqual(session.messages[0]['content'][1]['type'], 'input_image')
        self.assertEqual(session.messages[1]['encrypted_content'], 'opaque')
        self.assertEqual(session.messages[-1]['call_id'], 'c1')
        self.assertEqual(session.usage['in'], 10)

    def test_openrouter_uses_chat_tool_protocol_and_model_prefix(self):
        message = NS(content=None, tool_calls=[NS(id='c1', function=NS(name='check', arguments='{}'))],
                     model_dump=lambda **kw: {'role': 'assistant', 'tool_calls': []})
        create = Mock(return_value=NS(choices=[NS(message=message)], usage=NS(prompt_tokens=4, completion_tokens=2)))
        client = NS(base_url='https://openrouter.ai/api/v1/', chat=NS(completions=NS(create=create)))
        with patch.object(backend, 'OpenAI', return_value=client), patch.dict(os.environ, {}, clear=True):
            session = backend.Conversation()
            session.user('Read this', [b'png'])
            _, calls = session.request(tools=backend.TOOLS)
            session.result(calls[0][0], 'Every row adds up.')
        self.assertEqual(create.call_args.kwargs['model'], 'openai/gpt-6.1-sol')
        self.assertEqual(session.messages[0]['content'][1]['type'], 'image_url')
        self.assertEqual(session.messages[-1]['tool_call_id'], 'c1')


class PipelineTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.path = Path(self.tmp.name)
        ScriptedConversation.instances = []

    def test_text_pdf_and_excel_extract_expected_values_and_pointers(self):
        corpus = ROOT.parent / 'A3-Hackathon-Corpus'
        for name in ('f00_demo_02.pdf', 'f00_demo_01.xlsx'):
            with self.subTest(name=name):
                doc = next(corpus.rglob(name))
                old_layout = json.loads((ROOT / 'out_final' / (doc.stem + '.layout.json')).read_text())
                ScriptedConversation.steps = [(json.dumps(old_layout), [])]
                with patch.object(backend, 'Conversation', ScriptedConversation):
                    result = run.run_doc(doc, self.path)
                expected = json.loads((ROOT / 'out_final' / (doc.stem + '_2025.json')).read_text())
                actual = json.loads((self.path / (doc.stem + '_2025.json')).read_text())
                self.assertEqual(result['values'], len(expected))
                self.assertEqual(actual, expected)

    def scan(self):
        png = cv2.imencode('.png', np.full((100, 200), 255, dtype=np.uint8))[1].tobytes()
        return {'png': png, 'size': (200, 100), 'k': 0, 'w': 200, 'h': 100,
                'words': [{'id': 0, 'box': [10, 10, 80, 30], 'text': '1,00'}],
                'cache': self.path / 'cache.json', 'record': None}

    def page_layout(self):
        return {'page': 1, 'teur': False, 'strips': [(0, 48)], 'rules': [], 'kinds': [], 'headers': {}}

    def test_real_view_and_fix_tools_feed_review_cache_and_second_pass(self):
        scan = self.scan()
        ScriptedConversation.steps = [('', [('view1', 'view', '{}')]),
                                      ('', [('fix1', 'fix', '{"id":0,"text":"2,00"}')]), ('done', [])]
        with patch.object(backend, 'Conversation', ScriptedConversation), patch.dict(os.environ, {'OCR_AGENT': '1'}):
            words, _ = ocr.review_page(scan, self.page_layout())
        self.assertEqual(words[0].text, '2,00')
        self.assertEqual(json.loads(scan['cache'].read_text())['edits'][0]['old'], '1,00')
        self.assertEqual(len(ScriptedConversation.instances[0].inputs[1][1]), 1)
        ScriptedConversation.steps = [('', [('fix2', 'fix', '{"id":0,"text":"3,00"}')]), ('rechecked', [])]
        items = [{'box': (0, 0, 30, 12), 'text': '2,00', 'row': 'asset', 'field': 'ahk_close', 'note': 'sum differs'}]
        with patch.object(backend, 'Conversation', ScriptedConversation), patch.dict(os.environ, {'OCR_AGENT': '1'}):
            words, _ = ocr.recheck_page(scan, self.page_layout(), items)
        self.assertEqual(words[0].text, '3,00')
        record = json.loads(scan['cache'].read_text())
        self.assertEqual(record['recheck']['answer'], 'rechecked')
        self.assertEqual(record['words'][0]['text'], '3,00')
        # A cached second review must not start another model session.
        with patch.object(backend, 'Conversation', side_effect=AssertionError('cache missed')):
            words, _ = ocr.recheck_page(scan, self.page_layout(), items)
        self.assertEqual(words[0].text, '3,00')

    def test_exhausted_ocr_budget_does_not_cache_incomplete_review(self):
        scan = self.scan()
        ScriptedConversation.steps = [('', [('view1', 'view', '{}')])]
        with patch.object(backend, 'Conversation', ScriptedConversation), \
             patch.dict(os.environ, {'OCR_AGENT': '1', 'MAX_CALLS': '1'}):
            with self.assertRaisesRegex(RuntimeError, 'MAX_CALLS=1'):
                ocr.review_page(scan, self.page_layout())
        self.assertFalse(scan['cache'].exists())
        self.assertIsNone(scan['record'])

    def test_invalid_layout_json_is_retried(self):
        good = {'unit': 'EUR', 'pages': [], 'groups': [], 'kinds': []}
        ScriptedConversation.steps = [('not json', []), (json.dumps(good), [])]
        with patch.object(backend, 'Conversation', ScriptedConversation):
            result = backend.decide([], [], [])
        self.assertEqual(result['_problems'], [])
        self.assertIn('not a valid JSON', ScriptedConversation.instances[0].inputs[1][0])


if __name__ == '__main__':
    unittest.main()
