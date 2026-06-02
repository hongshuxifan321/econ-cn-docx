# -*- coding: utf-8 -*-
"""econ-cn-docx 核心函数单元测试。运行: python tests.py"""
import unittest, sys, os, tempfile

# Ensure we can import from the skill root
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from references.toolkit import (
    fix_chinese_quotes, _latex_to_unicode, _parse_latex_tokens, _build_omath,
    STYLE_PRESETS, _get_counters, add_regression_table, add_three_line_table,
    add_display_equation, setup_page, add_numbered_heading, reset_heading_counters,
)
from docx import Document
from docx.shared import Pt


class TestFixChineseQuotes(unittest.TestCase):
    def test_simple_chinese(self):
        result = fix_chinese_quotes('他说"你好"世界')
        self.assertIn('“', result)  # left curly quote
        self.assertIn('”', result)  # right curly quote
        self.assertNotIn('"', result)   # no straight quote left

    def test_no_chinese(self):
        self.assertEqual(fix_chinese_quotes('He said "hello"'), 'He said "hello"')

    def test_mixed_but_contains_chinese(self):
        self.assertEqual(fix_chinese_quotes('本文基于"CHAP数据"得出'), '本文基于“CHAP数据”得出')

    def test_empty(self):
        self.assertEqual(fix_chinese_quotes(''), '')

    def test_multiple_pairs(self):
        self.assertEqual(
            fix_chinese_quotes('所谓"污染"的"经济损害"'),
            '所谓“污染”的“经济损害”')


class TestLatexToUnicode(unittest.TestCase):
    def test_greek_lower(self):
        self.assertIn('α', _latex_to_unicode('\\alpha'))
        self.assertIn('β', _latex_to_unicode('\\beta'))

    def test_greek_upper(self):
        self.assertIn('Δ', _latex_to_unicode('\\Delta'))

    def test_operators(self):
        self.assertIn('·', _latex_to_unicode('\\cdot'))
        self.assertIn('×', _latex_to_unicode('\\times'))

    def test_text_command(self):
        self.assertEqual(_latex_to_unicode('\\text{GDP}'), 'GDP')

    def test_subscript_keeps_underscore(self):
        # _latex_to_unicode doesn't remove underscores—that's handled by tokenizer
        s = _latex_to_unicode('x_{it}')
        self.assertIn('_', s)
        self.assertNotIn('\\beta', s.replace('β', ''))

    def test_function_names(self):
        self.assertIn('ln', _latex_to_unicode('\\ln'))
        self.assertIn('log', _latex_to_unicode('\\log'))


class TestParseLatexTokens(unittest.TestCase):
    def test_plain_text(self):
        tokens = _parse_latex_tokens('y = α + βx')
        self.assertTrue(any(t[0] == 'text' for t in tokens))

    def test_single_subscript(self):
        tokens = _parse_latex_tokens('x_{it}')
        self.assertTrue(any(t[0] == 'sub' for t in tokens))

    def test_single_superscript(self):
        tokens = _parse_latex_tokens('x^{2}')
        self.assertTrue(any(t[0] == 'sup' for t in tokens))

    def test_subsup_combo(self):
        tokens = _parse_latex_tokens('β_{1}^{2}')
        self.assertTrue(any(t[0] == 'subsup' for t in tokens))

    def test_complex_equation(self):
        tokens = _parse_latex_tokens('y_{it} = β_{1} x_{it} + ε_{it}')
        self.assertTrue(len(tokens) > 5)
        sub_count = sum(1 for t in tokens if t[0] == 'sub')
        self.assertGreaterEqual(sub_count, 3)


class TestOmathBuild(unittest.TestCase):
    def test_build_simple(self):
        tokens = _parse_latex_tokens('y = αx')
        omath = _build_omath(tokens)
        self.assertIsNotNone(omath)

    def test_build_with_subscript(self):
        tokens = _parse_latex_tokens('β_{1}')
        omath = _build_omath(tokens)
        self.assertIsNotNone(omath)


class TestStylePresets(unittest.TestCase):
    def test_three_presets_exist(self):
        for key in ['thesis', 'journal', 'working_paper']:
            self.assertIn(key, STYLE_PRESETS)

    def test_thesis_margins(self):
        s = STYLE_PRESETS['thesis']
        # thesis: left 3.18 cm (binding), right 2.54 cm
        self.assertGreater(s['margin_left'], s['margin_right'])

    def test_all_styles_have_required_keys(self):
        required = ['margin_top', 'margin_bottom', 'margin_left', 'margin_right',
                     'body_size', 'body_line_spacing', 'body_indent',
                     'heading1_size', 'heading2_size',
                     'ref_size', 'ref_line_spacing', 'table_size']
        for name, s in STYLE_PRESETS.items():
            for key in required:
                self.assertIn(key, s, f'{name} missing {key}')


class TestRegressionTable(unittest.TestCase):
    def setUp(self):
        self.doc = Document()
        setup_page(self.doc, style='thesis')

    def test_basic_regression_table(self):
        models = [
            {'name': '(1) FE', 'N': 5000, 'R2': 0.123,
             'coef': {'x': 0.123, 'z': -0.045},
             'se':   {'x': 0.023, 'z': 0.018},
             'pvalues': {'x': 0.000, 'z': 0.031}},
        ]
        # Should not raise
        add_regression_table(self.doc, models, notes='注：标准误在括号内。')

    def test_regression_table_star_annotation(self):
        models = [
            {'name': 'M1', 'N': 100, 'R2': 0.5,
             'coef': {'v': 0.05},
             'se':   {'v': 0.02},
             'pvalues': {'v': 0.02}},  # p<0.05 → **
        ]
        add_regression_table(self.doc, models)


class TestNumberedHeadings(unittest.TestCase):
    def test_counters_per_document(self):
        doc1 = Document()
        setup_page(doc1, style='thesis')
        doc2 = Document()
        setup_page(doc2, style='thesis')

        add_numbered_heading(doc1, '引言', 1)
        add_numbered_heading(doc1, '数据', 1)
        add_numbered_heading(doc2, '引言', 1)

        c1 = _get_counters(doc1)
        c2 = _get_counters(doc2)
        self.assertEqual(c1[1], 2)  # doc1 has two level-1 headings
        self.assertEqual(c2[1], 1)  # doc2 has one level-1 heading (isolated)


class TestDisplayEquation(unittest.TestCase):
    def test_simple_equation(self):
        doc = Document()
        setup_page(doc, style='thesis')
        add_display_equation(doc, r'y_{it} = \beta x_{it} + \varepsilon_{it}')
        # Should have created 1 paragraph
        self.assertGreaterEqual(len(doc.paragraphs), 1)


class TestThreeLineTable(unittest.TestCase):
    def test_basic_table(self):
        doc = Document()
        setup_page(doc, style='thesis')
        add_three_line_table(doc, ['变量', '(1)'],
            [['x', '0.123***'], ['z', '0.045']])
        # Should have created 1 table + trailing empty paragraph
        self.assertGreaterEqual(len(doc.tables), 1)


class TestFixChineseQuotesEdge(unittest.TestCase):
    """Additional edge cases for fix_chinese_quotes."""

    def test_unmatched_quote_pair_stays(self):
        # Unpaired quotes inside Chinese context should not be modified
        result = fix_chinese_quotes('他说"你好')
        # The regex looks for paired quotes: "...Chinese..."
        # An unmatched quote won't match the pair pattern
        self.assertIn('"', result)  # unmatched stays as ASCII

    def test_empty_quotes(self):
        self.assertEqual(fix_chinese_quotes('""'), '""')

    def test_multiline(self):
        text = '第一段说"污染"。\n第二段说"经济"。'
        result = fix_chinese_quotes(text)
        self.assertIn('“', result)
        self.assertIn('”', result)
        self.assertNotIn('"污染"', result)


if __name__ == '__main__':
    unittest.main(verbosity=2)
