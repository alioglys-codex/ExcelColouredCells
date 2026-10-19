from pathlib import Path
import tempfile
import unittest
import zipfile
from copy import copy
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import Rule, FormulaRule, CellIsRule, ColorScaleRule
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.differential import DifferentialStyle
from zero_colored_cells import modify_workbook, UnsupportedFormatting

FILL = PatternFill('solid', fgColor='FFFF0000')


class ConditionalTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.source = Path(self.temp.name) / 'input.xlsx'
        self.wb = Workbook()
        self.ws = self.wb.active

    def run_workbook(self):
        self.wb.save(self.source)
        return modify_workbook(self.source)

    def test_relative_mixed_absolute_and_ln_sqrt(self):
        self.ws.append([10, 100, 3])
        self.ws.append([20, 400, 5])
        self.ws['E1'] = 2
        self.ws.conditional_formatting.add('B1:C2', FormulaRule(
            formula=['AND(SQRT(B1)>$E$1,LN($A1)>1)'], fill=FILL))
        result = self.run_workbook()
        self.assertEqual(result['processed']['Sheet'], 3)
        out = load_workbook(result['output'])
        self.assertEqual([out.active[c].value for c in ('B1', 'C1', 'B2', 'C2')], [0, 3, 0, 0])
        out.close()

    def test_background_only_differential_fill(self):
        self.ws['A1'] = 42
        self.ws.conditional_formatting.add('A1', Rule(
            type='cellIs', operator='greaterThan', formula=['30'],
            dxf=DifferentialStyle(fill=PatternFill(bgColor='FFFFC000'))))
        self.assertEqual(self.run_workbook()['total_replaced'], 1)

    def test_font_only_rule_with_stop_prevents_lower_fill(self):
        self.ws['A1'] = 42
        self.ws.conditional_formatting.add('A1', FormulaRule(
            formula=['TRUE'], font=Font(bold=True), stopIfTrue=True))
        self.ws.conditional_formatting.add('A1', CellIsRule(
            operator='greaterThan', formula=['0'], fill=FILL))
        self.assertEqual(self.run_workbook()['total_replaced'], 0)

    def test_false_rules_errors_blanks_and_font_only(self):
        self.ws['A1'] = 10
        self.ws['A2'] = 40
        self.ws['B1'] = 'header'
        self.ws.conditional_formatting.add('A1:A3', FormulaRule(formula=['1/0>0'], fill=FILL))
        self.ws.conditional_formatting.add('A1:A3', CellIsRule(operator='greaterThan', formula=['30'], fill=FILL))
        self.ws.conditional_formatting.add('B1', FormulaRule(formula=['TRUE'], font=Font(bold=True)))
        self.assertEqual(self.run_workbook()['total_replaced'], 1)

    def test_cross_sheet_and_disjoint_relative_reference(self):
        params = self.wb.create_sheet('Input values')
        params['A1'] = 10
        self.ws['A1'], self.ws['C1'] = 12, 3
        self.ws['A2'], self.ws['C2'] = 4, 15
        self.ws.conditional_formatting.add('A1:A2 C1:C2', FormulaRule(
            formula=["A1>'Input values'!$A$1"], fill=FILL))
        self.assertEqual(self.run_workbook()['processed']['Sheet'], 2)

    def test_no_fill_override_and_priority(self):
        self.ws['A1'] = 4
        self.ws['A1'].fill = FILL
        self.ws.conditional_formatting.add('A1', Rule(
            type='expression', formula=['TRUE'], dxf=DifferentialStyle(fill=PatternFill())))
        self.ws.conditional_formatting.add('A1', FormulaRule(formula=['TRUE'], fill=FILL))
        self.assertEqual(self.run_workbook()['total_replaced'], 0)

    def test_color_scale(self):
        for i in range(1, 4):
            self.ws.cell(i, 1, i)
        self.ws['A4'] = 'header'
        self.ws.conditional_formatting.add('A1:A5', ColorScaleRule(
            start_type='min', start_color='FF0000', end_type='max', end_color='00FF00'))
        self.assertEqual(self.run_workbook()['total_replaced'], 3)

    def test_between_and_not_between(self):
        for i, v in enumerate([5, 15, 25], 1):
            self.ws.cell(i, 1, v)
            self.ws.cell(i, 2, v)
        self.ws.conditional_formatting.add('A1:A3', CellIsRule(operator='between', formula=['10','20'],fill=FILL))
        self.ws.conditional_formatting.add('B1:B3', CellIsRule(operator='notBetween', formula=['10','20'],fill=FILL))
        self.assertEqual(self.run_workbook()['total_replaced'], 3)

    def test_missing_formula_cache_has_actionable_error(self):
        self.ws['A1'] = '=10+20'
        self.ws.conditional_formatting.add('A1', FormulaRule(formula=['A1>0'], fill=FILL))
        self.wb.save(self.source)
        with self.assertRaisesRegex(ValueError, 'recalculate, save, close'):
            modify_workbook(self.source)
        self.assertFalse(self.source.with_stem('input_zeroed').exists())

    def test_unsupported_rule_is_not_silently_ignored(self):
        self.ws['A1'] = 1
        self.ws.conditional_formatting.add('A1', FormulaRule(formula=['VLOOKUP(A1,B1:C2,2)>0'],fill=FILL))
        self.wb.save(self.source)
        with self.assertRaisesRegex(UnsupportedFormatting, 'VLOOKUP'):
            modify_workbook(self.source)

    def test_shared_master_replacement_preserves_uncolored_formulas(self):
        self.ws['A1'], self.ws['A2'], self.ws['A3'] = '=B1*2', '=B2*2', '=B3*2'
        self.ws['A1'].fill = FILL
        self.wb.create_sheet('Check')['A1'] = 'unchanged'
        self.wb.save(self.source)
        with zipfile.ZipFile(self.source) as original:
            files = [(copy(i), original.read(i.filename)) for i in original.infolist()]
        with zipfile.ZipFile(self.source, 'w') as archive:
            for info, data in files:
                if info.filename == 'xl/worksheets/sheet1.xml':
                    data = data.replace(b'<f>B1*2</f>', b'<f t="shared" ref="A1:A3" si="0">B1*2</f>')
                    data = data.replace(b'<f>B2*2</f>', b'<f t="shared" si="0"/>')
                    data = data.replace(b'<f>B3*2</f>', b'<f t="shared" si="0"/>')
                archive.writestr(info, data)
        result = modify_workbook(self.source)
        out = load_workbook(result['output'])
        self.assertEqual([out.active[c].value for c in ('A1','A2','A3')], [0, '=B2*2', '=B3*2'])
        out.close()
        with zipfile.ZipFile(self.source) as original, zipfile.ZipFile(result['output']) as edited:
            for path in original.namelist():
                if path != 'xl/worksheets/sheet1.xml':
                    self.assertEqual(original.read(path), edited.read(path), path)

    def test_existing_destination_is_preserved(self):
        self.wb.save(self.source)
        output = self.source.with_stem('input_zeroed')
        output.write_bytes(b'existing file')
        with self.assertRaisesRegex(ValueError, 'never overwritten'):
            modify_workbook(self.source)
        self.assertEqual(output.read_bytes(), b'existing file')

    def test_uncached_cross_sheet_reference_is_not_assumed_blank(self):
        inputs = self.wb.create_sheet('Check')
        inputs['A1'] = '=10+20'
        self.ws['A1'] = 50
        self.ws.conditional_formatting.add('A1', FormulaRule(
            formula=['A1>Check!$A$1'], fill=FILL))
        self.wb.save(self.source)
        with self.assertRaisesRegex(ValueError, 'Check!A1: missing saved formula result'):
            modify_workbook(self.source)

    def test_array_overlap_rejected_and_untouched_array_preserved(self):
        from openpyxl.worksheet.formula import ArrayFormula
        self.ws['A1'] = ArrayFormula(ref='A1:A2', text='=B1:B2*2')
        self.ws['B1'], self.ws['B2'] = 10, 20
        self.ws['B1'].fill = FILL
        result = self.run_workbook()
        out = load_workbook(result['output'])
        self.assertEqual(out.active['A1'].value.text, '=B1:B2*2')
        out.close()
        self.ws['A2'] = 40
        self.ws['A2'].fill = FILL
        self.wb.save(self.source)
        with self.assertRaisesRegex(ValueError, 'overlap array'):
            modify_workbook(self.source, Path(self.temp.name) / 'array_output.xlsx')
