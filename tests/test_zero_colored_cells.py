import tempfile
import unittest
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, GradientFill, Font, Border, Side
from zero_colored_cells import modify_workbook


class WorkbookTests(unittest.TestCase):
    def test_replacements_and_preservation(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'input.xlsx'
            wb = Workbook()
            ws = wb.active
            ws.title = 'Data'
            for index, color in enumerate(['FFA500', 'A52A2A', 'FF0000', '800080',
                                            '0000FF', 'FFFF00', '000000', 'FFFFFF'], 1):
                ws.cell(index, 1, '=1+2' if index == 1 else 'value')
                ws.cell(index, 1).fill = PatternFill('solid', fgColor=color)
            ws['B1'] = '=SUM(1,2)'
            ws['B1'].font = Font(bold=True)
            ws['B1'].border = Border(left=Side(style='thin'))
            ws['C1'].fill = PatternFill('solid', fgColor='FFFF00')
            ws['D1'] = 'pattern'
            ws['D1'].fill = PatternFill('darkGrid', fgColor='FF0000')
            ws['E1'] = 'gradient'
            ws['E1'].fill = GradientFill(stop=['FF0000', '0000FF'])
            ws['F1'] = 'merged'
            ws['F1'].fill = PatternFill('solid', fgColor='0000FF')
            ws.merge_cells('F1:G2')
            ws.row_dimensions[1].height = 30
            ws.column_dimensions['A'].width = 25
            for name in ['Check', 'Sheet3']:
                excluded = wb.create_sheet(name)
                excluded['A1'] = '=2+2'
                excluded['A1'].fill = PatternFill('solid', fgColor='FF0000')
            wb.save(source)
            original = source.read_bytes()
            result = modify_workbook(source)
            self.assertEqual(result['processed'], {'Data': 11})
            self.assertEqual(result['ignored'], ['Check', 'Sheet3'])
            self.assertEqual(source.read_bytes(), original)
            output = load_workbook(result['output'])
            self.assertEqual(output['Data']['A1'].value, 0)
            self.assertEqual(output['Data']['B1'].value, '=SUM(1,2)')
            self.assertIsNone(output['Data']['C1'].value)
            self.assertEqual(output['Check']['A1'].value, '=2+2')
            output.close()
            with self.assertRaises(ValueError):
                modify_workbook(source)

    def test_missing_exclusions_and_no_fill(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'input.xlsx'
            wb = Workbook()
            wb.active['A1'] = 42
            wb.active['A1'].fill = PatternFill(fgColor='FF0000')
            wb.save(source)
            result = modify_workbook(source)
            self.assertEqual(result['total_replaced'], 0)
            self.assertEqual(result['excluded_not_present'], ['Check', 'Sheet3'])

    def test_conditional_formatting_supported(self):
        from openpyxl.formatting.rule import CellIsRule
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'input.xlsx'
            wb = Workbook()
            wb.active['A1'] = 42
            wb.active.conditional_formatting.add('A1', CellIsRule(
                operator='greaterThan', formula=['0'],
                fill=PatternFill('solid', fgColor='FF0000')))
            wb.save(source)
            result = modify_workbook(source)
            self.assertEqual(result['total_replaced'], 1)
            out = load_workbook(result['output'])
            self.assertEqual(out.active['A1'].value, 0)
            self.assertEqual(len(out.active.conditional_formatting), 1)
            out.close()
