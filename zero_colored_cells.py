"""Zero nonempty Excel cells with direct or supported conditional background fills."""
import argparse
import ast
from copy import copy
import json
import math
import operator
from pathlib import Path
import posixpath
import re
import warnings
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
import zipfile

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.formula import Tokenizer
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.cell_range import CellRange

EXCLUDED = frozenset({'Check', 'Sheet3'})
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
M = '{' + NS + '}'
CELL_XML = re.compile(rb'<(?:\w+:)?c\b[^>]*?(?:/>|>.*?</(?:\w+:)?c>)', re.S)
FORMULA_XML = re.compile(rb'<(?:\w+:)?f\b[^>]*?(?:/>|>.*?</(?:\w+:)?f>)', re.S)


class UnsupportedFormatting(ValueError):
    pass


class ExcelError(Exception):
    """An Excel calculation error means a conditional rule does not apply."""


def flatten(values):
    for value in values:
        if isinstance(value, list):
            yield from flatten(value)
        else:
            yield value


class Formula:
    """Parse a limited Excel formula and interpret it without executing Python code."""
    FUNCTIONS = {'LN', 'LOG10', 'SQRT', 'ABS', 'EXP', 'SUM', 'MIN', 'MAX',
                 'AND', 'OR', 'NOT', 'IF', 'ISNUMBER', 'ISBLANK', 'IFERROR'}

    def __init__(self, text):
        parts = []
        for token in Tokenizer('=' + text.lstrip('=')).items:
            kind, subtype, value = token.type, token.subtype, token.value
            if kind == 'WHITE-SPACE':
                continue
            if kind == 'OPERAND':
                if subtype == 'RANGE':
                    # Defined names, external workbooks and structured references
                    # must not silently turn into empty cell references.
                    reference = value.rsplit('!', 1)[-1]
                    if '[' in value or not re.fullmatch(
                            r'\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?', reference):
                        raise UnsupportedFormatting(f'Unsupported formula reference: {value}')
                    parts.append('REF(' + repr(value) + ')')
                elif subtype == 'TEXT':
                    parts.append(repr(value[1:-1].replace('""', '"')))
                elif subtype == 'LOGICAL':
                    parts.append(str(value.upper() == 'TRUE'))
                elif subtype == 'NUMBER':
                    parts.append(value)
                elif subtype == 'ERROR':
                    parts.append('ERROR()')
                else:
                    raise UnsupportedFormatting(f'Unsupported formula token: {value}')
            elif kind == 'FUNC':
                if subtype == 'OPEN':
                    function = value[:-1].upper()
                    if function not in self.FUNCTIONS:
                        raise UnsupportedFormatting(f'Unsupported Excel function: {function}')
                    parts.append(function + '(')
                else:
                    parts.append(')')
            elif kind == 'PAREN':
                parts.append(value)
            elif kind == 'SEP' and subtype == 'ARG':
                parts.append(',')
            elif kind in ('OPERATOR-INFIX', 'OPERATOR-PREFIX'):
                translated = {'=': '==', '<>': '!=', '^': '**'}.get(value, value)
                if translated not in ('+', '-', '*', '/', '**', '==', '!=', '<', '>', '<=', '>='):
                    raise UnsupportedFormatting(f'Unsupported formula operator: {value}')
                parts.append(translated)
            else:
                raise UnsupportedFormatting(f'Unsupported formula syntax: {value}')
        try:
            self.tree = ast.parse(''.join(parts), mode='eval').body
        except SyntaxError as error:
            raise UnsupportedFormatting(f'Cannot parse Excel formula: {text}') from error

    def value(self, cached, sheet_name, origin, target, missing=frozenset()):
        row_delta = target.row - origin[0]
        col_delta = target.column - origin[1]

        def reference(value):
            try:
                value = Translator.translate_range(value, row_delta, col_delta)
            except Exception as error:
                raise ExcelError('Invalid shifted reference') from error
            if '!' in value:
                name, value = value.rsplit('!', 1)
                name = name.strip("'").replace("''", "'")
            else:
                name = sheet_name
            if name not in cached.sheetnames:
                raise UnsupportedFormatting(f'Missing referenced worksheet: {name}')
            sheet = cached[name]
            def read(address):
                address = address.replace('$', '')
                if (name, address) in missing:
                    raise UnsupportedFormatting(
                        f'{name}!{address}: missing saved formula result. Open the original workbook '
                        'in Excel, recalculate, save, close, and retry.')
                cell = sheet[address]
                if cell.data_type == 'e':
                    raise ExcelError(cell.value)
                return cell.value
            if ':' in value:
                left, top, right, bottom = range_boundaries(value)
                if (right - left + 1) * (bottom - top + 1) > 100000:
                    raise UnsupportedFormatting('Formula reference range is too large.')
                return [read(cell.coordinate) for row in sheet.iter_rows(
                    min_row=top, max_row=bottom, min_col=left, max_col=right) for cell in row]
            return read(value)

        def number(value):
            if value is None:
                return 0
            if isinstance(value, (int, float)):
                return value
            raise ExcelError('A numeric value was required')

        def visit(node):
            if isinstance(node, ast.Constant):
                return node.value
            if isinstance(node, ast.UnaryOp):
                value = number(visit(node.operand))
                if isinstance(node.op, ast.USub):
                    return -value
                if isinstance(node.op, ast.UAdd):
                    return value
            if isinstance(node, ast.BinOp):
                functions = {ast.Add: operator.add, ast.Sub: operator.sub,
                             ast.Mult: operator.mul, ast.Div: operator.truediv, ast.Pow: operator.pow}
                if type(node.op) in functions:
                    return functions[type(node.op)](number(visit(node.left)), number(visit(node.right)))
            if isinstance(node, ast.Compare) and len(node.ops) == 1:
                left, right = visit(node.left), visit(node.comparators[0])
                if left is None:
                    left = '' if isinstance(right, str) else 0
                if right is None:
                    right = '' if isinstance(left, str) else 0
                if isinstance(left, str) and isinstance(right, str):
                    left, right = left.casefold(), right.casefold()
                elif isinstance(left, str) or isinstance(right, str):
                    raise UnsupportedFormatting('Mixed text/numeric comparison needs Excel evaluation.')
                comparisons = {ast.Eq: operator.eq, ast.NotEq: operator.ne,
                               ast.Lt: operator.lt, ast.Gt: operator.gt,
                               ast.LtE: operator.le, ast.GtE: operator.ge}
                return comparisons[type(node.ops[0])](left, right)
            if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and not node.keywords:
                name = node.func.id
                if name == 'REF':
                    return reference(visit(node.args[0]))
                if name == 'ERROR':
                    raise ExcelError('Excel error literal')
                if name == 'IF':
                    return visit(node.args[1] if visit(node.args[0]) else node.args[2])
                if name == 'IFERROR':
                    try:
                        return visit(node.args[0])
                    except (ExcelError, ArithmeticError, ValueError) as error:
                        if isinstance(error, UnsupportedFormatting):
                            raise
                        return visit(node.args[1])
                args = [visit(arg) for arg in node.args]
                if name in ('SUM', 'MIN', 'MAX'):
                    values = [v for v in flatten(args) if isinstance(v, (int, float))]
                    return {'SUM': sum, 'MIN': lambda v: min(v, default=0),
                            'MAX': lambda v: max(v, default=0)}[name](values)
                if name in ('AND', 'OR'):
                    return (all if name == 'AND' else any)(flatten(args))
                if name == 'NOT':
                    return not args[0]
                if name == 'ISNUMBER':
                    return isinstance(args[0], (int, float)) and not isinstance(args[0], bool)
                if name == 'ISBLANK':
                    return args[0] is None
                functions = {'LN': math.log, 'LOG10': math.log10, 'SQRT': math.sqrt,
                             'ABS': abs, 'EXP': math.exp}
                if name in functions:
                    return functions[name](*[number(arg) for arg in args])
            raise UnsupportedFormatting('Unsupported formula expression.')
        return visit(self.tree)


def direct_fill(fill):
    return fill.fill_type is not None


def differential_fill(fill):
    # Excel differential styles commonly specify only bgColor with no patternType.
    if fill.fill_type:
        return True
    return any(color.type != 'rgb' or color.rgb != '00000000'
               for color in (fill.fgColor, fill.bgColor))


def conditional_rules(sheet):
    rules = []
    for group in sheet.conditional_formatting:
        first = str(group.sqref).split()[0]
        left, top, _, _ = range_boundaries(first)
        for rule in sheet.conditional_formatting[group]:
            fill = rule.dxf.fill if rule.dxf else None
            if fill is None and rule.type != 'colorScale' and not rule.stopIfTrue:
                continue
            if rule.type not in ('expression', 'cellIs', 'colorScale'):
                raise UnsupportedFormatting(f'{sheet.title}: unsupported conditional rule {rule.type}.')
            if rule.type == 'cellIs' and rule.operator not in (
                    'lessThan', 'lessThanOrEqual', 'greaterThan', 'greaterThanOrEqual',
                    'equal', 'notEqual', 'between', 'notBetween'):
                raise UnsupportedFormatting(f'Unsupported comparison: {rule.operator}')
            formulas = [Formula(formula) for formula in rule.formula]
            required = 1 if rule.type == 'expression' else (2 if rule.operator in ('between', 'notBetween') else 1)
            if rule.type != 'colorScale' and len(formulas) != required:
                raise UnsupportedFormatting(f'{sheet.title}: invalid conditional formula count.')
            rules.append((rule.priority, group.sqref, (top, left), rule, fill, formulas))
    return sorted(rules, key=lambda item: item[0])


def is_colored(cell, rules, cached, sheet_name, missing=frozenset()):
    result = direct_fill(cell.fill)
    fill_decided = False
    for _, ranges, origin, rule, fill, formulas in rules:
        if cell.coordinate not in ranges:
            continue
        try:
            if rule.type in ('cellIs', 'colorScale') and (sheet_name, cell.coordinate) in missing:
                raise UnsupportedFormatting(
                    f'{sheet_name}!{cell.coordinate}: missing saved formula result. Open the original '
                    'workbook in Excel, recalculate, save, close, and retry.')
            if rule.type == 'colorScale':
                value = cached[sheet_name][cell.coordinate].value
                applies = isinstance(value, (int, float)) and not isinstance(value, bool)
            elif rule.type == 'expression':
                applies = bool(formulas[0].value(cached, sheet_name, origin, cell, missing))
            else:
                value_cell = cached[sheet_name][cell.coordinate]
                if value_cell.data_type == 'e':
                    continue
                value = value_cell.value
                operands = [f.value(cached, sheet_name, origin, cell, missing) for f in formulas]
                if value is None:
                    value = 0
                if any(isinstance(v, str) != isinstance(value, str) for v in operands):
                    raise UnsupportedFormatting('Mixed text/numeric conditional comparison.')
                if isinstance(value, str):
                    value, operands = value.casefold(), [v.casefold() for v in operands]
                comparisons = {'lessThan': operator.lt, 'lessThanOrEqual': operator.le,
                               'greaterThan': operator.gt, 'greaterThanOrEqual': operator.ge,
                               'equal': operator.eq, 'notEqual': operator.ne}
                if rule.operator in ('between', 'notBetween'):
                    applies = min(operands) <= value <= max(operands)
                    if rule.operator == 'notBetween':
                        applies = not applies
                else:
                    applies = comparisons[rule.operator](value, operands[0])
        except UnsupportedFormatting:
            raise
        except (ExcelError, ArithmeticError, ValueError, TypeError):
            # Conditional formulas yielding #DIV/0!, #VALUE!, etc. are not applied.
            applies = False
        if applies:
            if not fill_decided and (fill is not None or rule.type == 'colorScale'):
                result = rule.type == 'colorScale' or differential_fill(fill)
                fill_decided = True
            if rule.stopIfTrue:
                break
    return result


def worksheet_paths(archive):
    relationships = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
    targets = {r.attrib['Id']: r.attrib['Target'] for r in relationships}
    root = ET.fromstring(archive.read('xl/workbook.xml'))
    result = {}
    for sheet in root.find(M + 'sheets'):
        key = sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
        path = targets[key]
        result[sheet.attrib['name']] = (path.lstrip('/') if path.startswith('/')
                                        else posixpath.normpath('xl/' + path))
    return result


def patch_sheet(data, sheet, selected):
    root = ET.fromstring(data)
    affected = set()
    for cell in root.findall('.//' + M + 'sheetData/' + M + 'row/' + M + 'c'):
        formula = cell.find(M + 'f')
        if (cell.attrib['r'] in selected and formula is not None
                and formula.attrib.get('t') == 'shared' and formula.text):
            affected.add(formula.attrib['si'])
    seen = set()
    def patch(match):
        raw = match.group()
        address = re.search(rb'\br="([^"]+)"', raw).group(1).decode()
        if address in selected:
            seen.add(address)
            tag = re.match(rb'<([^\s/>]+)', raw).group(1)
            prefix = tag[:-1]  # c -> empty; x:c -> x:
            if raw.endswith(b'/>'):
                raise ValueError('Unexpected nonempty cell with no XML value.')
            start, body = raw.split(b'>', 1)
            start = re.sub(rb'\s+t="[^"]*"', b'', start)
            payload = re.compile(rb'<(?:\w+:)?(?:f|v|is)\b[^>]*?(?:/>|>.*?</(?:\w+:)?(?:f|v|is)>)', re.S)
            body = payload.sub(b'', body)
            return start + b'>' + b'<' + prefix + b'v>0</' + prefix + b'v>' + body
        if affected:
            def expand(formula_match):
                fragment = formula_match.group()
                shared = re.search(rb'\bsi="([^"]+)"', fragment)
                if shared and shared.group(1).decode() in affected:
                    tag = re.match(rb'<([^\s/>]+)', fragment).group(1)
                    formula = sheet[address].value
                    if not isinstance(formula, str) or not formula.startswith('='):
                        raise ValueError(f'Cannot expand shared formula: {sheet.title}!{address}')
                    return b'<' + tag + b'>' + escape(formula[1:]).encode() + b'</' + tag + b'>'
                return fragment
            return FORMULA_XML.sub(expand, raw)
        return raw
    result = CELL_XML.sub(patch, data)
    if seen != selected:
        raise ValueError('Could not locate every target in worksheet XML.')
    return result


def normalized_xml(data, selected, sheet):
    root = ET.fromstring(data)
    for cell in root.findall('.//' + M + 'sheetData/' + M + 'row/' + M + 'c'):
        address = cell.attrib['r']
        if address in selected:
            cell.attrib.pop('t', None)
            for child in list(cell):
                if child.tag in (M + 'f', M + 'v', M + 'is'):
                    cell.remove(child)
        else:
            formula = cell.find(M + 'f')
            # Shared formulas may be expanded to preserve untargeted formulas.
            if formula is not None and formula.attrib.get('t') not in ('array', 'dataTable'):
                formula.attrib.clear()
                formula.text = sheet[address].value[1:]
    return ET.tostring(root)


def modify_workbook(source, destination=None):
    source = Path(source).resolve()
    destination = (Path(destination).resolve() if destination else
                   source.with_name(source.stem + '_zeroed.xlsx'))
    if source.suffix.lower() != '.xlsx' or destination.suffix.lower() != '.xlsx':
        raise ValueError('Input and output must be .xlsx files.')
    if source == destination or destination.exists():
        raise ValueError('Output must be a new file; existing files are never overwritten.')
    created = False
    with warnings.catch_warnings():
        warnings.simplefilter('error')
        workbook = load_workbook(source, data_only=False, rich_text=True)
        cached = load_workbook(source, data_only=True)
        selected = {}
        missing = {(sheet.title, cell.coordinate) for sheet in workbook for row in sheet for cell in row
                   if cell.data_type == 'f' and cached[sheet.title][cell.coordinate].value is None}
        try:
            for sheet in workbook:
                if sheet.title in EXCLUDED:
                    continue
                if sheet.tables:
                    raise UnsupportedFormatting(f'{sheet.title}: table styles are not supported.')
                rules = conditional_rules(sheet)
                coordinates = set()
                for row in sheet:
                    for cell in row:
                        if isinstance(cell, MergedCell) or cell.value is None or cell.value == '':
                            continue
                        if is_colored(cell, rules, cached, sheet.title, missing):
                            coordinates.add(cell.coordinate)
                # Replacing one array-formula cell could change untargeted spill cells.
                for row in sheet:
                    for cell in row:
                        if cell.data_type == 'f' and not isinstance(cell.value, str):
                            region = getattr(cell.value, 'ref', cell.coordinate)
                            array_range = CellRange(region)
                            if any(coordinate in array_range for coordinate in coordinates):
                                raise UnsupportedFormatting(
                                    f'{sheet.title}: targets overlap array/data-table formula {region}.')
                selected[sheet.title] = coordinates

            with zipfile.ZipFile(source) as original:
                paths = worksheet_paths(original)
                changed_paths = {paths[name]: name for name, cells in selected.items() if cells}
                try:
                    with destination.open('xb') as output:
                        created = True
                        with zipfile.ZipFile(output, 'w', compression=zipfile.ZIP_DEFLATED) as edited:
                            for info in original.infolist():
                                data = original.read(info.filename)
                                name = changed_paths.get(info.filename)
                                if name is not None:
                                    data = patch_sheet(data, workbook[name], selected[name])
                                edited.writestr(copy(info), data)
                            edited.comment = original.comment
                    with zipfile.ZipFile(destination) as edited:
                        if original.namelist() != edited.namelist():
                            raise ValueError('Workbook package structure changed.')
                        for path in original.namelist():
                            before, after = original.read(path), edited.read(path)
                            if path not in changed_paths:
                                if before != after:
                                    raise ValueError(f'Untargeted workbook part changed: {path}')
                            else:
                                name = changed_paths[path]
                                if normalized_xml(before, selected[name], workbook[name]) != normalized_xml(
                                        after, selected[name], workbook[name]):
                                    raise ValueError(f'Untargeted worksheet content changed: {name}')
                    verified = load_workbook(destination, data_only=False, rich_text=True)
                    try:
                        if workbook.sheetnames != verified.sheetnames:
                            raise ValueError('Sheet names or order changed.')
                        for sheet in workbook:
                            for row in sheet:
                                for cell in row:
                                    if isinstance(cell, MergedCell):
                                        continue
                                    actual = verified[sheet.title][cell.coordinate]
                                    if cell.coordinate in selected.get(sheet.title, set()):
                                        if actual.value != 0 or actual.data_type != 'n':
                                            raise ValueError(f'Target verification failed: {sheet.title}!{cell.coordinate}')
                                    elif (actual.data_type != cell.data_type or
                                          (vars(actual.value) != vars(cell.value)
                                           if cell.data_type == 'f' and not isinstance(cell.value, str)
                                           else actual.value != cell.value)):
                                        raise ValueError(f'Untargeted value/formula changed: {sheet.title}!{cell.coordinate}')
                    finally:
                        verified.close()
                except BaseException:
                    if created:
                        destination.unlink(missing_ok=True)
                    raise
        finally:
            cached.close()
            workbook.close()
    return {
        'output': str(destination),
        'processed': {name: len(cells) for name, cells in selected.items()},
        'ignored': [name for name in workbook.sheetnames if name in EXCLUDED],
        'excluded_not_present': sorted(EXCLUDED - set(workbook.sheetnames)),
        'total_replaced': sum(map(len, selected.values())),
        'verification': 'passed',
        'note': 'Conditional colors use original saved formula results. Formatting rules are preserved; '
                'displayed colors may change when Excel recalculates after values become zero.',
    }


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input', type=Path)
    parser.add_argument('--output', type=Path)
    args = parser.parse_args()
    try:
        print(json.dumps(modify_workbook(args.input, args.output), indent=2))
    except (ValueError, OSError, zipfile.BadZipFile) as error:
        parser.exit(1, f'Error: {error}\n')
