from io import StringIO
from openpyxl import Workbook
from xlsxq import __version__
import json


def test_version():
    assert __version__ == '0.1.1'


def test_sheet_list_query_prints_sheet_list(tmpdir):
    book_path = tmpdir.join("book.xlsx")
    wb = Workbook()
    ws = wb.active
    assert ws.title == 'Sheet'
    ws['A1'] = 1
    wb.save(str(book_path))
    from xlsxq import SheetListQuery
    query = SheetListQuery(infile=str(book_path), output='json')
    io = StringIO()
    query.execute(file=io)
    data = json.loads(io.getvalue())
    assert data == [{'name': 'Sheet'}]


def test_range_show_query_prints_values(tmpdir):
    book_path = tmpdir.join("book.xlsx")
    wb = Workbook()
    ws = wb.create_sheet("Mysheet")
    ws['A1'] = 11
    ws['B1'] = 12
    ws['A2'] = 21
    ws['B2'] = 22
    ws['A3'] = 31
    ws['B3'] = 32
    wb.save(str(book_path))
    from xlsxq import RangeShowQuery
    query = RangeShowQuery(
            infile=str(book_path), sheet='Mysheet', range_='A1:B3',
            output='tsv')
    io = StringIO()
    query.execute(file=io)
    text = io.getvalue()
    lines = text.splitlines()
    assert len(lines) == 3
    assert lines[0] == "11\t12"
    assert lines[1] == "21\t22"
    assert lines[2] == "31\t32"
