from openpyxl import Workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

patient = input("\nPatient name: " )

wb = Workbook()

# later we wil update filname to day of week?
dest_filename = 'testSoap.xlsx'

# worksheet title is patients name
ws = wb.active
ws.title = patient

ft_heading = Font(size=14, bold=True)

# merge cells and input patient name
ws.merge_cells('A1:D1')
a1 = ws['A1']
a1.font = ft_heading
a1.value = 'Patient: ' + patient











wb.save(filename=dest_filename)