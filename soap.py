'''
*************************************************
* Soapy                                         *
* 		                                        *
* A Simple script for creating SLP SOAP notes.  *
* 											    *
*************************************************
'''

import argparse
import os
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.compat import range
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

parser = argparse.ArgumentParser()
parser.add_argument('filename')
args = parser.parse_args()
filename = args.filename

# open workbook if exists
dest_filename = filename + '.xlsx'
if os.path.isfile(dest_filename):
	wb = load_workbook(filename = dest_filename)
	new_wb = False
else:
	wb = Workbook()
	new_wb = True

print('\n***********  SOAPY ***********')
print('Be sure that excel is not open!\n')
patient = input('Patient name: ' )
num_goals = int(input('Number of goals: ' ))
goals = []

# get goals from user
for i in range(0,num_goals):
	goal = input('Goal ' + str(i + 1) + ': ')
	goals.append(goal)


if new_wb:
	ws = wb.active
	ws.title = patient
else:
	if patient in wb.sheetnames: # delete sheet if it already exists
		wb.remove(wb[patient])
	ws = wb.create_sheet(title=patient)

ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
ws.print_options.gridLines = True
ws.page_margins.left = 0.25
ws.page_margins.right = 0.25
ws.page_margins.bottom = 0.5

ft_heading = Font(size=11, bold=True)

# merge cells and input patient name
ws.merge_cells('A1:F1')
a1 = ws['A1']
a1.font = ft_heading
a1.value = 'Pt: ' + patient

header_cells = {'A': 'Date', 'B': 'Subjective Notes', 'C': 'Short-term Goals', 'D': 'Data (+/-)', 'E': 'Cues Given', 'F': 'Assessment/Plan'}

# set appropriate column widths
column_widths = {'A': 10, 'B': 20, 'C': 35, 'D': 24, 'E': 20, 'F': 20}
for column, width in column_widths.items():
	ws.column_dimensions[column].width = width

# make cell heading
def insertHeading(startRow):
	for k, v in header_cells.items():
		cell = k + str(startRow)
		ws[cell].font = ft_heading
		ws[cell].value = v

# insert goals to cells
def insertGoals(startRow):
	for goal in range(startRow, num_goals + startRow):
		cell_num = 'C' + str(goal)
		cell = ws[cell_num].value = goals[goal - startRow]

def inputData(goal_index, heading_index, offset):
	for i in range(0, 4):
		ws.merge_cells('B' + str(goal_index) +':B' + str(goal_index + offset - 1)) # merges subjective note cells
		insertHeading(heading_index)
		insertGoals(goal_index)
		heading_index += offset + 3
		goal_index += offset + 3


heading_index = 4 # where to start inserting heading row
goal_index = 5	# where to start inserting goals


custom_offset = 4 # Use if num_goals < 4 to make room for subjective notes. 
use_custom_offset = False

# If num_goals < 4, we use a custom offset.
# The custom offset guarantees that the spreadsheet will have enough room
# to take subjective notes. 
# In other words, we will have at least four merged rows in column B to 
# take notes. 
if num_goals < custom_offset:
	use_custom_offset = True

if use_custom_offset == True:
	inputData(goal_index,heading_index,custom_offset)
else:
	inputData(goal_index,heading_index,num_goals)

wb.save(filename=dest_filename)
print('\n' + dest_filename + ' created.\n')
