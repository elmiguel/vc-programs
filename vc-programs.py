from os import listdir
from os import path
from os import mkdir
from os import walk
from json import dumps
from xlrd import open_workbook
import re

wb = open_workbook('./VC-Programs-2.xlsx')
vc_programs_sheet = wb.sheet_by_name('VC Programs')
program_template = {
    'courseId': '',
    'title': '',
    'typicalOnlinePeriodsOffered': '',
    'typicalHybridPeriodsOffered': '',
    'online': '',
    'hybrid': '',
}

num_cols = vc_programs_sheet.ncols   # Number of columns
programs = []

for row_idx in range(0, vc_programs_sheet.nrows):    # Iterate through rows
    if row_idx == 0:
        continue

    courseId = vc_programs_sheet.cell(row_idx,0).value
    _program = program_template.copy()
    _program['courseId'] = vc_programs_sheet.cell(row_idx,0).value
    _program['title'] = vc_programs_sheet.cell(row_idx, 1).value
    _program['typicalOnlinePeriodsOffered'] = vc_programs_sheet.cell(row_idx,2).value
    _program['typicalHybridPeriodsOffered'] = vc_programs_sheet.cell(row_idx, 3).value
    _program['online'] = vc_programs_sheet.cell(row_idx, 4).value
    _program['hybrid'] = vc_programs_sheet.cell(row_idx, 5).value

    programs += [_program]

with open('./vc_programs.json', 'w') as of:
    of.write(dumps({ 'programs': programs}, indent=4))
