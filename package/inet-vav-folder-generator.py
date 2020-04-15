# -*- coding: utf-8 -*-
"""
Created on Fri Jan 17 10:34:48 2020

@author: SESA431606

import the INET VAV DmprPos Search to a server containing an INET Interface which hosts MCI's, and execute it.
copy the results to one or more Excel workbooks, and save to convenient location. recommend creating workbooks of no more than 30 VAVs.
run this script and select the export template, then the workbook(s), and a new export will be created for each workbook.

"""


#%% import modules

import tkinter as tk
from tkinter import filedialog
tk.Tk().withdraw()  # hides annoying tkinter window

import xlrd


#%% define user functions

def parsePath(full_path, extract='file_name', point_suffix='-DmprPos'):
    # example_ebo_path = '../../../../../INET Interface/__TEMPLATE_MCI__/__TEMPLATE_ASC__-DmprPos'
    # example_win_path = 'C:/Users/me/folder/file.xlsx'

    if extract == 'mci_name':
        mci_name = full_path.split('/')[-2]
        return mci_name

    if extract == 'asc_name':
        # paths from assumed search are followed by the point name '-DmprPos', it is removed here
        asc_name = full_path.split('/')[-1].replace(point_suffix, '')
        return asc_name

    if extract == 'file_name':
        file_name = full_path.split('/')[-1].split('.')[-2]
        # file_extension = full_path.split('.')[-1]
        # full_file_name = file_name + file_extension
        return file_name


#%% define constants and parameters

# first popup is vav graphic folder export template, second popup is vav workbook created from DmprPos search
template_name = filedialog.askopenfilename(title='Select Export Template', initialdir='./input')
workbook_names = filedialog.askopenfilenames(title='Select VAV Workbooks', initialdir='./input')

# loop over all selected workbooks
for workbook_name in workbook_names:

    # contents copied from searches includes headers, so read columns from second row
    workbook = xlrd.open_workbook(workbook_name, on_demand=True)
    sheet = workbook.sheet_by_name('Sheet1')

    point_names = sheet.col_values(0, 1)
    ebo_paths = sheet.col_values(1, 1)
    mcu_numbers = sheet.col_values(2, 1)

    mcu_numbers[:] = [mcu[:-len('PP TT')] for mcu in mcu_numbers]

    # the export template was made from an actual vav, and the names of the mci and vav were replaced manually with the following
    template_mci = '__TEMPLATE_MCI__'
    template_asc = '__TEMPLATE_ASC__'
    template_mcu = '__TEMPLATE_MCU#__'

    # new export will be named using the template workbook
    new_file_name = f"./output/System Folder - {parsePath(workbook_name)} - Export Special (2.0.4.83).xml"


#%% open files

    with open(template_name, 'rt') as template, open(new_file_name, 'wt') as new_file:


#%% convert template into lists for restructuring output file

        header_lines = []
        export_lines = []
        end_lines = []

        header_written = False
        exports_written = False

        for line in template:

            if not header_written:
                if '<ExportedObjects>' in line:
                    header_lines.append(line)
                    header_written = True
                    continue
                else:
                    header_lines.append(line)

            if header_written and not exports_written:
                if '</ExportedObjects>' in line:
                    exports_written = True
                else:
                    export_lines.append(line)

            if header_written and exports_written:
                end_lines.append(line)


#%% create ouput file from template lists and spreadsheet

        for line in header_lines:
            new_file.write(line)

        for i in range(len(point_names)):
            for line in export_lines:
                    line = line.replace(template_mci, parsePath(ebo_paths[i], extract='mci_name'))
                    line = line.replace(template_asc, parsePath(ebo_paths[i], extract='asc_name'))
                    line = line.replace(template_mcu, 'MCU ' + mcu_numbers[i])
                    line = line.replace(' " ', '" ')
                    line = line.replace(' /', '/')

                    new_file.write(line)

        for line in end_lines:
            new_file.write(line)

