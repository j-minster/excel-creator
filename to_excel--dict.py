#!/usr/bin/env python3

import pandas as pd
import xlsxwriter
import functools
import itertools
import operator
from xlsxwriter.utility import xl_rowcol_to_cell
import sys

## import CSV file
filename = sys.argv[-1]

if not filename:
    input_df = pd.read_csv('./input/ex_input_5.csv')
else:
    input_df = pd.read_csv(filename)

## helper functions (to cull)
def compose(*functions):
    def compose2(f, g):
        return lambda x: f(g(x))
    return functools.reduce(compose2, functions, lambda x: x)

### a list containing booleans depending on whether a column in `df` has String values or not
def is_text(df):
    keys = list(df.keys())
    is_text = map(lambda key: type(df[key][0]) == str, keys)
    return list(is_text)

### determine which columns are groupings in `df`
### should return all column names except those with scenario names
def get_groups(df):
    groupdf = df.select_dtypes(include=['O'])
    groupnames = list(groupdf.columns)
    return groupnames

def get_scenarios(df):
    groups = get_groups(df)
    allcols = list(df.columns)
    for group in groups:
        allcols.remove(group)
    return allcols

# infinite dict
class NestedDict(dict):
    def __getitem__(self, key):
        if key in self:
            return self.get(key)
        else:
            value = NestedDict()
            self[key] = value
            return value

def df_to_dict(in_df):
    # create emtpy infinite dict to fill
    big_dict = NestedDict()

    # fill the dict with values from `input_df`
    for i, row in in_df.iterrows():
        dim_name = row.Dimension
        copyrow = row.copy()
        groupnames = get_groups(in_df)
        scenario_data = copyrow.drop(groupnames).to_list()
        upd_d = {dim_name : scenario_data}

        # create call to dict: ie. dict[group1][group2]...[groupn]
        gbrackets = [f'[row.{group}]' for group in groupnames]
        dict_accessor = 'big_dict' + functools.reduce(operator.concat, gbrackets[:-1])
        ex_d = eval(dict_accessor)            # old dictionary, if any
        new_d = {**ex_d, **upd_d}             # update and create new dictionary
        dict_set = dict_accessor + ' = new_d' # expression to assign old dict to new
        exec(dict_set)                        # execute dict_set

    return big_dict

# function to create the top block of a sheet
def create_header_block(sheetname, worksheet, sheet_dict, workbook, groupnames, scenarionames):
    # constants throughout
    bordercolor = "#9B9B9B"
    orangecolor = '#F7D8AA'
    greycolor =  '#F0F0F0'

    # merge cells A1 and A2
    # write the sheet name in the merged cells, large font, orange cell
    sheetname_format = workbook.add_format({'valign': 'vcenter',
                                            'bg_color': orangecolor,
                                            'font_name': 'Segoe UI Light (Heading)',
                                            'font_size': 14,
                                            'bottom': 1,
                                            'right': 1,
                                            'border_color': bordercolor,
                                            'indent': 1})

    worksheet.merge_range('A1:A2', sheetname, sheetname_format)

    # set height of sheet name cell to 36
    worksheet.set_row_pixels(0, 24)
    worksheet.set_row_pixels(1, 24)


    # set sheetname and metrics column (A) to width 43.33
    worksheet.set_column('A:A', 34.83)

    # write 'Metric' in cell A3, bold format, make the cell vertically taller (36 px), grey background
    metric_format = workbook.add_format({'bold': True,
                                         'bg_color': greycolor,
                                         'valign': 'vcenter',
                                         'font_name': 'Segoe UI (Body)',
                                         'font_size': 8,
                                         'border': 1,
                                         'border_color': bordercolor,
                                         'indent': 1})
    worksheet.write('A3', 'Metric', metric_format)
    worksheet.set_row_pixels(2, 48)


    # write all scenario* names in {H...}:3
    scenario_format = workbook.add_format({'bold': False,
                                           'bg_color': greycolor,
                                           'border_color': bordercolor,
                                           'align': 'centre',
                                           'valign': 'vcentre',
                                           'font_name': 'Segoe UI (Body)',
                                           'font_size': 8,
                                           'top': 1,
                                           'bottom': 1,
                                           'left': 0,
                                           'right': 0})
    l_scenario_format = workbook.add_format({'bold': False,
                                             'bg_color': greycolor,
                                             'border_color': bordercolor,
                                             'align': 'centre',
                                             'valign': 'vcentre',
                                             'font_name': 'Segoe UI (Body)',
                                             'font_size': 8,
                                             'top': 1,
                                             'bottom': 1,
                                             'left': 1,
                                             'right': 0})
    r_scenario_format = workbook.add_format({'bold': False,
                                             'bg_color': greycolor,
                                             'border_color': bordercolor,
                                             'align': 'centre',
                                             'valign': 'vcentre',
                                             'font_name': 'Segoe UI (Body)',
                                             'font_size': 8,
                                             'top': 1,
                                             'bottom': 1,
                                             'left': 0,
                                             'right': 1})

    scenario_col_offset = 6
    for offset, name in enumerate(scenarionames):
        if offset == 0:
            worksheet.write(2, scenario_col_offset+offset, name, l_scenario_format)

        elif offset == len(scenarionames)-1:
            worksheet.write(2, scenario_col_offset+offset, name, r_scenario_format)

        else:
            worksheet.write(2, scenario_col_offset+offset, name, scenario_format)

    # create dropdowns
    dropdown_format = workbook.add_format({'bottom': 1,
                                           'border_color': bordercolor,
                                           'align': 'centre',
                                           'valign': 'vcentre',
                                           'font_name': 'Segoe UI (Body)',
                                           'font_size': 8,
                                           'bg_color': '#DAEDF8',
                                           'fg_color': '#FFFFFF',
                                           'pattern': 16
                                           })
    input_cell_1 = 'B$3'
    worksheet.data_validation(input_cell_1, {'validate': 'list',
                                             'source': scenarionames,
                                             'input_title': 'Pick a scenario'
                                             })
    worksheet.write(input_cell_1, scenarionames[0], dropdown_format)

    input_cell_2 = 'C$3'
    worksheet.data_validation(input_cell_2, {'validate': 'list',
                                             'source': scenarionames,
                                             'input_title': 'Pick a scenario'
                                             })
    worksheet.write(input_cell_2, scenarionames[1], dropdown_format)

    # create +/- headings
    pmformat = workbook.add_format({'bottom': 1,
                                    'border_color': bordercolor,
                                    'align': 'right',
                                    'valign': 'vcentre',
                                    'font_name': 'Segoe UI (Body)',
                                    'font_size': 8,
                                    'bg_color': '#DAEDF8',})
    pmcell = 'D$3'
    worksheet.write(pmcell, '+/-', pmformat)
    pcell = 'E$3'
    worksheet.write(pcell, '%', pmformat)

    # merge cells {B, C, D, E}:2
    # write 'compare loaded scenarios...' in the merged cells, blue cell

    comp_format = workbook.add_format({'top': 1,
                                       'border_color': bordercolor,
                                       'align': 'left',
                                       'valign': 'vcentre',
                                       'font_name': 'Segoe UI Light (Headings)',
                                       'font_size': 8,
                                       'bg_color': '#DAEDF8',
                                       'indent': 1})
    worksheet.merge_range('B2:E2', 'Compare two loaded scenarios (use dropdowns)', comp_format)
    worksheet.write('F2', None, comp_format)
    worksheet.write('F3', None, pmformat)

def create_dynamic_block(worksheet, workbook, in_dict):

    # make column F small and G zero-width
    rformat = workbook.add_format({'right': 1,
                                   'border_color': '#9B9B9B'})
    worksheet.set_column('F:F', 2.33, rformat)
    # worksheet.set_column('G:G', 0)

# https://stackoverflow.com/questions/23499017/know-the-depth-of-a-dictionary
def dict_depth(d):
    if isinstance(d, dict):
        return 1 + (max(map(dict_depth, d.values())) if d else 0)
    return 0

test_dict = df_to_dict(input_df)
depth = dict_depth(test_dict)

def vals_are_lists(d):
    boollist = [isinstance(val, list) for _, val in d.items()]
    return all(boollist)

# test change
def create_data_rows(worksheet, in_dict, workbook, groupnames, scenarionames, ind_level):
    nums_offset = 6
    global row_offset
    groupformat = workbook.add_format({'bold': False,
                                       'font_name': 'Segoe UI (Body)',
                                       'font_size': 8,
                                       'right': 1,
                                       'border_color': '#9B9B9B'})
    numformat = workbook.add_format({'font_name': 'Segoe UI (Body)',
                                     'font_size': 8})
    lformat = workbook.add_format({'left': 1,
                                   'border_color': '#9B9B9B'})
    # if at leaf level, write row name and data at proper indentation, push row counter +1
    if vals_are_lists(in_dict):
        for name, datavec in in_dict.items():
            groupformat.set_indent(ind_level+1)
            worksheet.write(row_offset, 0, name, groupformat)
            worksheet.write_row(row_offset, nums_offset, datavec, numformat)

            formula_offset = row_offset + 1
            formulers = [f'=IFERROR(OFFSET($F{formula_offset}, 0, MATCH(B$3, $G$3:$DB$3, 0)), "-")',
                         f'=IFERROR(OFFSET($F{formula_offset}, 0, MATCH(C$3, $G$3:$DB$3, 0)), "-")',
                         f'=IFERROR(C{formula_offset}-B{formula_offset}, "-")',
                         f'=IFERROR(C{formula_offset}/B{formula_offset}-1, "-")']
            worksheet.write_row(row_offset, 1, formulers, numformat)

            row_offset += 1
        worksheet.write(row_offset, 0, None, groupformat)
        row_offset += 1
    else:
        for name, nested_dict in in_dict.items():
            groupformat.set_indent(ind_level)
            if ind_level == 0:
                bigname = '-- ' + name + ' --'
                groupformat.set_bold(True)
                groupformat.set_font_size(9)
                worksheet.write(row_offset, 0, bigname, groupformat)
                worksheet.write(row_offset, nums_offset, None, lformat)
                row_offset += 1
                next_ind = ind_level + 1
                create_data_rows(worksheet, nested_dict, workbook, groupnames, scenarionames, next_ind)
            elif ind_level == 1:
                groupformat.set_bold(True)
                worksheet.write(row_offset, 0, name, groupformat)
                worksheet.write(row_offset, nums_offset, None, lformat)
                row_offset += 1
                next_ind = ind_level + 1
                create_data_rows(worksheet, nested_dict, workbook, groupnames, scenarionames, next_ind)
            else:
                worksheet.write(row_offset, 0, name, groupformat)
                row_offset += 1
                next_ind = ind_level + 1
                create_data_rows(worksheet, nested_dict, workbook, groupnames, scenarionames, next_ind)


    # if not at leaf level, write row name at indentation, push row counter +1, recurse
# function to take a 'sheet dict' (from big_dict) and actually create a sheet
def add_data_to_sheet(sheetname, sheet_dict):
    worksheet = workbook.get_worksheet_by_name(sheetname)


def create_xl_from_dict(in_dict):
    # create the excel `Workbook` to write to
    workbook = xlsxwriter.Workbook('dict_test.xlsx')
    sheetnames = list(in_dict.keys())

    for sheetname in sheetnames:
        global row_offset
        row_offset = 3

        workbook.add_worksheet(sheetname)
        worksheet = workbook.get_worksheet_by_name(sheetname)
        worksheet.hide_gridlines(2)

        sheet_dict = in_dict[sheetname]
        groupnames = get_groups(input_df)
        scenarionames = get_scenarios(input_df)

        create_header_block(sheetname, worksheet, sheet_dict, workbook, groupnames, scenarionames)
        create_dynamic_block(worksheet, workbook, in_dict)
        create_data_rows(worksheet, sheet_dict, workbook, groupnames, scenarionames, 0)

    workbook.close()


create_xl_from_dict(test_dict)
