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

## TODO: determine if unused and delete
def df_from_indices(df, indices):
    return df.iloc[indices]

def drop_columns(df, columns):
    return df.drop(columns=columns)


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
                                             'font_name': 'Segoe UI (Body)',
                                             'font_size': 8,
                                             'top': 1,
                                             'bottom': 1,
                                             'left': 0,
                                             'right': 1})

    for offset, name in enumerate(scenarionames):
        if offset == 0:
            worksheet.write(2, 7+offset, name, l_scenario_format)

        elif offset == len(scenarionames)-1:
            worksheet.write(2, 7+offset, name, r_scenario_format)

        else:
            worksheet.write(2, 7+offset, name, scenario_format)



def create_dynamic_block(worksheet, in_dict):
    # merge cells {B, C, D, E}:2
    # write 'compare loaded scenarios...' in the merged cells, blue cell
    worksheet.merge_range('B2:E2', 'Compare loaded scenarios...')

    # make column F small and G zero-width
    worksheet.set_column('F:F', 2.33)
    worksheet.set_column('G:G', 0)

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
    # if at leaf level, write row name and data at proper indentation, push row counter +1
    global row_offset
    if vals_are_lists(in_dict):
        print('dict')
        for name, datavec in in_dict.items():
            #TODO write name
            worksheet.write_row(row_offset, 7, datavec)
            row_offset += 1
    else:
        print('nodict')
        for name, nested_dict in in_dict.items():
            worksheet.write(row_offset, 0, name)
            row_offset += 1
            ind_level += 1
            create_data_rows(worksheet, nested_dict, workbook, groupnames, scenarionames, ind_level)


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

        sheet_dict = in_dict[sheetname]
        groupnames = get_groups(input_df)
        scenarionames = get_scenarios(input_df)

        create_header_block(sheetname, worksheet, sheet_dict, workbook, groupnames, scenarionames)
        create_dynamic_block(worksheet, in_dict)
        create_data_rows(worksheet, sheet_dict, workbook, groupnames, scenarionames, 0)
        # add_data_to_sheet(sheetname, sheet_dict)

    workbook.close()


create_xl_from_dict(test_dict)
