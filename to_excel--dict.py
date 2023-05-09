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
## TODO: fix logic aroudn which columns are dropped, they're not going to contain 'Scenario'...
def get_groups(df):
    in_df = df.loc[:,~df.columns.str.startswith('Scenario')]
    groupnames = list(in_df.columns)
    return groupnames

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

def df_to_dict(input_df):
    # create emtpy infinite dict to fill
    big_dict = NestedDict()

    # fill the dict with values from `input_df`
    for i, row in input_df.iterrows():
        dim_name = row.Dimension
        copyrow = row.copy()
        groupnames = get_groups(input_df)
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


# create emtpy infinite dict to fill
big_dict = NestedDict()

# fill the dict with values from `input_df`
for i, row in input_df.iterrows():
    dim_name = row.Dimension
    copyrow = row.copy()
    groupnames = get_groups(input_df)
    scenario_data = copyrow.drop(groupnames).to_list()
    upd_d = {dim_name : scenario_data}

    # create call to dict: ie. dict[group1][group2]...[groupn]
    gbrackets = [f'[row.{group}]' for group in groupnames]
    dict_accessor = 'big_dict' + functools.reduce(operator.concat, gbrackets[:-1])
    ex_d = eval(dict_accessor)            # old dictionary, if any
    new_d = {**ex_d, **upd_d}             # update and create new dictionary
    dict_set = dict_accessor + ' = new_d' # expression to assign old dict to new
    exec(dict_set)                        # execute dict_set


# function to create the top block of a sheet
def create_header_block(sheetname, worksheet, sheet_dict):
    # merge cells A1 and A2
    # write the sheet name in the merged cells, large font, orange cell
    sheetname_format = workbook.add_format({'align': 'center',
                                            'valign': 'vcenter',
                                            'bg_color': '#F7D8AA',
                                            'font_name': 'Segoe UI Light (Heading)',
                                            'font_size': 14})

    worksheet.merge_range('A1:A2', sheetname, sheetname_format)

    # set height of sheet name cell to 36
    worksheet.set_row_pixels(0, 24)
    worksheet.set_row_pixels(1, 24)


    # set sheetname and metrics column (A) to width 43.33
    worksheet.set_column('A:A', 43.33)

    # write 'Metric' in cell A3, bold format, make the cell vertically taller (36 px), grey background
    metric_format = workbook.add_format({'bold': True,
                                         'bg_color': '#F0F0F0',
                                         'valign': 'vcenter',
                                         'font_name': 'Segoe UI (Body)',
                                         'font_size': 8})
    worksheet.write('A3', 'Metric', metric_format)
    worksheet.set_row_pixels(2, 48)

    # merge cells {B, C, D, E}:2
    # write 'compare loaded scenarios...' in the merged cells, blue cell
    worksheet.merge_range('B2:E2', 'Compare loaded scenarios...')

    # make column F small and G zero-width
    worksheet.set_column('F:F', 2.33)
    worksheet.set_column('G:G', 0)

    # write all scenario* names in {H...}:3
    scenario_format = workbook.add_format({'bold': False,
                                           'align': 'centre',
                                           'font_name': 'Segoe UI (Body)',
                                           'font_size': 8})
    scenario_names = input_df.columns


# def create_dynamic_block(worksheet, in_dict):

# def create_data_rows(worksheet, in_dict):


# create the excel `Workbook` to write to
workbook = xlsxwriter.Workbook('dict_test.xlsx')


# function to take a 'sheet dict' (from big_dict) and actually create a sheet
def add_data_to_sheet(sheetname, sheet_dict):
    worksheet = workbook.get_worksheet_by_name(sheetname)

    create_header_block(sheetname, worksheet, sheet_dict)
    # create_dynamic_block(worksheet, in_dict)
    # create_data_rows(worksheet, in_dict)


sheetnames = list(big_dict.keys())

# create all sheets
for sheetname, sheet_dict in zip(sheetnames, big_dict):
    if sheetname not in workbook.sheetnames:
        workbook.add_worksheet(sheetname)

    add_data_to_sheet(sheetname, sheet_dict)


workbook.close()
