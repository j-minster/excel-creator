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
def get_groups(df):
    in_df = df.loc[:,~df.columns.str.startswith('Scenario')]
    groupnames = list(in_df.columns)
    return groupnames

def get_offsets(groups):
    sgroups = sorted(groups)
    length = len(sgroups)
    biglist = [(k, list(g)) for k, g in itertools.groupby(sgroups, lambda group: group[0][0])]
    lens = [[l[2] for l in lst] for g, lst in biglist]
    accs = [list(itertools.accumulate(lst, initial=0)) for lst in lens]
    final = [lst[:-1] for lst in accs]
    spacer = [list(itertools.accumulate([1] * len(lst), initial=0))[:-1] for lst in final]
    flat_spacer = list(itertools.chain(*spacer))
    flat_final = list(itertools.chain(*final))
    final = list(map(operator.add, flat_final, flat_spacer))
    return final

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


group_headers = get_groups(input_df)
gdfs = input_df.groupby(group_headers)

df_lengths = [len(gdf) for _, gdf in gdfs]

groups = gdfs.groups

group_indices_dict = dict(groups).values()
group_names = list(dict(groups))

indices = [list(vals) for vals in group_indices_dict]

group_names_and_indices = list(zip(group_names, indices, df_lengths))

offsets = get_offsets(group_names_and_indices)

sheet_indices_offset = list(zip(group_names, indices, offsets))

## main function
def append_to_sheet(metric, df, writer):
    sheet_name = metric[0][0]
    metric_name = metric[0][1]
    indices = metric[1]
    row_offset = metric[2]

    workbook = writer.book
    sheets = writer.sheets

    write_header = True
    dimension_row_offset = 1
    metric_spacer = 0
    cpanel_row_offset = 2
    cpanel_column_offset = 7
    merges = True

    if sheet_name in sheets.keys():
        write_header = False
        dimension_row_offset = 0
        metric_spacer = 1
        merges = False

    sliced_df = df_from_indices(df, indices)
    # TODO: fix this and make automatic, the columns to drop

    metric_df = sliced_df['Dimension']
    num_df = drop_columns(sliced_df, ['Group', 'Metric', 'Dimension'])

    scenario_names = list(num_df.columns)

    # metric_df.to_excel(writer,
    #                    sheet_name=sheet_name,
    #                    startrow=row_offset + metric_spacer + cpanel_row_offset + dimension_row_offset,
    #                    startcol=1,
    #                    header=False,
    #                    index=False)


    num_df.to_excel(writer,
                    sheet_name=sheet_name,
                    startrow=row_offset + metric_spacer + cpanel_row_offset,
                    startcol=1+cpanel_column_offset,
                    header=write_header,
                    index=False)


    metric_format = workbook.add_format({'bold': True, 'bg_color': '#F0F0F0'})

    worksheet = sheets[sheet_name]
    worksheet.hide_gridlines(2)

    dim_format = workbook.add_format({
        'bg_color': '#F0F0F0',
        'right': 1,
        'italic': 1
    })
    dim_startrow = row_offset + metric_spacer + cpanel_row_offset + dimension_row_offset
    dim_startcol = 1
    for delta_row, metric in enumerate(metric_df.values):
         worksheet.write(dim_startrow + delta_row,
                         dim_startcol,
                         metric,
                         dim_format)

    # Add a header format.
    num_df_header_format = workbook.add_format({
        'bold': True,
        'valign': 'centre',
        'align': 'centre',
        'fg_color': '#D7D4F0',
        'border': 0})

    # Write the column headers with the defined format.
    if write_header:
        for col_num, value in enumerate(num_df.columns.values):
            worksheet.write(row_offset + metric_spacer + cpanel_row_offset,
                            1 + cpanel_column_offset + col_num,
                            value,
                            num_df_header_format)

    # write sheet name in top left corner
    if merges:
        sheetname_format = workbook.add_format( {
            "align": "center",
            "valign": "vcenter",
            "bg_color": '#F7D8AA'
        })
        worksheet.merge_range('A1:A2', sheet_name, sheetname_format)

    metric_row_start = row_offset + dimension_row_offset + metric_spacer + cpanel_row_offset
    worksheet.write(metric_row_start, 0,
                    metric_name,
                    metric_format)

    # write dropdown menus for choosing scenarios
    input_cell_1 = 'C$3'
    worksheet.data_validation(input_cell_1, {'validate': 'list',
                                     'source': scenario_names,
                                     'input_title': 'Pick a scenario'
                                    })
    scenario_format = workbook.add_format({"bold": True, "align": 'centre'})
    worksheet.write(input_cell_1, scenario_names[0], scenario_format)

    input_cell_2 = 'D$3'
    worksheet.data_validation(input_cell_2, {'validate': 'list',
                                     'source': scenario_names,
                                     'input_title': 'Pick a scenario'
                                    })
    worksheet.write(input_cell_2, scenario_names[1], scenario_format)

    # write the 'lookup' block
    lookup_row_range = list(range(metric_row_start, metric_row_start + len(metric_df)))
    lookup_col_range = list(range(2, 3+1))
    lookup_indices = list(itertools.product(lookup_row_range, lookup_col_range))

    # set column widths of columns next to the 'comparison' block
    worksheet.set_column(cpanel_column_offset, cpanel_column_offset, 0)
    worksheet.set_column(cpanel_column_offset-1, cpanel_column_offset-1, 2)


    for index in lookup_indices:
        zero_w_cell = xl_rowcol_to_cell(index[0], cpanel_column_offset, col_abs=True)
        if index[1] == 2:
            lookup_formula = f'=IFERROR(OFFSET({zero_w_cell}, 0, MATCH({input_cell_1}, $I$3:$DB$3, 0)), \"-\")'
        else:
            lookup_formula = f'=IFERROR(OFFSET({zero_w_cell}, 0, MATCH({input_cell_2}, $I$3:$DB$3, 0)), \"-\")'
        worksheet.write_formula(index[0], index[1], lookup_formula)

    # write the 'comparison' block`
    comparison_col_range = [index + 2 for index in lookup_col_range]
    comparison_row_range = lookup_row_range
    comparison_indices = list(itertools.product(comparison_row_range, comparison_col_range))

    ## write names of comparisons
    comp_format = workbook.add_format({'bold': True, 'right': 1, 'align': 'centre'})
    worksheet.write('E3', '+/-', scenario_format)
    worksheet.write('F3', '%', comp_format)
    worksheet.set_column('F:F', 7)
    worksheet.set_column('E:E', 9)

    for index in comparison_indices:
        ccell = 'C' + str(index[0] + 1)
        dcell = 'D' + str(index[0] + 1)
        if index[1] == 4:
            comparison_formula = f'=IFERROR({dcell}-{ccell}, "-")'
        else:
            comparison_formula = f'=IFERROR({dcell}/{ccell}-1, "-")'
        worksheet.write_formula(index[0], index[1], comparison_formula)

    # format columns to add borders
    left_border_format = workbook.add_format({ 'left': 1 })
    right_border_format = workbook.add_format({ 'right': 1 })

    worksheet.set_column(1, 1, cell_format = right_border_format)

    # second column line
    worksheet.set_column(cpanel_column_offset-2, cpanel_column_offset-2,
                         6,
                         cell_format = right_border_format)


# running the main function on each group
writer = pd.ExcelWriter('test.xlsx', engine='xlsxwriter')
workbook = writer.book
bold = workbook.add_format({'bold': True})

for metric in sheet_indices_offset:
    append_to_sheet(metric, input_df, writer)

for sheet in writer.sheets:
    writer.sheets[sheet].autofit()
    writer.sheets[sheet].merge_range('C2:F2', 'Compare two loaded scenarios (use dropdowns)')
        # worksheet.set_column('C:C', 8)

writer.close()


d = NestedDict()
for i, row in input_df.iterrows():
    dim_name = row.Dimension
    copyrow = row.copy()
    scenario_data = copyrow.drop(['Group', 'Group2', 'Metric', 'Dimension']).to_list()
    upd_d = {dim_name : scenario_data}
    ex_d = d[row.Group][row.Group2][row.Metric]
    new_d = {**ex_d, **upd_d}
    d[row.Group][row.Group2][row.Metric] = new_d
