import pandas as pd
from itertools import filterfalse
import xlsxwriter
import functools
import itertools
import operator
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell
import sys
import re

## import CSV file
## helper functions (need to cull)
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
# should return all column names except those with scenario names
def get_groups(df):
    # assuming that the columns for scenarios contain years, ie '2056'
    # groupdf = df.select_dtypes(include=['O'])
    allcols = list(df.columns)
    groupnames = [name for name in allcols if not re.search('(19|[2-9][0-9])\d{2}', name)]
    return groupnames

def drop_rows_containing(df, string):
    rowlists = [row.to_list() for _, row in df.iterrows()]
    indices = [i for i, row in enumerate(rowlists) if string in row]
    return df.drop(indices)


def df_from_clargs():
    global excel_out_path
    num_args = len(sys.argv) - 1

    match num_args:
        case 0:
            print("No input CSV file given, terminating")
            sys.exit()
            # csv_path = './input/data_larger.csv'
            # excel_out_path = './output.xlsx'
        case 1:
            csv_path = f'{sys.argv[1]}'
            excel_out_path = './output.xlsx'
            print(f'Reading CSV from: {csv_path}')
            print(f'Will write excel file to: {excel_out_path}')
        case 2:
            csv_path = f'{sys.argv[1]}'
            excel_out_path = f'{sys.argv[2]}'
            print(f'Reading CSV from: {csv_path}')
            print(f'Will write excel file to: {excel_out_path}')

    input_df = pd.read_csv(csv_path)

    return input_df

def get_scenarios(df):
    groups = get_groups(df)
    allcols = list(df.columns)
    for group in groups:
        allcols.remove(group)
    return allcols

### infinite dict
class NestedDict(dict):
    def __getitem__(self, key):
        if key in self:
            return self.get(key)
        else:
            value = NestedDict()
            self[key] = value
            return value

### convert dataframe to insane nested dictionary. Slow but works
def df_to_dict(in_df):
    # create emtpy infinite dict to fill
    big_dict = NestedDict()
    groupnames = get_groups(in_df)
    scenarionames = get_scenarios(in_df)
    # fill the dict with values from `input_df`
    for i, row in in_df.iterrows():
        dim_name = row.Mode
        scenario_data = row[scenarionames].to_list()
        upd_d = {dim_name : scenario_data}

        # create call to dict: ie. dict[group1][group2]...[groupn]
        gbrackets = [f'[row.{group}]' for group in groupnames]
        dict_accessor = 'big_dict' + functools.reduce(operator.concat, gbrackets[:-1])
        ex_d = eval(dict_accessor)            # old dictionary, if any
        new_d = {**ex_d, **upd_d}             # update and create new dictionary
        dict_set = dict_accessor + ' = new_d' # expression to assign old dict to new
        exec(dict_set)                        # execute dict_set

    return big_dict

### used to shorten the names of sheets to meet Excel's 31char limit
def shorten_long_sheetnames(in_df):
    rlist = [('Average'     , 'Avg.'),
             ('Distance'    , 'Dist.'),
             ('Distances'   , 'Dists'),
             ('Terminating' , 'Term.'),
             ('Originating' , 'Orig.'),
             ('Population'  , 'Pop.')]

    def replace_multi(rlist, name):
        for frm, to in rlist:
            if len(name) > 31:
                name = name.replace(frm, to)
        return name

    groups = get_groups(in_df)
    # sheetname_col = groups[0]
    sheetname_col = list(in_df.columns)[0]
    sheetnames = list(in_df.loc[:, sheetname_col])
    replaced_sheetnames = [replace_multi(rlist, name) for name in sheetnames]

    if in_df[sheetname_col].to_list() != replaced_sheetnames:
        print('Sheet names shortened to be < 31 chars')

    in_df[sheetname_col] = replaced_sheetnames
    return in_df

### return a list of 'sheetnames' for the excel file
def get_sheetnames(in_df):
    groups = get_groups(in_df)
    sheetname_col = groups[0]
    sheetnames = in_df.loc[:, sheetname_col]
    sheetset = set(sheetnames)
    return sheetset

### create individual dictionaries for sheets rather than one huge dictionary for the whole dataframe
def create_sheet_dict(in_df, sheetname):
    sheetnames = get_sheetnames(in_df)
    majorcol = list(in_df.columns)[0]
    sub_df = in_df[in_df[majorcol].isin([sheetname])]

    d1 = df_to_dict(sub_df)

    idx = list(sheetnames).index(sheetname)
    print(f'made \'{sheetname}\' dict')
    print(f'(Creating {idx+1} of {len(sheetnames)} total sheets)')

    return d1

### function to create the top block of a sheet
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
                                             # 'source': scenarionames,
                                             'source': '=$G$3:$XFD$3',
                                             'input_title': 'Pick a scenario'
                                             })
    worksheet.write(input_cell_1, scenarionames[0], dropdown_format)

    input_cell_2 = 'C$3'
    worksheet.data_validation(input_cell_2, {'validate': 'list',
                                             # 'source': scenarionames,
                                             'source': '=$G$3:$XFD$3',
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
    lilcell_format = workbook.add_format({'top': 1,
                                          'right': 1,
                                          'border_color': bordercolor,
                                          'align': 'left',
                                          'valign': 'vcentre',
                                          'font_name': 'Segoe UI Light (Headings)',
                                          'font_size': 8,
                                          'bg_color': '#DAEDF8',
                                          'indent': 1})

    worksheet.merge_range('B2:E2', 'Compare two loaded scenarios (use dropdowns)', comp_format)
    worksheet.write('F2', None, lilcell_format)
    worksheet.write('F3', None, pmformat)


def create_dynamic_block(worksheet, workbook):

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

### check whether the values (not the keys) in the dictionary `d` are lists
def vals_are_lists(d):
    boollist = [isinstance(val, list) for _, val in d.items()]
    return all(boollist)

### input the data for each sheet (warning: recursion)
def create_data_rows(worksheet, in_dict, workbook, groupnames, scenarionames, ind_level, sheetname, writeIndexHeader):
    nums_offset = 6
    global row_offset
    global index_row_offset

    groupformat = workbook.add_format({'bold': False,
                                       'font_name': 'Segoe UI (Body)',
                                       'font_size': 8,
                                       'right': 1,
                                       'border_color': '#9B9B9B'})
    index_groupformat = workbook.add_format({'bold': True,
                                             'font_name': 'Arial Narrow',
                                             'font_size': 11,
                                             'font_color': 'blue',
                                             'underline': 1,
                                             'indent': 1})
    index_headerformat = workbook.add_format({'bold': True,
                                              'bottom': 1,
                                              'font_name': 'Arial Narrow',
                                              'font_size': 16,
                                              'indent': 0})
    index_ulformat = workbook.add_format({'bottom': 1})
    numformat = workbook.add_format({'font_name': 'Segoe UI (Body)',
                                     'font_size': 8,
                                     'num_format': '#,##0.000'})
    pctformat = workbook.add_format({'font_name': 'Segoe UI (Body)',
                                     'font_size': 8,
                                     'num_format': '0.0%'})
    lformat = workbook.add_format({'left': 1,
                                   'border_color': '#9B9B9B'})
    if writeIndexHeader:
        index_row_offset += 1
        link_string = f'internal:{sheetname!r}!A1'
        index_sheet.write_url(index_row_offset, 1, link_string)
        index_sheet.write(index_row_offset, 1, sheetname, index_headerformat)
        index_row_offset += 1

    # if at leaf level, write row name and data at proper indentation, push row counter +1
    # else, write metric name and recursively call `create_data_rows`, on items in `in_dict` pushing indentation counter +1
    if vals_are_lists(in_dict):
        for name, datavec in in_dict.items():
            if name == '--':
                row_offset -= 1
                worksheet.write_row(row_offset, nums_offset, datavec, numformat)

                formula_offset = row_offset + 1
                formulers = [f'=IFERROR(OFFSET($F{formula_offset}, 0, MATCH(B$3, $G$3:$DB$3, 0)), "-")',
                             f'=IFERROR(OFFSET($F{formula_offset}, 0, MATCH(C$3, $G$3:$DB$3, 0)), "-")',
                             f'=IFERROR(C{formula_offset}-B{formula_offset}, "-")']
                pct_cell = f'=IFERROR(C{formula_offset}/B{formula_offset}-1, "-")'

                worksheet.write_row(row_offset, 1, formulers, numformat)
                worksheet.write(row_offset, 4, pct_cell, pctformat)

                row_offset += 1
            else:
                groupformat.set_indent(ind_level+1)
                worksheet.write(row_offset, 0, name, groupformat)
                worksheet.write_row(row_offset, nums_offset, datavec, numformat)
                formula_offset = row_offset + 1
                formulers = [f'=IFERROR(OFFSET($F{formula_offset}, 0, MATCH(B$3, $G$3:$DB$3, 0)), "-")',
                             f'=IFERROR(OFFSET($F{formula_offset}, 0, MATCH(C$3, $G$3:$DB$3, 0)), "-")',
                             f'=IFERROR(C{formula_offset}-B{formula_offset}, "-")']
                pct_cell = f'=IFERROR(C{formula_offset}/B{formula_offset}-1, "-")'
                worksheet.write_row(row_offset, 1, formulers, numformat)
                worksheet.write(row_offset, 4, pct_cell, pctformat)
                row_offset += 1

        worksheet.write(row_offset, 0, None, groupformat)
        row_offset += 1
    else:
        for name, nested_dict in in_dict.items():
            groupformat.set_indent(ind_level)
            if ind_level == 0:
                # write `bigname` to sheet
                bigname = '-- ' + name + ' --'
                groupformat.set_bold(True)
                groupformat.set_font_size(9)
                worksheet.write(row_offset, 0, bigname, groupformat)
                worksheet.write(row_offset, nums_offset, None, lformat)

                # create index links
                to_cell = xl_rowcol_to_cell(row_offset, 0)
                link_string = f'internal:{sheetname!r}!{to_cell}'
                index_sheet.write_url(index_row_offset, 1, link_string)
                index_sheet.write(index_row_offset, 1, name, index_groupformat)
                index_row_offset += 1

                # bump `row_offset` and `next_ind` and recurse again for each `nested_dict`
                row_offset += 1
                next_ind = ind_level + 1
                create_data_rows(worksheet, nested_dict, workbook, groupnames, scenarionames, next_ind, sheetname, False)
            elif ind_level == 1:
                groupformat.set_bold(True)
                worksheet.write(row_offset, 0, name, groupformat)
                worksheet.write(row_offset, nums_offset, None, lformat)
                row_offset += 1
                next_ind = ind_level + 1
                create_data_rows(worksheet, nested_dict, workbook, groupnames, scenarionames, next_ind, sheetname, False)
            else:
                if name == '--':
                    next_ind = ind_level
                    create_data_rows(worksheet, nested_dict, workbook, groupnames, scenarionames, next_ind, sheetname, False)
                else:
                    worksheet.write(row_offset, 0, name, groupformat)
                    row_offset += 1
                    next_ind = ind_level + 1
                    create_data_rows(worksheet, nested_dict, workbook, groupnames, scenarionames, next_ind, sheetname, False)

### ties together all previously defined functions
def create_xl_from_df(in_df):
    global excel_out_path
    workbook = xlsxwriter.Workbook(excel_out_path)

    # create the index sheet
    workbook.add_worksheet('Index')
    global index_sheet
    index_sheet = workbook.get_worksheet_by_name('Index')
    global index_row_offset
    index_row_offset = 0

    in_df = shorten_long_sheetnames(in_df)
    in_df = in_df.fillna('')
    scenarionames = get_scenarios(in_df)
    sheetnames = get_sheetnames(in_df)

    for sheetname in sheetnames:
        print(f'creating {sheetname} sheet')
        global row_offset
        row_offset = 3

        workbook.add_worksheet(sheetname)
        worksheet = workbook.get_worksheet_by_name(sheetname)
        worksheet.set_default_row(18)
        worksheet.hide_gridlines(2)

        sheet_dict = create_sheet_dict(in_df, sheetname)
        groupnames = get_groups(in_df)

        create_data_rows(worksheet, sheet_dict, workbook, groupnames, scenarionames, 0, sheetname, writeIndexHeader=True)
        create_header_block(sheetname, worksheet, sheet_dict, workbook, groupnames, scenarionames)
        create_dynamic_block(worksheet, workbook)

        worksheet.set_default_row(hide_unused_rows=True)
        worksheet.autofit()

        # setting column widths
        worksheet.set_column('A:A', 33.33)
        worksheet.set_column('B:B', 10.67)
        worksheet.set_column('C:C', 10.67)
        worksheet.set_column('D:D', 10.67)
        worksheet.set_column('E:E', 5)
        worksheet.set_column('G:XFD', 10.67)
        print(f'{sheetname} sheet done')

    print(f'Writing excel file to disk as {excel_out_path}, please wait')
    index_sheet.autofit()
    index_sheet.hide_gridlines(2)
    workbook.close()


### do everything
input_df = df_from_clargs()
create_xl_from_df(input_df)
