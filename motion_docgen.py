""" Produce documentation for motion axes from commissioning spreadsheet.

Retrieves information from commissioning spreadsheet and reformats into a more
readable layout. Output as PDF file for beamline documentation.
"""

import xlrd
import argparse
import pandas as pd
from subprocess import call

def process_arguments():
    desc = 'Generate motion documentation'
    parser = argparse.ArgumentParser(description = desc)
    parser.add_argument('--file-path', required=True, dest='file_path')
    parser.add_argument('--output-file', required=True, dest='output_file')

    return parser.parse_args()


def read_data_from_file(data_file):
    keys = ['Component', 
            'Axis', 
            'Controller', 
            'Axis number', 
            'Approach movement direction', 
            'Maximum speed (EU/s)', 
            'Phase current (mA)',
            'Holding current (%)']
    with open(data_file) as f:
        df = pd.io.excel.read_excel(
                f, sheetname = 'Axis data', header = 0,
                parse_cols = 'A:AW', index_col=0)
    return df

def create_data_subset(df):
    keys = ['Component', 'Axis', 'Controller', 'Motor type',
            'Axis number', 'Approach movement direction', 
            'Maximum speed (EU/s)', 'Phase current (mA)', 
            'Holding current (%)']
    index_list = ['Component', 'Axis']
    return df.loc[keys,:].transpose().set_index(index_list)

def create_tex_file(df, output_file):
    output_file_base = output_file.split('.')[0]
    output_file_tex = output_file_base + '.tex'
    of = open (output_file_tex, 'w')
    ifile = open('table_file_begin.tex', 'r')
    of.writelines([l for l in ifile.readlines()])
    of.write(df.to_latex())
    ifile = open('table_file_end.tex', 'r')
    of.writelines([l for l in ifile.readlines()])
    of.close()

def create_csv(df, output_file):
    output_file_base = output_file.split('.')[0]
    output_file_full = output_file_base + '.csv'
    df.to_csv(open(output_file_full, 'w'), encoding='utf-8')

def create_pdf(output_file):
    output_file_base = output_file.split('.')[0]
    call(['pdflatex', output_file_base + '.tex'])
        
def main():
    args = process_arguments()
    df = read_data_from_file(args.file_path)
    dfs = create_data_subset(df)
    #create_tex_file(dfs, args.output_file)
    #create_pdf(args.output_file)
    create_csv(dfs, args.output_file)

if __name__ == '__main__':
    main()

