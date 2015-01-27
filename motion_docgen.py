""" Produce documentation for motion axes from commissioning spreadsheet.

Retrieves information from commissioning spreadsheet and reformats into a more
readable layout. Output as PDF file for beamline documentation.
"""

import xlrd
import argparse
import pandas as pd

def process_arguments():
	desc = 'Generate motion documentation'
	parser = argparse.ArgumentParser(description = desc)
	parser.add_argument('--file_path', required=True)

	return parser.parse_args(


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

def create_tex_file(df):
	of = open (output, 'w')
	ifile = open('table_file_begin.tex', 'r')
	of.writelines([l for l in ifile.readlines()])
	of.write(df.loc[keys,:].to_latex())
	ifile = open('table_file_end.tex', 'r')
	of.writelines([l for l in
	ifile.readlines()])
	of.close()
		
		


def main():
	args = process_arguments()
	df = read_data_from_file(args.file)
	create_tex_file(df)

if __name__ = '__main__':
	main()

