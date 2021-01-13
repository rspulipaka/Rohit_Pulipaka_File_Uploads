# importing pandas and os.path to file existing files
import pandas as pd
import os.path
from os import path

# asks for the user to input an excel file name and checks for its existence
file_name = input("Enter excel file name: ") + ".xlsx"
does_exist = path.exists(file_name)

if not does_exist:
	print("No such file or directory: '{}".format(file_name)+"'")
else:
	df = pd.read_excel(file_name, engine='openpyxl', sheet_name = None)

	sheet_names_list = []
	number_of_rows = []
	number_of_columns = []

	# extracts the individual sheet names in a given excel sheet
	sheet_names = df.keys()

	# appends the sheet names to a list
	for keys in sheet_names:
		sheet_names_list.append(keys)

	# loops through each sheet within the excel sheet and prints the number of rows and columns in each
	for sheets in sheet_names_list:
		df = pd.read_excel(file_name, engine='openpyxl', sheet_name = sheets)
		df.drop(df.columns[df.columns.str.contains('unnamed', case = False, na = False)], axis = 'columns', inplace = True)

		# adds the number of rows and columns for each sheet into a list
		number_of_rows.append(len(df))
		number_of_columns.append(len(df.columns))

	# creates a dictionary of the output lists
	formatted_report_dict = {
		'Sheet Name':sheet_names_list,
		'Number of Rows':number_of_rows,
		'Number of Columns':number_of_columns
	}

	# outputs a data frame to display the number of rows and columns of each sheet
	report_df = pd.DataFrame(formatted_report_dict)
	print(report_df)