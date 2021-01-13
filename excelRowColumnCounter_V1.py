# importing pandas
import pandas as pd

# creates a dictionary for each sheet in the Data Used for Modelling_final1 excel sheet
sheets_dict = {
	'sheet 1':'2014DSPHY',
	'sheet 2':'2015DSRSQ',
	'sheet 3':'Breeding Lines',
	'sheet 4':'Premium Varieties',
	'sheet 5':'2014WSRSQ',
	'sheet 6':'Japonica',
	'sheet 7':'Sensory Data Summary'
}

# lists to input into an output table
sheet_names_list = []
number_of_rows = []
number_of_columns = []

# loops through each sheet within the excel sheet and prints the number of rows and columns in each
for key in sheets_dict:

	# accesses each sheet of the excel sheet 
	sheet_name = sheets_dict[key]
	sheet_names_list.append(sheet_name)

	# creates a temporary dataframe and extracts the number of rows and columns in each sheet
	df = pd.read_excel('Modelling_Data.xlsx', engine='openpyxl', skiprows = 1, sheet_name = sheet_name)
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