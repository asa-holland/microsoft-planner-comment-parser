# comment_parser.py

# import pandas to handle the excel file
import pandas as pd

# import numpy to handle the missing values
import numpy as np

def parse(excel_file):
	"""
	Separates comments in an Apps4Pro Planner Manager Microsoft Planner exported excel file
	:param excel_file: takes a string of a path to an excel file exported from Microsoft Planner (using Apps4Pro Planner Manager) with comments that should be separated
	:return: string of a path to the newly created cleaned excel file
	"""

	# Read the excel file into pandas
	raw_df = pd.read_excel(io=excel_file)

	# Split comments based on their #### dividers, then explode out the rows so that each row entry is backfilled based on the task details
	new_df = (raw_df.set_index(['Plan Name', 'Task Title', 'Bucket', 'Assigned to', 'Start date', 'Due date', 
		'Completed by', 'Completed date']).apply(lambda x: x.str.split('###############').explode()).reset_index())

	# Add a new column for the commenter
	new_df['Commenter'] = new_df['Comments'].str.extract('Posted by: ([A-Za-z -]+) - ')

	# Add a new column for the comment date
	new_df['Comment Date'] = new_df['Comments'].str.extract('- ([0-9\/]*) [0-9: AP]*M')

	# add a new column for comment time
	new_df['Comment Time'] = new_df['Comments'].str.extract('- [0-9\/]* ([0-9: AP]*M)')

	# add a new column for the comment itself
	new_df['Comment'] = new_df['Comments'].str.extract('-\d\d:\d\d\n\n([\w\d\s\/,\(\)-]*)(?=These comments are about the task|Reply)')

	# Clean up the comment, removing any whitespace and newlines
	new_df["Comment"] = new_df["Comment"].str.strip()

	# remove the original comments column
	new_df.drop('Comments', axis='columns', inplace=True)

	# create a new filename (in the same folder as the original file)
	new_excel_file = str(excel_file).replace('.xlsx', ' Cleaned Comments.xlsx')

	# save the cleaned file to the original directory with a cleaned name
	new_df.to_excel(excel_writer=new_excel_file)

	# output to the user that the parsing has been completed, and return the path of the newly created file as a string
	print('Comments for exported Microsoft Planner worksheet have been successfully separated into individual rows.')
	return new_excel_file 