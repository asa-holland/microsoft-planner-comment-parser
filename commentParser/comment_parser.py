# comment_parser.py

# import pandas to handle the excel file
import pandas as pd

# import regex to handle the searching
import re

# While this could be done natively in pandas we bypass this to avoid having to deal with Series/dataframe transformation

def parse(excel_file):

	# Read the excel file into pandas
	raw_df = pd.read_excel(io=excel_file)

	# Convert all NANs to None so that python (not pandas) can handle the missing values
	df = raw_df.where(pd.notnull(raw_df), None)

	# Convert the list of excel cells in the comments column to a standard python list
	comments_column_as_list = df['Comments'].tolist()

	

	# Define the pattern we need to extract the comment information
	comment_pattern = r'Posted by: ([A-Za-z -]+) - ([0-9\/]*) ([0-9: AP]*M) -\d\d:\d\d\n\n([\w\d\s\/,\(\)-]*)(?=These comments are about the task|Reply)'

	# Separated comments
	separated_comments = []

	# Iterate over the data
	for cell in comments_column_as_list:

		# For each individual cell, check if the content is empty. If so, we can skip over any parsing for this cell
		if cell is None:
			separated_comments += []

		# Otherwise, parse as usual
		else:

			# Within each cell, a new dictionary using the values we want to parse out
			new_columns = {
				'commenter': [],
				'date': [],
				'time': [],
				'comment': [],
			}
			
			# perform a search using our regex code to return a list of tuples, each tuple containing a set of matched information in the form:
			# (Commenter Name, Comment Date, Comment Time, Comment)
			cell_results = re.findall(pattern=comment_pattern, string=cell)

			print(cell_results)

			# append the separated comments to the list
			separated_comments += [cell_results]


	print(separated_comments)

	pass