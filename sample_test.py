# sample_test.py

# Import os to navigate to current installation location
import os

# Local import of comment parser
from commentParser import comment_parser

# Get the path to the current script file and use it to find the path to the sample data provided
path_to_test = os.path.realpath(__file__)
path_to_sample_data = path_to_test.replace('\\sample_test.py', '\\tests\\Project Workflow Export (Sample Data).xlsx')

# Run the code to parse the file, creating a new version of the parsed data
comment_parser.parse(excel_file=path_to_sample_data)