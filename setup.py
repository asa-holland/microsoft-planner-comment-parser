# setup.py

from distutils.core import setup

from setuptools import find_packages

import os


# Pull project description (long) from project readme
current_directory = os.path.dirname(os.path.abspath(__file__))

# Linke to the requirements text file
requirementPath = current_directory + '/requirements.txt'
install_requires = []
if os.path.isfile(requirementPath):
    with open(requirementPath) as f:
        install_requires = f.read().splitlines()

try:

    with open(os.path.join(current_directory,'README.md'), encoding='utf-8') as f:

        long_description = f.read()

except Exception:

    long_description = ''

setup(

	# Project name: 

	name='Microsoft Planner Coment Parser',

	# Packages to include in the distribution: 

	packages=find_packages('pandas', 'numpy', 'openpyxl'),

	# Project version number:

	version='1.0.0',

	# List a license for the project, eg. MIT License

	license='MIT License',

	# Short description of your library: 

	description='Provides a script to separate comments in an Apps4Pro export of Microsoft Planner Excel workbook.',

	# Long description of your library: 

	long_description=long_description,

	long_description_content_type='text/markdown',

	# Your name: 

	author='Asa Holland',

	# Your email address:

	author_email='hollandasa@gmail.com',

	# Link to your github repository or website: 

	url='https://github.com/asa-holland/microsoft-planner-comment-parser',

	# Download Link from where the project can be downloaded from:

	download_url='https://github.com/asa-holland/microsoft-planner-comment-parser',

	# List of keywords: 

	keywords=['Python', 'pandas', 'numpy', 'openpyxl', 'Microsoft', 'planner', 'comments', 'comment', 'export', 'Apps4Pro', 'Excel', 'workbook', 'parse', 'separate', 'clean', 'manager'],

	# List project dependencies: 

	install_requires=install_requires,

	# https://pypi.org/classifiers/ 

	classifiers=['Environment :: Win32 (MS Windows)', 'Development Status :: 5 - Production/Stable', 'Intended Audience :: End Users/Desktop', 'License :: Free For Educational Use', 'Programming Language :: Python :: 3.7']

)