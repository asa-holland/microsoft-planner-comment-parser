# Microsoft Planner Comment Parser
Short script to separate comments in Apps4Pro Planner Manager Excel files created from exporting Microsoft Planner tasks.


<!-- ABOUT THE PROJECT -->
## About The Project

![Comment Parser Example Use][use-example]

Microsoft Planner is a useful tool used to track Tasks in various stages of completion. Microsft Planner's [web app](https://tasks.office.com) offers the feature to [export an entire plan](https://docs.microsoft.com/en-us/power-platform/admin/using-word-templates-dynamics-365) to better manipulate and analyzed Plan data. However, Microsoft's exported Excel file does not support the inclusion of comments made on Tasks.

If tracking task comments is useful to you, [Apps4Pro](https://apps4.pro/Home.aspx) created software called the [Planner Manager](https://apps4.pro/planner-manager.aspx) which at the time of this script had a downloadable free trial version. The Planner Manager offers many added features to manipulate Microsoft Planner data, including the ability to extract comments from all Tasks in a given Plan.

However, the Planner Manager exports all comments for a given Task as a single cell in the exported `.xlsx` file:

![Sample Comments Uncleaned][use-file-before]

In an effort to make comments more easy to work with in Excel, this script parses the single cell comment for each task and separates it into four columns (Commenter, Comment Date, Comment Time, and Comment), and separates individual comments for any particular Task into unique rows (backfilling Task information as needed):

![Sample Comments Cleaned][use-file-after]


### Built With

* [pandas](https://pandas.pydata.org/): a Python library for data analysis
* [numpy](https://numpy.org/): a Python library for handling arrays and scientific computing
* [openpyxl](https://openpyxl.readthedocs.io/en/stable/): a Python library that reads/writes Excel 2010 `.xlsx`/`.xlsm`/`.xltx`/`.xltm` files


<!-- GETTING STARTED -->
## Getting Started

To get a local copy up and running follow these simple steps.

### Prerequisites

In order to use the Microsoft Planner Comment Parser, you must first have Python and pip installed on your system. If you need assistance installing these prerequisites, see the folowing steps:
* Python is a programming language. All of this project's code base is written in Python. Download the latest version of [Python](https://www.python.org/downloads/) and install onto your local machine.

* Pip is the package installer for Python. Once Python is installed, open your local machine's command line and use the following command to utilize Python to install Pip:
```sh
python get-pip.py -g
```

Git is a version control system. In this project, Git is used to clone (copy) the most up-to-date project files from GitHub to your local machine. Download the latest version of [git](https://git-scm.com/download/win) and install on your local machine.

This project manipulates Microsoft Excel `.xlsx` files, but a local installation of Microsoft Excel is <b>not required</b> to run this script.


### Installation

1. Open the command line on your local machine.

2. Enter the following command to use Git to clone this repository to your local machine.
```sh
git clone https://github.com/asa-holland/microsoft-planner-comment-parser.git
```
3. Enter the following command to use Pip to install this repository's dependencies.
```sh
pip install -r requirements.txt
```



<!-- USAGE EXAMPLES -->
## Usage

To use the Microsoft Planner Comment Parser once installed, add following line item to your project:

```sh
import microsoft-planner-comment-parser as mpcp
```

Locate the `.xlsx` file created by Apps4Pro Planner Manager that contains the exported task information from your desired plan.

Create a local string variable for the path to this saved `.xlsx` file.

Then, add the following line to your script:
```sh
mpcp.commentParser.parse(excel_file='{your path to excel file here}')
```

When run, this will create an additional `.xlsx` file in the same directory as your initial file, and will use the same filename but will append ` Cleaned Comments` to the filename.

## Usage Example with Sample File

To see an example of the Ficrosoft Planner Comment Parser, navigate to the `//test` directory of this project. This folder contains a single `.xlsx` file that contains sample data exported from Microsoft Planner using the Planner Manager.

![Sample Comments Uncleaned][use-file-before]

Next, open the command line, navigate to the installation folder and run:
```sh
python sample_test.py
```

This will create another `.xlsx` file in the same directory where comments have been separated by column and entry.

![Sample Comments Cleaned][use-file-after]

<!-- ROADMAP -->
## Roadmap

See the [open issues](https://github.com/asa-holland/microsoft-planner-comment-parser/issues) for a list of proposed features (and known issues).



<!-- CONTRIBUTING -->
## Contributing

Contributions are what make the open source community such an amazing place to be learn, inspire, and create. Any contributions you make are **greatly appreciated**.

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request



<!-- LICENSE -->
## License

Distributed under the MIT License. See [LICENSE](https://github.com/asa-holland/microsoft-planner-comment-parser/LICENSE.txt) for more information.



<!-- CONTACT -->
## Contact

Asa Holland - [@AsaHolland404](https://twitter.com/AsaHolland404) - hollandasa@gmail.com

Project Link: [https://github.com/asa-holland/microsoft-planner-comment-parser](https://github.com/asa-holland/microsoft-planner-comment-parser)



<!-- ACKNOWLEDGEMENTS -->
## Acknowledgements

* [Apps4Pro](https://apps4.pro/Home.aspx) for doing the heavy lifting in building the Planner Manager.
* [othneildrew](https://github.com/othneildrew) for creating the [template README file](https://github.com/othneildrew/Best-README-Template) that was used as the starting point for the README for this project. 





<!-- MARKDOWN LINKS & IMAGES -->
[linkedin-shield]: https://img.shields.io/badge/-LinkedIn-black.svg?style=flat-square&logo=linkedin&colorB=555
[linkedin-url]: https://www.linkedin.com/in/asa-holland-a2a0b5b7/
[use-file-after]: images/comments-cleaned.JPG
[use-file-before]: images/comments-whole.JPG
[use-example]: images/use.gif