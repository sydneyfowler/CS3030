# Excel Command Line Tool

This is a command line tool used to work with Excel files, written in python. This tool aims to create a navigable interface so that users can quickly clean and manipulate excel files. Functionality includes: 
- Importing / exporting Excel files from and to csv
- Data validation and cleanup features for common cleanup functions (phone numbers, email addresses, state codes, zip codes, dates, web addresses, character limits, etc.)
- Compressing files
- Sending a copy of a file to a list of email addresses
- Removing duplicate rows based on a set of criteria
- Data Analysis: Sum, Count, Max, Min, Unique Values, Average, etc.
- Graphing: Taking data from an excel file and plotting to a chart

## Getting Started

Download the source files or clone to your own project. Once you have installed the prerequisites, simply execute "main.py" to start the tool. The sections below will guide you through the process.

### Prerequisites

Make sure you are executing in python 3 or later. This project was tested and developed in python 3. 

In order for the tool to work properly with all functions, you will the following:
	openpyxl 
	pandas
	matplotlib
	numpy

### Installing

Instead of manually installing the above one at a time, we have already included a file within the project that will allow you to install all the required dependences in one command. This is included in a file called "requirements.txt."

```
$ pip3 install -r requirements.txt
```

Once the dependencies are installed, all you need do is run main.py to start the tool.

```
$ python main.py
```

And it should be working! Follow the instructions to get to a desired sub-tool.

## Running Demos

You will notice a file directory named "test_files." Here there are files that have issues that you can test with this tool. You can choose to select different files for testing throughout the program when prompted.

### Importing a file

When asked for a file, a relative path is all you need to enter. Make sure you include the file type!

```
test_files/example.xlsx
```

## References

* [RuterNPath](https://help.returnpath.com/hc/en-us/articles/220560587-What-are-the-rules-for-email-address-syntax-) - Derived email rules
* [RegexTester](https://www.regextester.com/93652) - Derived web adress rules
* [OpenPyxl](https://openpyxl.readthedocs.io/en/stable/) - Used for opening and manipulating Excel documents in Python
* [Pandas](https://pandas.pydata.org/) - Used for importing and exporting excel files
* [Matplotlib](https://matplotlib.org/) - Used for the graphing / charting tool
* [NumPy](https://numpy.org/) - Libraries for data handling and computation

## Authors

* **Matthew Hileman**
* **Sydney Fowler**

## License

[Licences here]

## Acknowledgments

[Acknowledgements here]
