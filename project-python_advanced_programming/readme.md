# Excel data Extractor Python Project
These are the main files of the project which will be used in the project. This Project prompts the user for the PS number which he wants to get details of. Then the user chooses the PS number and all details are printed and copied to the output excel file.
The output Excel file will have its name as ***PS Number_output_data.xlsx*** where PS_number will be the PS number selected by the user from the given options.
## Steps to run the project
* To run the project you must have installed python and openpyxl library of python.
* To Install ***openpyxl*** use command ***pip install openpyxl***.
* After installing the openpyxl, run the Python file named ***main_project_file.py***.
* This will show all available PS numbers and ask for choose PS number to get output.
* After checking that input is valid it prints the data of PS number in tabular format and copies that data to the Excel file.
## Steps to run pytest
* In order to run pytest you must have pytest module installed.
* To install pytest using pip type ***pip install pytest***.
* After installing pytest, type pytest in the command line it will collect and run all tests of the project.
## Description of files
|             File name            |                                       Description                                      |
|:--------------------------------:|:--------------------------------------------------------------------------------------:|
|        main_project_file.py      |     File to get PS number input from user to print output and generate   Excel file    |
|          test_project.py         |                       File to test the project using pytest module                     |
|           datafile.xlsx          |                        File containing 5 sheets of required data                       |
|     99004413_output_data.xlsx    |                   Sample output file generated using the project file                  |