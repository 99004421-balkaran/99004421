# 99004421 Python Advanced Programming Project
This is the project created by me for the Python Advanced Programming Module of the genesis. This project uses excel read and write operations to store random data created by the program. This data can is stored in the excel file using ***openpyxl*** library. First, the random data is created using the python file in the folder ***Random_Data_Generator*** then this data is read by another python file and data is extracted from the excel file to print on console and copying that data to new Excel output file.
## Objectives of the project
* User can select the PS number of the employee he desires.
* The user input will be checked and data will be extarcted from the excel file
* The extracted data will be printed on the console and written to the excel file.
* Then name of the excel file is ***PS_number_output_data.xlsx*** where PS_number is the number selected by the user.
* User can Access this file for any future use.
## Pytest and Pylint Working
 [![Pytest - Unit Testing](https://github.com/99004421-balkaran/99004421/actions/workflows/pytest.yml/badge.svg)](https://github.com/99004421-balkaran/99004421/actions/workflows/pytest.yml) 

 [![Pylint Check](https://github.com/99004421-balkaran/99004421/actions/workflows/pylint.yml/badge.svg)](https://github.com/99004421-balkaran/99004421/actions/workflows/pylint.yml)
 
 [![Git Inspector](https://github.com/99004421-balkaran/99004421/actions/workflows/gitInspector.yml/badge.svg)](https://github.com/99004421-balkaran/99004421/actions/workflows/gitInspector.yml)
## Folder Structure
|                 Folder Name                |                                         Description                                       |
|:------------------------------------------:|:-----------------------------------------------------------------------------------------:|
|     project-python_advanced_programming    |       Folder Containing all python files   including main project file and test file      |
|            Screenshots_Of_Project          |     Screenshots of the pytest, pylint,   running project and outputs shown to the user    |
## Steps to run the project
* To run the project you must have installed python and openpyxl library of python.
* To Install ***openpyxl*** use command ***pip install openpyxl***.
* After installing the openpyxl, run the Python file named ***main_project_file.py*** by opening the ***project-python_advanced_programming*** folder of the project.
* This will show all available PS numbers and ask for choose PS number to get output.
* After checking that input is valid it prints the data of PS number in tabular format and copies that data to the Excel file.
## Steps to run pytest
* In order to run pytest you must have pytest module installed.
* To install pytest using pip type ***pip install pytest***.
* After installing pytest, Open the ***project-python_advanced_programming*** folder of the project.
* Type ***pytest*** in the command line it will collect and run all tests of the project.
