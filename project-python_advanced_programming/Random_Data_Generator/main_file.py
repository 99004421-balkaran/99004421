"""This is the main file which generates the random data for the excel sheets
This data will be read by the other python file for the output
"""
from random import randint

import openpyxl
from generate_sheet1_data import SheetNumberOne
from generate_sheet2_data import SheetNumberTwo
from generate_sheet3_data import SheetNumberThree
from generate_sheet4_data import SheetNumberFour
from generate_sheet5_data import SheetNumberFive


def create_workbook():
    """Function to create the workbook with 5 sheets as required"""
    workbook = openpyxl.Workbook()
    sheet_names = ['Semester Marks Sheet', 'Random Hobbies Sheet', 'Cities Visited Sheet',
                   'Programming Expertise Sheet', 'Domain Areas Sheet']
    for sheet_name in sheet_names:
        workbook.create_sheet(sheet_name)
    workbook.remove(workbook['Sheet'])
    workbook.save("datafile.xlsx")
    return workbook


def generate_random_ps_numbers():
    """function to generate 15 random PS numbers using random module"""
    list_of_ps_numbers = []
    for _ in range(16):
        list_of_ps_numbers.append(randint(99004400, 99004450))
    return list_of_ps_numbers


if __name__ == '__main__':
    main_workbook = create_workbook()
    sheet_1 = main_workbook['Semester Marks Sheet']
    ps_numbers_list = generate_random_ps_numbers()
    sheet_1_class = SheetNumberOne(ps_numbers_list, sheet_1, main_workbook)
    sheet_1_class.write_to_excel_file()

    sheet_2 = main_workbook['Random Hobbies Sheet']
    sheet_2_class = SheetNumberTwo(ps_numbers_list, sheet_2, main_workbook)
    sheet_2_class.write_to_excel_file()

    sheet_3 = main_workbook['Cities Visited Sheet']
    sheet_3_class = SheetNumberThree(ps_numbers_list, sheet_3, main_workbook)
    sheet_3_class.write_to_excel_file()

    sheet_4 = main_workbook['Programming Expertise Sheet']
    sheet_4_class = SheetNumberFour(ps_numbers_list, sheet_4, main_workbook)
    sheet_4_class.write_to_excel_file()

    sheet_5 = main_workbook['Domain Areas Sheet']
    sheet_5_class = SheetNumberFive(ps_numbers_list, sheet_5, main_workbook)
    sheet_5_class.write_to_excel_file()
