"""File to generate the data for the Sheet 1 which contains random PS numbers
and random semester marks of 15 rows
"""
from random import randint
# import openpyxl


class SheetNumberOne:
    """Class for Sheet 1 with functions to generate data and add it to excel sheet"""
    def __init__(self):
        self.list_of_ps_numbers = []
        self.semester_marks = []

    def generate_random_ps_numbers(self):
        """function to generate 15 random PS numbers using random module"""
        for _ in range(15):
            self.list_of_ps_numbers.append(randint(99004400, 99004450))
        return self.list_of_ps_numbers

    def generate_random_marks(self):
        """Function to generate random marks to add to the excel file"""
        for _ in range(15):
            row_marks = []
            for _ in range(20):
                row_marks.append(randint(60, 100))
            self.semester_marks.append(row_marks)
        return self.semester_marks

    def write_to_excel_file(self):
        """Function to add data generate to the excel file"""
        pass
