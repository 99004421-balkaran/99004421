"""Module to generate data for the sheet 3 with random PS numbers
and getting random programming expertise from the list of programming languages
for each PS number"""
from random import randint, choice
import openpyxl


class SheetNumberFour:
    """Class for Sheet 4 with functions to generate data and add it to excel sheet"""
    def __init__(self):
        self.list_of_ps_numbers = []
        self.programming_expertise_list = []
        self.random_programming_expertise = ["Novice", "Expert", "Advanced", "Intermediate"]
        self.list_of_programming_languages = ["Python", "GO", "Java", "C", "C++", "C#",
                                              "Swift", "Julia", "MATLAB", "Assembly", "MIPS",
                                              "Ruby","R", "Kotlin", "XML", "CSS", "JavaScript",
                                              "BASIC","Objective C"]

    def generate_random_ps_numbers(self):
        """function to generate 15 random PS numbers using random module"""
        for _ in range(15):
            self.list_of_ps_numbers.append(randint(99004400, 99004450))
        return self.list_of_ps_numbers

    def generate_random_expertise_levels(self):
        """Function to generate random expertise levels for the programming languages"""
        for _ in range(15):
            row_data = []
            for _ in range(20):
                row_data.append(choice(self.random_programming_expertise))
            self.programming_expertise_list.append(row_data)
        return self.programming_expertise_list

    def write_to_excel_file(self):
        """Function to add data generate to the excel file"""
        pass
