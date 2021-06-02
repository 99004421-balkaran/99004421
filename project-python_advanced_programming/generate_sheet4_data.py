"""Module to generate data for the sheet 3 with random PS numbers
and getting random programming expertise from the list of programming languages
for each PS number"""
from random import choice


class SheetNumberFour:
    """Class for Sheet 4 with functions to generate data and add it to excel sheet"""

    def __init__(self, list_ps_numbers, sheet_obj, workbook_obj):
        self.list_of_ps_numbers = list_ps_numbers
        self.sheet_obj = sheet_obj
        self.workbook_obj = workbook_obj
        self.programming_expertise_list = []
        self.random_programming_expertise = ["Novice", "Expert", "Advanced", "Intermediate"]
        self.list_of_programming_languages = ["Python", "GO", "Java", "C", "C++", "C#",
                                              "Swift", "Julia", "MATLAB", "Assembly", "MIPS",
                                              "Ruby", "R", "Kotlin", "XML", "CSS", "JavaScript",
                                              "BASIC", "Objective C", "HTML"]

    def get_random_ps_numbers(self):
        """function to get 15 random PS numbers using random module"""
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
        self.sheet_obj["A1"] = "PS Number"
        col = 66
        for num in range(1, 21):
            cell_pos = chr(col) + '1'
            self.sheet_obj[cell_pos] = self.list_of_programming_languages[num-1]
            col += 1
        index_of_ps = 0
        col = 65
        for row in range(2, 17):
            cell_pos = chr(col) + str(row)
            self.sheet_obj[cell_pos] = self.list_of_ps_numbers[index_of_ps]
            index_of_ps += 1
        self.generate_random_expertise_levels()
        list_col = 0
        list_row = 0
        for row in range(2, 17):
            list_col = 0
            for col in range(66, (66 + 20)):
                cell_pos = chr(col) + str(row)
                self.sheet_obj[cell_pos] = self.programming_expertise_list[list_row][list_col]
                list_col += 1
            list_row += 1
        self.workbook_obj.save("datafile.xlsx")

