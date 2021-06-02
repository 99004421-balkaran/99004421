"""File to generate the data for the Sheet 1 which contains random PS numbers
and random semester marks of 15 rows
"""
from random import randint


class SheetNumberOne:
    """Class for Sheet 1 with functions to generate data and add it to excel sheet"""

    def __init__(self, list_ps_numbers, sheet_obj, workbook_obj):
        self.list_of_ps_numbers = list_ps_numbers
        self.semester_marks = []
        self.sheet_obj = sheet_obj
        self.workbook_obj = workbook_obj

    def get_random_ps_numbers(self):
        """function to get 15 random PS numbers using random module"""
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
        self.sheet_obj["A1"] = "PS Number"
        col = 66
        for num in range(1, 21):
            cell_pos = chr(col) + '1'
            self.sheet_obj[cell_pos] = "Subject " + str(num)
            col += 1
        index_of_ps = 0
        col = 65
        for row in range(2, 17):
            cell_pos = chr(col) + str(row)
            self.sheet_obj[cell_pos] = self.list_of_ps_numbers[index_of_ps]
            index_of_ps += 1
        self.generate_random_marks()
        list_col = 0
        list_row = 0
        for row in range(2, 17):
            list_col = 0
            for col in range(66, (66 + 20)):
                cell_pos = chr(col) + str(row)
                self.sheet_obj[cell_pos] = self.semester_marks[list_row][list_col]
                list_col += 1
            list_row += 1
        self.workbook_obj.save("datafile.xlsx")
