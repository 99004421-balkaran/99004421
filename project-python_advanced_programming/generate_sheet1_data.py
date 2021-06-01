from random import randint
import openpyxl


class SheetNumberOne:
    def __init__(self):
        self.list_of_ps_numbers = []
        self.semester_marks = []

    def generate_random_ps_numbers(self):
        for num in range(15):
            self.list_of_ps_numbers.append(randint(99004400, 99004450))
        return self.list_of_ps_numbers

    def generate_random_marks(self):
        for row in range(15):
            row_marks = []
            for column in range(20):
                row_marks.append(randint(60, 100))
            self.semester_marks.append(row_marks)
        return self.semester_marks

    def write_to_excel_file(self):
        pass
