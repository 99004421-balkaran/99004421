from random import randint, choice
import openpyxl


class SheetNumberFour:
    def __init__(self):
        self.list_of_ps_numbers = []
        self.programming_expertise_list = []
        self.random_programming_expertise = ["Novice", "Expert", "Advanced", "Intermediate"]
        self.list_of_programming_languages = ["Python", "GO", "Java", "C", "C++", "C#", "Objective C", "Swift",
                                              "Julia", "MATLAB", "Assembly", "MIPS", "Ruby", "R", "Kotlin", "XML",
                                              "CSS", "JavaScript", "BASIC"]

    def generate_random_ps_numbers(self):
        for num in range(15):
            self.list_of_ps_numbers.append(randint(99004400, 99004450))
        return self.list_of_ps_numbers

    def generate_random_expertise_levels(self):
        for row in range(15):
            row_data = []
            for column in range(20):
                row_data.append(choice(self.random_programming_expertise))
            self.programming_expertise_list.append(row_data)
        return self.programming_expertise_list

    def write_to_excel_file(self):
        pass
