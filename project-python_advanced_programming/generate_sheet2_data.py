"""Module to generate data for the sheet 2 with random PS numbers
and getting random hobbies from the list of hobbies for each PS number"""
from random import randint, choice
import openpyxl


class SheetNumberTwo:
    """Class for Sheet 2 with functions to generate data and add it to excel sheet"""

    def __init__(self):
        self.list_of_ps_numbers = []
        self.hobbies_list = []
        self.random_hobbies = ["Acting", "Animation", "Art", "Baking", "Basketball", "Beat boxing",
                               "Blogging", "Bowling", "Calligraphy", "Candle making",
                               "Candy making", "Card games", "Ceramics", "Chatting", "Cheese making",
                               "Chess", "Cleaning", "Collecting", "Coloring", "Communication",
                               "Computer programming", "Confectionery", "Construction",
                               "Cooking", "Craft", "Creative writing", "Crossword puzzles",
                               "Cryptography", "Cue sports", "Dance", "Decorating", "Digital arts",
                               "Dining", "Distro Hopping", "Reading", "Poetry", "Video Games",
                               "Singing", "Surfing", "Gardening"]

    def generate_random_ps_numbers(self):
        """function to generate 15 random PS numbers using random module"""
        for _ in range(15):
            self.list_of_ps_numbers.append(randint(99004400, 99004450))
        return self.list_of_ps_numbers

    def generate_random_hobbies(self):
        """Function to generate random hobbies for the PS numbers"""
        for _ in range(15):
            row_data = []
            for _ in range(20):
                row_data.append(choice(self.random_hobbies))
            self.hobbies_list.append(row_data)
        return self.hobbies_list

    def write_to_excel_file(self):
        """Function to add data generate to the excel file"""
        pass
