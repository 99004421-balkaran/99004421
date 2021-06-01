from random import randint, choice
import openpyxl


class SheetNumberTwo:
    def __init__(self):
        self.list_of_ps_numbers = []
        self.hobbies_list = []
        self.random_hobbies = ["Acting", "Animation", "Art", "Baking", "Basketball", "Beat boxing", "Blogging",
                               "Board games", "Bowling", "Calligraphy", "Candle making", "Candy making",
                               "Card games", "Ceramics", "Chatting""Cheese making", "Chess", "Cleaning", "Collecting",
                               "Coloring", "Communication", "Computer programming", "Confectionery", "Construction",
                               "Cooking", "Craft", "Creative writing", "Crossword puzzles", "Cryptography",
                               "Cue sports", "Dance", "Decorating", "Digital arts", "Dining", "Distro Hopping",
                               "Reading", "Poetry", "Video Games", "Singing", "Surfing", "Gardening"]

    def generate_random_ps_numbers(self):
        for num in range(15):
            self.list_of_ps_numbers.append(randint(99004400, 99004450))
        return self.list_of_ps_numbers

    def generate_random_hobbies(self):
        for row in range(15):
            row_data = []
            for column in range(20):
                row_data.append(choice(self.random_hobbies))
            self.hobbies_list.append(row_data)
        return self.hobbies_list

    def write_to_excel_file(self):
        pass
