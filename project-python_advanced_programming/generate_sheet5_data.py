"""Module to generate data for the sheet 5 with random PS numbers
and getting random domain expertise area from the list of area of expertise
 for each PS number"""
from random import randint, choice
import openpyxl


class SheetNumberFive:
    """Class for Sheet 5 with functions to generate data and add it to excel sheet"""
    def __init__(self):
        self.list_of_ps_numbers = []
        self.domain_expertise_list = []
        self.random_domain_expertise = ["Programming", "Design", "Testing", "Architect",
                                        "Project Manager", "Medical", "Communication", "Aerospace",
                                        "Media", "Electronics", "Sheet Metal", "Embedded", "IOT",
                                        "Tourism", "Simulations", "ML", "AI", "Gaming",
                                        "Web Development", "Networking", "Block Chain", "Big data",
                                        "Cloud Computing", "Data Mining", "Image Processing",
                                        "Security", "Bio Medical", "Signal Processing",
                                        "Power Electronics", "VLSI"]

    def generate_random_ps_numbers(self):
        """function to generate 15 random PS numbers using random module"""
        for _ in range(15):
            self.list_of_ps_numbers.append(randint(99004400, 99004450))
        return self.list_of_ps_numbers

    def generate_random_expertise_levels(self):
        """To generate of random expertise levels list and return it"""
        for _ in range(15):
            row_data = []
            for _ in range(20):
                row_data.append(choice(self.random_domain_expertise))
            self.domain_expertise_list.append(row_data)
        return self.domain_expertise_list

    def write_to_excel_file(self):
        """Function to add data generate to the excel file"""
        pass
