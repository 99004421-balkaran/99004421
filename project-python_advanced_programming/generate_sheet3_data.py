"""Module to generate data for the sheet 3 with random PS numbers
and getting random cities from the list of cities for each PS number"""
from random import randint, choice
import openpyxl


class SheetNumberThree:
    """Class for Sheet 3 with functions to generate data and add it to excel sheet"""
    def __init__(self):
        self.list_of_ps_numbers = []
        self.cities_list = []
        self.random_cities = ["Kathmandu", "London", "Chicago", "San Jose", "New York", "York",
                              "Jersey", "Delhi", "Douala", "Quito", "Kotkapura", "Faridkot", "Moga",
                              "Amritsar", "Mohali", "Landran", "Kharar", "Jhanjeri", "Dalian",
                              "Dubai", "Portland", "Toronto", "Kingston", "Surrey", "Atlin",
                              "Miami", "Denver", "Boston", "Detroit", "LA", "Vice City", "Caracas",
                              "Shiraz", "Hyderabad", "Calgary", "Tokyo", "Beijing", "Accra"]

    def generate_random_ps_numbers(self):
        """Function to generate and return random PS numbers"""
        for _ in range(15):
            self.list_of_ps_numbers.append(randint(99004400, 99004450))
        return self.list_of_ps_numbers

    def generate_random_cities(self):
        """Function to generate and return random cities names"""
        for _ in range(15):
            row_data = []
            for _ in range(20):
                row_data.append(choice(self.random_cities))
            self.cities_list.append(row_data)
        return self.cities_list

    def write_to_excel_file(self):
        """Function to add data generate to the excel file"""
        pass
