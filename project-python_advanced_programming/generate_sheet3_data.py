from random import randint, choice
import openpyxl


class SheetNumberThree:
    def __init__(self):
        self.list_of_ps_numbers = []
        self.cities_list = []
        self.random_cities = ["Kathmandu", "London", "Chicago", "San Jose", "New York", "York", "Jersey", "Delhi",
                              "Douala", "Quito", "Kotkapura", "Faridkot", "Moga", "Amritsar", "Mohali", "Landran",
                              "Kharar", "Jhanjeri", "Dalian", "Dubai", "Portland", "Toronto", "Kingston", "Surrey",
                              "Atlin", "Miami", "Denver", "Boston", "Detroit", "LA", "Vice City", "Caracas",
                              "Shiraz", "Hyderabad", "Calgary", "Tokyo", "Beijing", "Accra"]

    def generate_random_ps_numbers(self):
        for num in range(15):
            self.list_of_ps_numbers.append(randint(99004400, 99004450))
        return self.list_of_ps_numbers

    def generate_random_cities(self):
        for row in range(15):
            row_data = []
            for column in range(20):
                row_data.append(choice(self.random_cities))
            self.cities_list.append(row_data)
        return self.cities_list

    def write_to_excel_file(self):
        pass
