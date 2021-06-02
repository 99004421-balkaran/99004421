"""Module to generate data for the sheet 3 with random PS numbers
and getting random cities from the list of cities for each PS number"""
from random import choice


class SheetNumberThree:
    """Class for Sheet 3 with functions to generate data and add it to excel sheet"""

    def __init__(self, list_ps_numbers, sheet_obj, workbook_obj):
        self.list_of_ps_numbers = list_ps_numbers
        self.sheet_obj = sheet_obj
        self.workbook_obj = workbook_obj
        self.cities_list = []
        self.random_cities = ["Kathmandu", "London", "Chicago", "San Jose", "New York", "York",
                              "Jersey", "Delhi", "Douala", "Quito", "Kotkapura", "Faridkot", "Moga",
                              "Amritsar", "Mohali", "Landran", "Kharar", "Jhanjeri", "Dalian",
                              "Dubai", "Portland", "Toronto", "Kingston", "Surrey", "Atlin",
                              "Miami", "Denver", "Boston", "Detroit", "LA", "Vice City", "Caracas",
                              "Shiraz", "Hyderabad", "Calgary", "Tokyo", "Beijing", "Accra"]

    def get_random_ps_numbers(self):
        """function to get 15 random PS numbers using random module"""
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
        self.sheet_obj["A1"] = "PS Number"
        col = 66
        for num in range(1, 21):
            cell_pos = chr(col) + '1'
            self.sheet_obj[cell_pos] = "City " + str(num)
            col += 1
        index_of_ps = 0
        col = 65
        for row in range(2, 17):
            cell_pos = chr(col) + str(row)
            self.sheet_obj[cell_pos] = self.list_of_ps_numbers[index_of_ps]
            index_of_ps += 1
        self.generate_random_cities()
        list_col = 0
        list_row = 0
        for row in range(2, 17):
            list_col = 0
            for col in range(66, (66 + 20)):
                cell_pos = chr(col) + str(row)
                self.sheet_obj[cell_pos] = self.cities_list[list_row][list_col]
                list_col += 1
            list_row += 1
        self.workbook_obj.save("datafile.xlsx")