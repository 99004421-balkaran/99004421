"""Module to generate data for the sheet 5 with random PS numbers
and getting random domain expertise area from the list of area of expertise
 for each PS number"""
from random import choice


class SheetNumberFive:
    """Class for Sheet 5 with functions to generate data and add it to excel sheet"""

    def __init__(self, list_ps_numbers, sheet_obj, workbook_obj):
        self.list_of_ps_numbers = list_ps_numbers
        self.sheet_obj = sheet_obj
        self.workbook_obj = workbook_obj
        self.domain_expertise_list = []
        self.random_domain_expertise = ["Programming", "Design", "Testing", "Architect",
                                        "Project Manager", "Medical", "Communication", "Aerospace",
                                        "Media", "Electronics", "Sheet Metal", "Embedded", "IOT",
                                        "Tourism", "Simulations", "ML", "AI", "Gaming",
                                        "Web Development", "Networking", "Block Chain", "Big data",
                                        "Cloud Computing", "Data Mining", "Image Processing",
                                        "Security", "Bio Medical", "Signal Processing",
                                        "Power Electronics", "VLSI"]

    def get_random_ps_numbers(self):
        """function to get 15 random PS numbers using random module"""
        return self.list_of_ps_numbers

    def generate_random_expertise_domains(self):
        """To generate of random expertise levels list and return it"""
        for _ in range(15):
            row_data = []
            for _ in range(20):
                row_data.append(choice(self.random_domain_expertise))
            self.domain_expertise_list.append(row_data)
        return self.domain_expertise_list

    def write_to_excel_file(self):
        """Function to add data generate to the excel file"""
        self.sheet_obj["A1"] = "PS Number"
        col = 66
        for num in range(1, 21):
            cell_pos = chr(col) + '1'
            self.sheet_obj[cell_pos] = "Domain Area " + str(num)
            col += 1
        index_of_ps = 0
        col = 65
        for row in range(2, 17):
            cell_pos = chr(col) + str(row)
            self.sheet_obj[cell_pos] = self.list_of_ps_numbers[index_of_ps]
            index_of_ps += 1
        self.generate_random_expertise_domains()
        list_col = 0
        list_row = 0
        for row in range(2, 17):
            list_col = 0
            for col in range(66, (66 + 20)):
                cell_pos = chr(col) + str(row)
                self.sheet_obj[cell_pos] = self.domain_expertise_list[list_row][list_col]
                list_col += 1
            list_row += 1
        self.workbook_obj.save("datafile.xlsx")
