from random import randint, choice
import openpyxl


class SheetNumberFive:
    def __init__(self):
        self.list_of_ps_numbers = []
        self.domain_expertise_list = []
        self.random_domain_expertise = ["Programming", "Design", "Testing", "Architect", "Project Manager",
                                        "Medical", "Communication", "Aerospace", "Media", "Electronics",
                                        "Sheet Metal", "Embedded", "IOT", "Tourism", "Simulations", "ML", "AI",
                                        "Gaming", "Web Development", "Networking", "Block Chain", "Big data",
                                        "Cloud Computing", "Data Mining", "Image Processing", "Security",
                                        "Bio Medical", "Signal Processing", "Power Electronics", "VLSI"]

    def generate_random_ps_numbers(self):
        for num in range(15):
            self.list_of_ps_numbers.append(randint(99004400, 99004450))
        return self.list_of_ps_numbers

    def generate_random_expertise_levels(self):
        for row in range(15):
            row_data = []
            for column in range(20):
                row_data.append(choice(self.random_domain_expertise))
            self.domain_expertise_list.append(row_data)
        return self.domain_expertise_list

    def write_to_excel_file(self):
        pass
