"""File to test all functions of the Project one by one"""
import openpyxl

from main_project_file import get_cities_ps_number, get_marks_ps_number, \
    get_hobbies_ps_number, get_programming_ps_number, get_domain_areas_ps_number, \
    get_ps_numbers, write_marks_to_excel, write_hobbies_to_excel, \
    write_domain_areas_to_excel, write_cities_visited_to_excel, write_programming_skills_to_excel

input_workbook = openpyxl.open("datafile.xlsx")
output_workbook = openpyxl.Workbook()
output_workbook.create_sheet("Output Sheet")
sheet = output_workbook["Output Sheet"]


def test_get_ps_numbers():
    """Function to test PS numbers function"""
    assert isinstance(get_ps_numbers(input_workbook), list)
    assert isinstance(get_ps_numbers(input_workbook), list)


def test_marks_function():
    """Function to test get semester marks list function"""
    assert isinstance((get_marks_ps_number(6, input_workbook)), list)
    assert isinstance((get_marks_ps_number(7, input_workbook)), list)


def test_hobbies_function():
    """Function to test get hobbies list function"""
    assert isinstance((get_hobbies_ps_number(6, input_workbook)), list)
    assert isinstance((get_hobbies_ps_number(7, input_workbook)), list)


def test_cities_function():
    """Function to test get cities list function"""
    assert isinstance((get_cities_ps_number(6, input_workbook)), list)
    assert isinstance((get_cities_ps_number(7, input_workbook)), list)


def test_programming_function():
    """Function to test get programming skills list function"""
    assert isinstance((get_programming_ps_number(6, input_workbook)), list)
    assert isinstance((get_programming_ps_number(7, input_workbook)), list)


def test_domain_function():
    """Function to test get domain areas list function"""
    assert isinstance((get_domain_areas_ps_number(6, input_workbook)), list)
    assert isinstance((get_domain_areas_ps_number(7, input_workbook)), list)


def test_write_marks_to_excel():
    """Function to test function for writing marks to the excel file"""
    marks_list_1 = get_marks_ps_number(6, input_workbook)
    marks_list_2 = get_marks_ps_number(10, input_workbook)
    assert write_marks_to_excel(sheet, "A", marks_list_1) is True
    assert write_marks_to_excel(sheet, "A", marks_list_2) is True


def test_write_hobbies_to_excel():
    """Function to test function for writing hobbies to the excel file"""
    hobbies_list_1 = get_hobbies_ps_number(6, input_workbook)
    hobbies_list_2 = get_hobbies_ps_number(10, input_workbook)
    assert write_hobbies_to_excel(sheet, "B", hobbies_list_1) is True
    assert write_hobbies_to_excel(sheet, "B", hobbies_list_2) is True


def test_write_cities_to_excel():
    """Function to test function for writing cities to the excel file"""
    cities_list_1 = get_cities_ps_number(6, input_workbook)
    cities_list_2 = get_cities_ps_number(10, input_workbook)
    assert write_cities_visited_to_excel(sheet, "C", cities_list_1) is True
    assert write_cities_visited_to_excel(sheet, "C", cities_list_2) is True


def test_write_programming_to_excel():
    """Function to test function for writing programming skills
     to the excel file"""
    programming_list_1 = get_programming_ps_number(6, input_workbook)
    programming_list_2 = get_programming_ps_number(10, input_workbook)
    assert write_programming_skills_to_excel(sheet, "D", programming_list_1) is True
    assert write_programming_skills_to_excel(sheet, "D", programming_list_2) is True


def test_write_domain_to_excel():
    """Function to test function for writing domain areas to the excel file"""
    domain_list_1 = get_domain_areas_ps_number(6, input_workbook)
    domain_list_2 = get_domain_areas_ps_number(10, input_workbook)
    assert write_domain_areas_to_excel(sheet, "E", domain_list_1) is True
    assert write_domain_areas_to_excel(sheet, "E", domain_list_2) is True
