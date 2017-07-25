import os
import xlrd
from xlrd import open_workbook
from xlutils.copy import copy


def read_file(file_path):
    book = xlrd.open_workbook(file_path)
    return book


class ConvertFile(object):

    def __init__(self, input_file, schema_file, export_file_name):
        self.export_file_name = export_file_name
        self.input_file = input_file
        self.book = read_file(self.input_file)
        self.sheet = self.book.sheet_by_name('Outlook_Contacts')
        self.titles = self.sheet.col_values(0, 1)
        self.first_names = self.sheet.col_values(1, 1)
        self.middle_name = self.sheet.col_values(2, 1)
        self.last_names = self.sheet.col_values(3, 1)
        self.suffixes = self.sheet.col_values(4, 1)
        self.companies = self.sheet.col_values(5, 1)
        self.departments = self.sheet.col_values(6, 1)
        self.job_titles = self.sheet.col_values(7, 1)
        self.business_streets = self.sheet.col_values(8, 1)
        self.business_streets2 = self.sheet.col_values(9, 1)
        self.business_streets3 = self.sheet.col_values(10, 1)
        self.business_cities = self.sheet.col_values(11, 1)
        self.business_states = self.sheet.col_values(12, 1)
        self.business_postals = self.sheet.col_values(13, 1)
        self.business_countries = self.sheet.col_values(14, 1)
        self.home_streets = self.sheet.col_values(15, 1)
        self.home_streets2 = self.sheet.col_values(16, 1)
        self.home_streets3 = self.sheet.col_values(17, 1)
        self.home_cities = self.sheet.col_values(18, 1)
        self.home_states = self.sheet.col_values(19, 1)
        self.home_postals = self.sheet.col_values(20, 1)
        self.home_countries = self.sheet.col_values(21, 1)
        self.other_streets = self.sheet.col_values(22, 1)
        self.other_streets2 = self.sheet.col_values(23, 1)
        self.other_streets3 = self.sheet.col_values(24, 1)
        self.other_cities = self.sheet.col_values(25, 1)
        self.other_states = self.sheet.col_values(26, 1)
        self.other_postals = self.sheet.col_values(27, 1)
        self.other_countries = self.sheet.col_values(28, 1)
        self.assistant_phones = self.sheet.col_values(29, 1)
        self.business_faxes = self.sheet.col_values(30, 1)
        self.business_phones = self.sheet.col_values(31, 1)
        self.business_phones2 = self.sheet.col_values(32, 1)
        self.callbacks = self.sheet.col_values(33, 1)
        self.car_phones = self.sheet.col_values(34, 1)
        self.company_main_phones = self.sheet.col_values(35, 1)
        self.home_faxes = self.sheet.col_values(36, 1)
        self.home_phones = self.sheet.col_values(37, 1)
        self.home_phones2 = self.sheet.col_values(38, 1)
        self.ISDNs = self.sheet.col_values(39, 1)
        self.mobiles = self.sheet.col_values(40, 1)
        self.other_faxes = self.sheet.col_values(41, 1)
        self.other_phones = self.sheet.col_values(42, 1)
        self.pagers = self.sheet.col_values(43, 1)
        self.primary_phones = self.sheet.col_values(44, 1)
        self.radio_phones = self.sheet.col_values(45, 1)
        self.tty_phones = self.sheet.col_values(46, 1)
        self.telex = self.sheet.col_values(47, 1)
        self.accounts = self.sheet.col_values(48, 1)
        self.anniversaries = self.sheet.col_values(49, 1)
        self.assistant_names = self.sheet.col_values(50, 1)
        self.billing = self.sheet.col_values(51, 1)
        self.birthdays = self.sheet.col_values(52, 1)
        self.business_po = self.sheet.col_values(53, 1)
        self.categories = self.sheet.col_values(54, 1)
        self.children = self.sheet.col_values(55, 1)
        self.directories = self.sheet.col_values(56, 1)
        self.emails = self.sheet.col_values(57, 1)
        self.email_types = self.sheet.col_values(58, 1)
        self.email_displays = self.sheet.col_values(59, 1)
        self.emails2 = self.sheet.col_values(60, 1)
        self.emails_types2 = self.sheet.col_values(61, 1)
        self.emails_displays2 = self.sheet.col_values(62, 1)
        self.emails3 = self.sheet.col_values(63, 1)
        self.emails_types3 = self.sheet.col_values(64, 1)
        self.emails_displays3 = self.sheet.col_values(65, 1)
        self.genders = self.sheet.col_values(66, 1)
        self.gov_ids = self.sheet.col_values(67, 1)
        self.hobbies = self.sheet.col_values(68, 1)
        self.home_pos = self.sheet.col_values(69, 1)
        self.initials = self.sheet.col_values(70, 1)
        self.free_busy = self.sheet.col_values(71, 1)
        self.keywords = self.sheet.col_values(72, 1)
        self.languages = self.sheet.col_values(73, 1)
        self.locations = self.sheet.col_values(74, 1)
        self.manager_names = self.sheet.col_values(75, 1)
        self.milages = self.sheet.col_values(76, 1)
        self.notes = self.sheet.col_values(77, 1)
        self.office_locations = self.sheet.col_values(78, 1)
        self.organization_ids = self.sheet.col_values(79, 1)
        self.other_pos = self.sheet.col_values(80, 1)
        self.priorities = self.sheet.col_values(81, 1)
        self.is_private = self.sheet.col_values(82, 1)
        self.professions = self.sheet.col_values(83, 1)
        self.referred = self.sheet.col_values(84, 1)
        self.sensitivities = self.sheet.col_values(85, 1)
        self.spouses = self.sheet.col_values(86, 1)
        self.user = self.sheet.col_values(87, 1)
        self.user2 = self.sheet.col_values(88, 1)
        self.user3 = self.sheet.col_values(89, 1)
        self.user4 = self.sheet.col_values(90, 1)
        self.web_pages = self.sheet.col_values(91, 1)
        self.realhound_path = os.path.join(os.getcwd(), schema_file)
        self.num_rows = self.sheet.nrows
        self.num_columns = self.sheet.ncols

    def write_to(self):
        print('writing')
        realhound = open_workbook(self.realhound_path, on_demand=True)
        writer = copy(realhound)
        realhound_sheet = writer.get_sheet(0)
        record = 0

        for row in range(1, self.num_rows):
            realhound_sheet.write(row, 0, self.companies[record])
            realhound_sheet.write(row, 1, self.last_names[record])
            realhound_sheet.write(row, 2, self.first_names[record])
            realhound_sheet.write(row, 4, 'Mobile')
            realhound_sheet.write(row, 5, self.mobiles[record])
            realhound_sheet.write(row, 6, 'Business Phones')
            realhound_sheet.write(row, 7, self.business_phones[record])
            realhound_sheet.write(row, 8, 'Business Phones2')
            realhound_sheet.write(row, 9, self.business_phones2[record])
            realhound_sheet.write(row, 10, 'Home Phones')
            realhound_sheet.write(row, 11, self.home_phones[record])
            realhound_sheet.write(row, 12, 'Company Phones')
            realhound_sheet.write(row, 13, self.company_main_phones[record])
            realhound_sheet.write(row, 14, self.emails[record])
            realhound_sheet.write(row, 15, self.emails2[record])
            realhound_sheet.write(row, 16, self.business_streets[record])
            realhound_sheet.write(row, 18, self.business_cities[record])
            realhound_sheet.write(row, 19, self.business_states[record])
            realhound_sheet.write(row, 20, self.business_postals[record])
            realhound_sheet.write(row, 22, self.home_streets[record])
            realhound_sheet.write(row, 23, self.home_cities[record])
            realhound_sheet.write(row, 24, self.home_states[record])
            realhound_sheet.write(row, 26, self.home_postals[record])
            realhound_sheet.write(row, 29, self.user[record])
            realhound_sheet.write(row, 30, self.user2[record])
            realhound_sheet.write(row, 30, self.user3[record])
            realhound_sheet.write(row, 31, self.user4[record])
            realhound_sheet.write(row, 38, self.web_pages[record])
            record += 1

        writer.save(os.path.join(os.getcwd(), 'Realhound_filled.xls'))

