import re
import main
import unittest
from openpyxl import load_workbook, Workbook
from openpyxl.cell.read_only import ReadOnlyCell


class TestBaseFunctionality(unittest.TestCase):
    def test_root_pattern_regex(self):
        test_strings = {"'Ace-SP&L'!C20*C$56": 3,
                        "C45/R34": 3, "=365*SUM('Ace-SP&L'!C20:C$56)": 5}
        i = 0
        for key, value in test_strings.items():
            with self.subTest(i=i):
                spilt_list = list(
                    filter(None, re.split(main.root_pattern_regex, key)))
                self.assertEqual(len(spilt_list), value)
            i += 1

    def test_extract_regex(self):
        test_strings = {
            "'Ace-SP&L'!C260": ["Ace-SP&L", "260"], "Q456": ["456"]}
        i = 0
        for key, value in test_strings.items():
            with self.subTest(i=i):
                split_list = list(
                    filter(None, re.split(main.extract_regex, key)))
                self.assertEqual(split_list, value)

    def test_formula_cleaner(self):
        self.assertEqual(main.clean_formula(["="]), [])
        self.assertEqual(main.clean_formula(["=SUM("]), ["SUM("])
        self.assertEqual(main.clean_formula(
            ["=365*AVERAGE("]), ["365*AVERAGE("])
        self.assertEqual(main.clean_formula(["=("]), ["("])

    def test_dict_generator(self):
        root_dict = {}
        # Open the given file to parse
        read_book = load_workbook('NF-SA Template 160519.xlsx', read_only=True)

        # List of helper sheets needs to be parsed
        sheets = ['Ace-SP&L', 'Ace-SBS', 'Ace-SCFS']
        # generate the dict for lookup of row names
        main.generate_dict(read_book, root_dict, ['SA-Ratios'], 2, 2, 7)
        main.generate_dict(read_book, root_dict, sheets, 1, 1, 4)

        self.assertEqual(len(root_dict), 4)


if __name__ == "__main__":
    unittest.main()
