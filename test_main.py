import unittest
import main
from openpyxl import load_workbook, Workbook


class TestMainFunctionality(unittest.TestCase):
    # open the book and generate the dictionary
    def setUp(self):
        # List of helper sheets needs to be parsed
        sheets = ['Ace-SP&L', 'Ace-SBS', 'Ace-SCFS']
        # Root dict to hold the parsed data
        self.root_dict = {}
        # Open the given file to parse
        self.read_book = load_workbook(
            'NF-SA Template 160519.xlsx', read_only=True)
        # generate the dict for lookup of row names
        main.generate_dict(self.read_book, self.root_dict,
                           ['SA-Ratios'], 2, 2, 7)
        main.generate_dict(self.read_book, self.root_dict, sheets, 1, 1, 4)

    # Unit Test to test the root functionality of the solution, to generate the textual representation of the given formula
    def test_formatter(self):
        self.assertEqual(main.format_formula(
            "='Ace-SP&L'!C20*C450", self.root_dict), ["Trading Sales", "*", "Ratio"])
        self.assertEqual(main.format_formula("=SUM(C868:C872)", self.root_dict), [
                         "SUM(", "Lease Adjustment A/c", ":", "Other Non Current Assets", ")"])
        self.assertEqual(main.format_formula("=SUM(D810:D812)-SUM(C810:C812)", self.root_dict), [
                         "SUM(", "Total GB", ":", "Capital Advances", ")-SUM(", "Total GB", ":", "Capital Advances", ")"])
        # Net proceeds/(repayment) from debt
        self.assertEqual(main.format_formula(
            "=SUM('Ace-SCFS'!C58:C63)+'Ace-SCFS'!C73+'Ace-SCFS'!C75", self.root_dict),
            ["SUM(", "Increase / (Decrease) in Loan Funds", ":", "Bad and doubtful debts as % of revenues", ")+",
             "Changes in working capital borrowings", "+", "Net Inc/Dec in cash / Export credit facilities and other short term loans"])
        # 10 year CAGR
        self.assertTrue(main.format_formula("=(S69/I69)^(1/10)-1", self.root_dict),
                        ["(", "Net Sales", "/", "Net Sales", ")^(1/10)-1"])
        self.assertTrue(main.format_formula("=365*AVERAGE(D930)/C929",
                                            self.root_dict), ['365*AVERAGE(', 'Inventories', ')/', 'COGS'])

    # Unit Test to test if the function to find the textual representation of a given coordinate is returned properly
    def test_find_name(self):
        self.assertEqual(main.find_name("D42", self.root_dict), "Net Sales")
        # ref to another sheet
        self.assertEqual(main.find_name(
            "C$947", self.root_dict), "Short Term Provisions")
        # 2 level nested example
        self.assertEqual(main.find_name(
            "C478", self.root_dict), "Total Asset turnover")
        # fail case, 457 is not an indexable row
        self.assertEqual(main.find_name("C457", self.root_dict), "457")

        self.assertEqual(main.find_name("C957", self.root_dict),
                         "Total Equity and Liabilities")
        self.assertEqual(main.find_name("C97", self.root_dict),
                         "General and Administration Expenses")


if __name__ == "__main__":
    unittest.main()
