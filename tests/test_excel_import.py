import unittest
import os
from openpyxl import Workbook
from wedding_planner.models import SeatingPlan
from wedding_planner.excel_io import ExcelIO

class TestExcelImport(unittest.TestCase):
    def setUp(self):
        self.test_dir = os.path.dirname(os.path.abspath(__file__))
        self.data_dir = os.path.join(self.test_dir, 'data')
        os.makedirs(self.data_dir, exist_ok=True)
        
        self.filename = os.path.join(self.data_dir, "test_groups.xlsx")
        self.hebrew_filename = os.path.join(self.data_dir, "test_hebrew_import.xlsx")
        self.seating_plan = SeatingPlan()
        
        # Create a sample Excel file
        wb = Workbook()
        ws = wb.active
        ws.append(["Group Name", "Count", "Other Column"])
        ws.append(["Family A", 3, "ignore"])
        ws.append(["Couple B", 2, "ignore"])
        ws.append(["Single C", 1, "ignore"])
        wb.save(self.filename)

    def tearDown(self):
        if os.path.exists(self.filename):
            os.remove(self.filename)
        if os.path.exists(self.hebrew_filename):
            os.remove(self.hebrew_filename)

    def test_get_headers(self):
        headers = ExcelIO.get_headers(self.filename)
        self.assertEqual(headers, ["Group Name", "Count", "Other Column"])

    def test_import_groups(self):
        ExcelIO.import_groups_to_plan(self.filename, "Group Name", "Count", self.seating_plan)
        
        # Check total guests: 3 groups -> 3 guest entries (with different sizes)
        # Family A (3), Couple B (2), Single C (1)
        self.assertEqual(len(self.seating_plan.guests), 3)
        
        # Check specific guests
        guests = list(self.seating_plan.guests.values())
        
        # Family A (Size 3)
        family_a = [g for g in guests if g.name == "Family A"]
        self.assertEqual(len(family_a), 1)
        self.assertEqual(family_a[0].size, 3)

    def test_hebrew_import(self):
        # Create dummy Hebrew Excel file
        wb = Workbook()
        ws = wb.active
        ws.title = "Guests"
        
        # Headers
        headers = ["שם", "קטגוריה", "מספר"]
        ws.append(headers)
        
        # Data
        data = [
            ["משפחת כהן", "חברים", 3],
            ["דוד לוי", "משפחה", 1]
        ]
        
        for row in data:
            ws.append(row)
            
        wb.save(self.hebrew_filename)

        ExcelIO.import_groups_to_plan(self.hebrew_filename, group_col=headers[0], count_col=headers[2], category_col=headers[1], seating_plan=self.seating_plan)
        
        # Verify counts and categories and sizes
        # Cohen family (3)
        cohens = [g for g in self.seating_plan.guests.values() if "משפחת כהן" in g.name]
        self.assertEqual(len(cohens), 1)
        self.assertEqual(cohens[0].size, 3)
        self.assertEqual(cohens[0].category, "חברים")

        # David Levy (1)
        levy = [g for g in self.seating_plan.guests.values() if "דוד לוי" in g.name]
        self.assertEqual(len(levy), 1)
        self.assertEqual(levy[0].size, 1)

if __name__ == '__main__':
    unittest.main()
