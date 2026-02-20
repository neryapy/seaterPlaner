import unittest
import tkinter as tk
from tkinter import ttk
import sys
import os

# Ensure the current directory is in the python path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from wedding_planner.gui import WeddingPlannerGUI
from wedding_planner.models import SeatingPlan

class TestGUIIntegration(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        # Hide window during tests
        self.root.withdraw() 
        self.app = WeddingPlannerGUI(self.root)
        self.plan = self.app.seating_plan
        # Clear defaults
        self.plan.guests.clear()
        self.plan.tables.clear()

    def tearDown(self):
        self.root.destroy()

    def test_guest_list_structure(self):
        # 1. Verify Treeview existence
        self.assertTrue(hasattr(self.app, 'guest_tree'), "app.guest_tree does not exist.")
        self.assertIsInstance(self.app.guest_tree, ttk.Treeview, "app.guest_tree is not a Treeview")
            
        # 2. Verify Columns
        columns = self.app.guest_tree["columns"]
        expected_columns = ("name", "category", "size")
        self.assertEqual(tuple(columns), expected_columns, f"Unexpected columns. Got {columns}")

    def test_guest_list_population(self):
        # 3. Verify Data Population
        g1 = self.plan.add_guest("Family Cohen", "Friends", size=3)
        g2 = self.plan.add_guest("David", "Family", size=1)
        
        self.app.refresh_guest_list()
        
        children = self.app.guest_tree.get_children()
        self.assertEqual(len(children), 2, "Expected 2 items in tree")

        # Check Item Values
        item1 = self.app.guest_tree.item(str(g1.id))
        values1 = item1['values']
        
        # Values come back as strings/ints depending on insertion, usually converted to str by tk
        self.assertIn("Family Cohen", str(values1[0]))
        self.assertEqual(str(values1[2]), "3")

    def test_stats_update(self):
        # 4. Verify Stats Logic
        g1 = self.plan.add_guest("Family Cohen", "Friends", size=3)
        self.app.update_stats()
        
        # Guests string format: "Guests: 0/3 Seated" (or similar structure)
        stats_text = self.app.stats_label.cget("text")
        self.assertIn("Guests: 0/3 Seated", stats_text)

    def test_sort_persistence(self):
        # Add guests
        g1 = self.plan.add_guest("C", "General", 3)
        g2 = self.plan.add_guest("A", "General", 1)
        g3 = self.plan.add_guest("B", "General", 2)
        
        self.app.refresh_guest_list()
        
        # Default sort is Name ASC (A, B, C)
        children = self.app.guest_tree.get_children()
        names = [self.app.guest_tree.item(c)['values'][0] for c in children]
        self.assertEqual(names, ["A", "B", "C"])
        
        # Simulate clicking Size column header -> Size ASC
        self.app.treeview_sort_column(self.app.guest_tree, "size", False)
        
        children = self.app.guest_tree.get_children()
        sizes = [int(self.app.guest_tree.item(c)['values'][2]) for c in children]
        self.assertEqual(sizes, [1, 2, 3])
        
        # Reverse -> Size DESC
        self.app.treeview_sort_column(self.app.guest_tree, "size", True)
        children = self.app.guest_tree.get_children()
        sizes = [int(self.app.guest_tree.item(c)['values'][2]) for c in children]
        self.assertEqual(sizes, [3, 2, 1])

    def test_capacity_enforcement(self):
        # Create Table with capacity 2
        t1 = self.plan.add_table("Small Table", 2)
        
        # Create Guest with size 3
        g1 = self.plan.add_guest("Big Group", "General", 3)
        
        # Try to assign via logic
        success = self.plan.assign_guest_to_table(g1.id, t1.id)
        self.assertFalse(success, "Logic allowed assignment exceeding capacity.")
        self.assertIsNone(g1.table_id)
        
        # Create Guest with size 2
        g2 = self.plan.add_guest("Fits", "General", 2)
        success = self.plan.assign_guest_to_table(g2.id, t1.id)
        self.assertTrue(success, "Logic prevented valid assignment.")

if __name__ == "__main__":
    unittest.main()
