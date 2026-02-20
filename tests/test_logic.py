import unittest
from wedding_planner.models import SeatingPlan

class TestSeatingPlan(unittest.TestCase):
    def setUp(self):
        self.plan = SeatingPlan()

    def test_add_guest(self):
        guest = self.plan.add_guest("John Doe", "Friends")
        self.assertEqual(guest.name, "John Doe")
        self.assertEqual(guest.id, 1)
        self.assertIn(1, self.plan.guests)

    def test_add_table(self):
        table = self.plan.add_table("Table 1", 5)
        self.assertEqual(table.name, "Table 1")
        self.assertEqual(table.capacity, 5)
        self.assertIn(1, self.plan.tables)

    def test_assign_guest(self):
        guest = self.plan.add_guest("Alice")
        table = self.plan.add_table("T1", 2)
        
        success = self.plan.assign_guest_to_table(guest.id, table.id)
        self.assertTrue(success)
        self.assertEqual(guest.table_id, table.id)
        self.assertIn(guest.id, table.guest_ids)

    def test_capacity_limit(self):
        g1 = self.plan.add_guest("G1")
        g2 = self.plan.add_guest("G2")
        g3 = self.plan.add_guest("G3")
        table = self.plan.add_table("T1", 2)

        self.plan.assign_guest_to_table(g1.id, table.id)
        self.plan.assign_guest_to_table(g2.id, table.id)
        
        success = self.plan.assign_guest_to_table(g3.id, table.id)
        self.assertFalse(success)
        self.assertIsNone(g3.table_id)
        self.assertEqual(len(table.guest_ids), 2)

    def test_move_guest(self):
        g1 = self.plan.add_guest("G1")
        t1 = self.plan.add_table("T1", 5)
        t2 = self.plan.add_table("T2", 5)

        self.plan.assign_guest_to_table(g1.id, t1.id)
        self.assertIn(g1.id, t1.guest_ids)

        self.plan.assign_guest_to_table(g1.id, t2.id)
        self.assertIn(g1.id, t2.guest_ids)
        self.assertNotIn(g1.id, t1.guest_ids)
        self.assertEqual(g1.table_id, t2.id)

    def test_unseat_guest(self):
        g1 = self.plan.add_guest("G1")
        t1 = self.plan.add_table("T1", 5)
        self.plan.assign_guest_to_table(g1.id, t1.id)
        
        self.plan.unseat_guest(g1.id)
        self.assertIsNone(g1.table_id)
        self.assertNotIn(g1.id, t1.guest_ids)

if __name__ == '__main__':
    unittest.main()
