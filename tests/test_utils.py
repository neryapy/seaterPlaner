import unittest
import arabic_reshaper
from bidi.algorithm import get_display

def fix_text(text):
    if not text: return text
    reshaped_text = arabic_reshaper.reshape(str(text))
    return get_display(reshaped_text)

class TestUtils(unittest.TestCase):
    def test_rtl_logic(self):
        original = "שלום"
        fixed = fix_text(original)
        
        # Expected visual order for "שלום" when reversed for LTR rendering logic
        # visually: Mem Sofit, Vav, Lamed, Shin
        expected_visual = "םולש" 
        
        self.assertEqual(fixed, expected_visual, f"Expected {expected_visual}, got {fixed}")

if __name__ == '__main__':
    unittest.main()
