
import sys
import os

# Add the project root to sys.path
sys.path.append(os.getcwd())

try:
    from wedding_planner.styles import Styles
    print(f"Styles.accent_color: {Styles.accent_color}")
    
    # Also verify imports work
    from wedding_planner.gui import WeddingPlannerGUI
    print("GUI imported successfully")
    
except Exception as e:
    print(f"Error: {e}")
    sys.exit(1)
