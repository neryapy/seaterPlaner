import tkinter as tk
from wedding_planner.gui import WeddingPlannerGUI
import sys

def test_gui_launch():
    try:
        root = tk.Tk()
        app = WeddingPlannerGUI(root)
        # Schedule the window to close after 1000ms
        root.after(1000, root.destroy)
        root.mainloop()
        print("GUI launched successfully")
    except Exception as e:
        print(f"GUI launch failed: {e}")
        sys.exit(1)

if __name__ == "__main__":
    test_gui_launch()
