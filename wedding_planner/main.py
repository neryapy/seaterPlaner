import tkinter as tk
from wedding_planner.gui import WeddingPlannerGUI

def main():
    root = tk.Tk()
    app = WeddingPlannerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
