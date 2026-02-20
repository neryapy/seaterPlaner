import tkinter as tk
from tkinter import ttk

class Styles:
    # --- Modern Palette (Tailwind-inspired) ---
    bg_color = "#f3f4f6"            # Gray-100
    sidebar_bg = "#ffffff"          # White
    
    primary_color = "#4f46e5"       # Indigo-600
    primary_hover = "#4338ca"       # Indigo-700
    secondary_color = "#e5e7eb"     # Gray-200
    secondary_hover = "#d1d5db"     # Gray-300
    
    text_color = "#111827"          # Gray-900
    muted_text_color = "#6b7280"    # Gray-500
    accent_color = "#b91c1c"        # Red-700 (High contrast for full tables)
    
    success_color = "#10b981"       # Emerald-500
    warning_color = "#f59e0b"       # Amber-500
    error_color = "#ef4444"         # Red-500
    
    table_fill_color = "#ffffff"
    table_outline_color = "#9ca3af" # Gray-400
    table_seated_color = "#ecfdf5"  # Emerald-50
    table_full_color = "#fee2e2"    # Red-50

    # --- Fonts ---
    font_family = "Segoe UI" if "win" in __import__('sys').platform else "Inter" 
    # Fallback to Helvetica/Arial if Inter not installed, but let's try for a modern sans.
    
    header_font = (font_family, 18, "bold")
    sub_header_font = (font_family, 14, "bold")
    normal_font = (font_family, 11)
    small_font = (font_family, 9)
    
    @classmethod
    def configure_ttk_styles(cls):
        style = ttk.Style()
        style.theme_use('clam') # 'clam' provides a good base for custom styling on Linux/Windows

        # General
        style.configure(".", background=cls.bg_color, foreground=cls.text_color, font=cls.normal_font)
        
        # Frames
        style.configure("TFrame", background=cls.bg_color)
        style.configure("White.TFrame", background=cls.sidebar_bg)
        style.configure("Card.TFrame", background=cls.sidebar_bg, relief="flat") # We can simulate card with frame

        # Buttons (Primary)
        style.configure("Primary.TButton",
                        background=cls.primary_color,
                        foreground="white",
                        borderwidth=0,
                        focuscolor=cls.primary_color,
                        padding=(15, 8),
                        font=(cls.font_family, 10, "bold"))
        style.map("Primary.TButton",
                  background=[('active', cls.primary_hover), ('pressed', cls.primary_hover)])

        # Buttons (Secondary/Outline)
        style.configure("Secondary.TButton",
                        background="white",
                        foreground=cls.text_color,
                        borderwidth=1,
                        bordercolor=cls.secondary_color,
                        focuscolor=cls.secondary_hover,
                        padding=(12, 6),
                        font=(cls.font_family, 10))
        style.map("Secondary.TButton",
                  background=[('active', "#f9fafb"), ('pressed', "#f3f4f6")],
                  bordercolor=[('active', cls.primary_color)])
                  
        # Treeview
        style.configure("Treeview", 
                        background="white",
                        fieldbackground="white",
                        foreground=cls.text_color,
                        rowheight=45,
                        borderwidth=0,
                        font=cls.normal_font)
        style.configure("Treeview.Heading",
                        background="#f9fafb",
                        foreground=cls.muted_text_color,
                        font=(cls.font_family, 10, "bold"),
                        borderwidth=1,
                        relief="flat")
        style.map("Treeview",
                  background=[('selected', cls.primary_color)],
                  foreground=[('selected', 'white')])
        
        # PanedWindow
        style.configure("TPanedwindow", background=cls.bg_color)
        style.configure("Sash", background=cls.secondary_color, sashthickness=2)

        # Labelframes
        style.configure("TLabelframe", background=cls.bg_color, relief="flat")
        style.configure("TLabelframe.Label", background=cls.bg_color, foreground=cls.text_color, font=cls.sub_header_font)

