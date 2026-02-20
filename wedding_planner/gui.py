from .exporter import GoogleSheetsExporter
from .excel_io import ExcelIO
import arabic_reshaper
from bidi.algorithm import get_display
import os
import tkinter as tk
from tkinter import ttk, simpledialog, messagebox, filedialog
from .models import SeatingPlan, Guest, Table
from .styles import Styles

def fix_text(text):
    if not text: return text
    reshaped_text = arabic_reshaper.reshape(str(text))
    return get_display(reshaped_text)

class RTLStringDialog(simpledialog.Dialog):
    def __init__(self, parent, title, prompt, initialvalue=None):
        self.prompt = prompt
        self.initialvalue = initialvalue
        self.result = None
        super().__init__(parent, title)

    def body(self, master):
        ttk.Label(master, text=self.prompt, font=Styles.normal_font).pack(pady=(10, 5))
        
        self.var = tk.StringVar()
        if self.initialvalue:
            self.var.set(self.initialvalue)
            
        self.entry = ttk.Entry(master, textvariable=self.var, font=Styles.normal_font, justify="right")
        self.entry.pack(padx=20, pady=5, fill=tk.X)
        
        self.preview_label = ttk.Label(master, text="", font=Styles.normal_font, foreground=Styles.primary_color)
        self.preview_label.pack(pady=(0, 10))
        
        self.var.trace("w", self.update_preview)
        self.update_preview() # Init
        
        return self.entry

    def update_preview(self, *args):
        text = self.var.get()
        self.preview_label.config(text=fix_text(text))

    def apply(self):
        self.result = self.entry.get()
    
    def buttonbox(self):
        '''add standard button box. override to style buttons'''
        box = ttk.Frame(self)

        w = ttk.Button(box, text="OK", width=10, command=self.ok, style="Primary.TButton", default=tk.ACTIVE)
        w.pack(side=tk.LEFT, padx=5, pady=5)
        w = ttk.Button(box, text="Cancel", width=10, command=self.cancel, style="Secondary.TButton")
        w.pack(side=tk.LEFT, padx=5, pady=5)

        self.bind("<Return>", self.ok)
        self.bind("<Escape>", self.cancel)

        box.pack()

class WeddingPlannerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Wedding Seating Planner")
        self.root.geometry("1000x800")
        self.root.configure(bg=Styles.bg_color)

        self.seating_plan = SeatingPlan()
        
        # Drag and drop state
        self.drag_data = {"item": None, "x": 0, "y": 0, "type": None}
        
        # Sort state
        self.sort_col = "name"
        self.sort_reverse = False

        self.setup_ui()

    def setup_ui(self):
        self.fix_text = fix_text
        
        # Apply TTK Styles
        Styles.configure_ttk_styles()

        # --- Top Toolbar (Modern White Bar) ---
        toolbar = ttk.Frame(self.root, style="White.TFrame", padding=(20, 10))
        toolbar.pack(side=tk.TOP, fill=tk.X)
        
        # Shadow / Separator effect (using a thin frame)
        ttk.Frame(self.root, style="TFrame", height=1).pack(side=tk.TOP, fill=tk.X) # Placeholder for shadow if needed, or just let color diff handle it
        
        # Title/Logo Area (Left)
        title_label = ttk.Label(toolbar, text="Seater Planner", font=Styles.header_font, background=Styles.sidebar_bg, foreground=Styles.primary_color)
        title_label.pack(side=tk.LEFT, padx=(0, 30))

        def create_btn(parent, text, cmd, style="Primary.TButton"):
            btn = ttk.Button(parent, text=text, command=cmd, style=style, cursor="hand2")
            btn.pack(side=tk.LEFT, padx=5)
            return btn

        create_btn(toolbar, "+ Guest", self.add_guest_dialog)
        create_btn(toolbar, "+ Table", self.add_table_dialog)
        
        # Spacer
        ttk.Frame(toolbar, style="White.TFrame").pack(side=tk.LEFT, expand=True, fill=tk.X)
        
        # Right side actions
        def create_btn_right(parent, text, cmd, style="Secondary.TButton"):
             btn = ttk.Button(parent, text=text, command=cmd, style=style, cursor="hand2")
             btn.pack(side=tk.RIGHT, padx=5)

        self.default_table_capacity = 12
        self.auto_use_default_capacity = True

        create_btn_right(toolbar, "Settings", self.settings_dialog)
        create_btn_right(toolbar, "Import Groups", self.import_groups_dialog)
        create_btn_right(toolbar, "Load XLSX", self.load_excel)
        create_btn_right(toolbar, "Save XLSX", self.save_excel)


        # --- Statistics Bar (Floating or Subtle) ---
        self.stats_frame = ttk.Frame(self.root, style="TFrame", padding=(20, 5))
        self.stats_frame.pack(side=tk.TOP, fill=tk.X)
        
        self.stats_label = ttk.Label(self.stats_frame, text="Ready", font=Styles.normal_font, foreground=Styles.muted_text_color)
        self.stats_label.pack(side=tk.LEFT)

        # --- Main Layout ---
        paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))

        # Left Panel: Guest Chips (Canvas based)
        left_frame = ttk.Frame(paned_window, style="White.TFrame", width=320)
        # left_frame.pack_propagate(False) # ttk frame doesn't support pack_propagate nicely sometimes, but let's try
        paned_window.add(left_frame, weight=0) # Fixed width
        
        ttk.Label(left_frame, text="Guest List", font=Styles.header_font, background=Styles.sidebar_bg).pack(pady=(20, 10), padx=15, anchor="w")
        
        # Search Bar
        search_frame = ttk.Frame(left_frame, style="White.TFrame", padding=(15, 0, 15, 10))
        search_frame.pack(fill=tk.X)
        
        self.search_var = tk.StringVar()
        self.search_var.trace("w", lambda name, index, mode, sv=self.search_var: self.refresh_guest_list())
        
        self.search_col_var = tk.StringVar(value="All")
        
        # UI: [Entry] [Combo(Col)]
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, font=Styles.normal_font)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        cols = ["All", "Name", "Category"]
        search_combo = ttk.Combobox(search_frame, textvariable=self.search_col_var, values=cols, state="readonly", width=8, font=Styles.normal_font)
        search_combo.pack(side=tk.RIGHT, padx=(5, 0))
        search_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_guest_list())

        # Scrollable Treeview for Guests
        columns = ("name", "category", "size")
        self.guest_tree = ttk.Treeview(left_frame, columns=columns, show="headings", selectmode="browse")
        
        self.guest_tree.heading("name", text="Name", anchor="w")
        self.guest_tree.heading("category", text="Category", anchor="w")
        self.guest_tree.heading("size", text="Size", anchor="center")
        
        self.guest_tree.column("name", width=140, anchor="w")
        self.guest_tree.column("category", width=100, anchor="w")
        self.guest_tree.column("size", width=40, anchor="center")
        
        # Enable Sorting
        for col in columns:
            self.guest_tree.heading(col, text=col.title(), command=lambda c=col: self.treeview_sort_column(self.guest_tree, c, False))

        self.guest_scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.guest_tree.yview)
        self.guest_tree.configure(yscrollcommand=self.guest_scrollbar.set)
        
        self.guest_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.guest_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        


        # Bind events for Treeview
        self.guest_tree.bind("<ButtonPress-1>", self.on_guest_press)
        self.guest_tree.bind("<B1-Motion>", self.on_guest_drag)
        self.guest_tree.bind("<ButtonRelease-1>", self.on_guest_release)
        self.guest_tree.bind("<Button-3>", self.on_guest_right_click)
        # Mousewheel scrolling
        self.guest_tree.bind_all("<MouseWheel>", self._on_mousewheel)

        # Right Panel: Table Map
        right_frame = ttk.Frame(paned_window, style="TFrame")
        paned_window.add(right_frame, weight=1)
        
        # Tooltip/Header for map
        map_header = ttk.Frame(right_frame, style="TFrame", padding=20)
        map_header.pack(fill=tk.X)
        ttk.Label(map_header, text="Floor Plan", font=Styles.header_font).pack(side=tk.LEFT)
        
        # Zoom control
        zoom_frame = ttk.Frame(map_header, style="TFrame")
        zoom_frame.pack(side=tk.RIGHT)
        ttk.Label(zoom_frame, text="Zoom:", font=Styles.normal_font).pack(side=tk.LEFT, padx=5)
        self.zoom_var = tk.DoubleVar(value=1.0)
        self.zoom_scale = ttk.Scale(zoom_frame, from_=0.5, to=2.0, variable=self.zoom_var, command=lambda v: self.refresh_canvas())
        self.zoom_scale.pack(side=tk.LEFT)

        self.canvas = tk.Canvas(right_frame, bg=Styles.bg_color, highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.canvas.bind("<ButtonPress-1>", self.on_canvas_press)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_release)
        self.canvas.bind("<Double-Button-1>", self.on_table_double_click)
        # Right Click for Context Menu (Linux/Windows use Button-3, macOS might need Button-2 but sticking to standard for now)
        self.canvas.bind("<Button-3>", self.on_canvas_right_click)

        self.update_stats()



    def treeview_sort_column(self, tv, col, reverse):
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        
        # Save sort state
        self.sort_col = col
        self.sort_reverse = reverse
        
        # Try to convert to int for size column or if data is numeric
        try:
            l.sort(key=lambda t: int(t[0]), reverse=reverse)
        except ValueError:
            l.sort(reverse=reverse)

        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)

        # Reverse sort next time
        tv.heading(col, command=lambda: self.treeview_sort_column(tv, col, not reverse))



    def _on_mousewheel(self, event):
        # Determine which widget is under mouse to decide what to scroll
        # or just scroll both/active? 
        # For simplicity, if over left frame, scroll tree, else canvas (if we implemented zoom/pan there)
        # But Treeview handles its own scrolling usually if focused. 
        # Let's just try to pass it to the tree if it's the target.
        pass # Treeview usually handles this natively on most OS, or we need specific binding. 
        # To avoid conflict with default bindings, let's leave it to the widget for now or explicit check.
        # self.guest_tree.yview_scroll(int(-1*(event.delta/120)), "units")

    def update_stats(self):
        total_guests = sum(g.size for g in self.seating_plan.guests.values())
        seated_guests = sum(g.size for g in self.seating_plan.guests.values() if g.table_id is not None)
        unseated = total_guests - seated_guests
        
        total_tables = len(self.seating_plan.tables)
        total_capacity = sum(t.capacity for t in self.seating_plan.tables.values())
        total_occupancy = sum(sum(self.seating_plan.guests[g_id].size for g_id in t.guest_ids) for t in self.seating_plan.tables.values())
        
        text = (f"Guests: {seated_guests}/{total_guests} Seated  •  {unseated} Waiting  |  "
                f"Tables: {total_tables} Active  •  {total_occupancy}/{total_capacity} Seats Used")
        self.stats_label.config(text=text)

    def refresh_guest_list(self):
        # Clear existing items
        for item in self.guest_tree.get_children():
            self.guest_tree.delete(item)
            
        self.current_list_guest_ids = []
        
        self.current_list_guest_ids = []
        
        # Filter unseated guests
        guests = [g for g in self.seating_plan.guests.values() if g.table_id is None]
        
        # Filter unseated guests
        guests = [g for g in self.seating_plan.guests.values() if g.table_id is None]
        
        # Filter by Search
        search_query = self.search_var.get().lower().strip() if hasattr(self, 'search_var') else ""
        search_col = self.search_col_var.get() if hasattr(self, 'search_col_var') else "All"
        
        if search_query:
            filtered = []
            for g in guests:
                match = False
                if search_col == "All":
                    # Check name and category
                    if search_query in g.name.lower() or search_query in g.category.lower():
                        match = True
                elif search_col == "Name":
                     if search_query in g.name.lower(): match = True
                elif search_col == "Category":
                     if search_query in g.category.lower(): match = True
                
                if match:
                    filtered.append(g)
            guests = filtered
        
        # Apply sort
        key_func = lambda g: getattr(g, self.sort_col)
        # Handle case sensitivity for strings if needed, but strict getattr is fine for now.
        # However, for 'size' it's int, for others str.
        # Also 'category' might be empty? Models says default "General".
        
        try:
             # If sorting by size, standard sort works (ints). 
             # If sorting by name/category (strings), we might want case insensitive? 
             # The previous logic was just lambda g: g.name (which is case sensitive).
             # Let's match implicit behavior but support columns.
             guests.sort(key=lambda g: getattr(g, self.sort_col), reverse=self.sort_reverse)
        except:
             # Fallback if attribute missing or comparison fails
             guests.sort(key=lambda g: g.name)
        
        for guest in guests:
                # Insert into Treeview
                # Use guest ID as item ID (iid)
                self.guest_tree.insert("", "end", iid=str(guest.id), values=(
                    self.fix_text(guest.name),
                    self.fix_text(guest.category),
                    guest.size
                ))
                self.current_list_guest_ids.append(guest.id)
        
        self.update_stats()

    # Need to update on_guest_press to work with treeview
    def on_guest_press(self, event):
        item = self.guest_tree.identify_row(event.y)
        if not item:
             # Clicked on empty space, deselect
             self.guest_tree.selection_remove(self.guest_tree.selection())
             return

        # Select the item
        self.guest_tree.selection_set(item)
        
        guest_id = int(item)
        self.drag_data["item"] = guest_id
        self.drag_data["type"] = "guest_list"
        self.drag_data["x"] = event.x
        self.drag_data["y"] = event.y
        
        # Create Ghost
        guest = self.seating_plan.guests[guest_id]
        self.drag_window = tk.Toplevel(self.root)
        self.drag_window.overrideredirect(True)
        self.drag_window.attributes('-alpha', 0.8) # Transparent ghost
        self.drag_window.geometry(f"+{event.x_root}+{event.y_root}")
        
        # Ghost styling
        text = f"{self.fix_text(guest.name)} ({guest.size})"
        # Use a label with solid background for the ghost
        l = tk.Label(self.drag_window, text=text, bg=Styles.primary_color, fg="white", 
                 font=Styles.normal_font, padx=10, pady=5)
        l.pack()
        
        self.drag_window.lift()

    def refresh_canvas(self):
        self.canvas.delete("all")
        # Draw dotted grid
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        # for i in range(0, 2000, 40):
        #     self.canvas.create_line(i, 0, i, 2000, fill="#e5e7eb", dash=(2, 4))
        #     self.canvas.create_line(0, i, 2000, i, fill="#e5e7eb", dash=(2, 4))
            
        for table in self.seating_plan.tables.values():
            self.draw_table(table)
        self.update_stats()


    def draw_table(self, table):
        z = getattr(self, "zoom_var", tk.DoubleVar(value=1.0)).get()
        x, y = table.x * z, table.y * z
        r = 70 * z
        
        # Shadow connection
        self.canvas.create_oval(x-r+4*z, y-r+4*z, x+r+4*z, y+r+4*z, fill="#bdc3c7", outline="", tags=("table", f"table_{table.id}"))
        
        # Table Body
        
        # Color based on fullness
        occupancy = sum(self.seating_plan.guests[g_id].size for g_id in table.guest_ids)
        if occupancy >= table.capacity:
            fill = Styles.table_full_color
            border = Styles.error_color if occupancy > table.capacity else Styles.table_outline_color
        elif occupancy > 0:
            fill = Styles.table_seated_color
            border = Styles.primary_color
        else:
            fill = Styles.table_fill_color
            border = Styles.table_outline_color # Use table_outline_color instead of secondary_hover if not found

        self.canvas.create_oval(x-r, y-r, x+r, y+r, 
                              fill=fill, 
                              outline=border, 
                              width=int(3*z) if occupancy > 0 else int(2*z), 
                              tags=("table", f"table_{table.id}"))
        
        # Table Info
        info_text = f"{self.fix_text(table.name)}\n{occupancy}/{table.capacity}"
        text_color = Styles.accent_color if occupancy >= table.capacity else Styles.text_color
        
        main_font_size = max(8, int(14 * z)) if occupancy < table.capacity else max(10, int(16 * z))
        font_style = (Styles.font_family, main_font_size) if occupancy < table.capacity else (Styles.font_family, main_font_size, "bold")

        self.canvas.create_text(x, y, text=info_text, 
                              font=font_style, 
                              fill=text_color, 
                              justify=tk.CENTER,
                              tags=("table", f"table_{table.id}"))

                # Visual Chairs/Guests
        import math
        if table.capacity > 0:
            angle_step = 360 / table.capacity
            seat_r = 16 * z
            dist = r + (28 * z)
            
            # Flatten guest IDs for seat mapping
            seat_assignments = []
            for g_id in table.guest_ids:
                guest = self.seating_plan.guests[g_id]
                for _ in range(guest.size):
                    seat_assignments.append(g_id)

            for i in range(table.capacity):
                angle = math.radians(i * angle_step - 90) # Start from top
                sx = x + dist * math.cos(angle)
                sy = y + dist * math.sin(angle)
                
                # Check if this seat is occupied
                is_occupied = i < len(seat_assignments)
                
                if is_occupied:
                    guest_id = seat_assignments[i]
                    tags = ("seated_guest", f"guest_{guest_id}")
                    
                    # Seat circle (Filled)
                    self.canvas.create_oval(sx-seat_r, sy-seat_r, sx+seat_r, sy+seat_r,
                                          fill=Styles.primary_color, outline="", width=0,
                                          tags=tags)
                                          
                    # Initials
                    guest = self.seating_plan.guests[guest_id]
                    initial = guest.name[0] if guest.name else "?"
                    small_font_size = max(6, int(10 * z))
                    self.canvas.create_text(sx, sy, text=self.fix_text(initial), font=(Styles.font_family, small_font_size), fill="white", tags=tags)
                else:
                    tags = ("table", f"table_{table.id}")
                    # Empty seat (Outline)
                    self.canvas.create_oval(sx-seat_r+2*z, sy-seat_r+2*z, sx+seat_r-2*z, sy+seat_r-2*z,
                                          fill=Styles.bg_color, outline=Styles.secondary_hover, width=max(1, int(2*z)),
                                          tags=tags)

    # --- Actions --- (unchanged, skipping)

    # --- Drag and Drop Logic --- (on_guest_press/drag are unchanged)

    def on_canvas_press(self, event):
        items = self.canvas.find_closest(event.x, event.y)
        if not items: return
        
        tags = self.canvas.gettags(items[0])
        
        # Check for seated guest first
        guest_id = None
        for tag in tags:
            if tag.startswith("guest_"):
                try:
                    guest_id = int(float(tag.split("_")[1]))
                except ValueError:
                    continue
                break
        
        if guest_id is not None:
            self.drag_data["item"] = guest_id
            self.drag_data["type"] = "seated_guest"
            self.drag_data["x"] = event.x
            self.drag_data["y"] = event.y
            
            # Create Ghost
            guest = self.seating_plan.guests[guest_id]
            self.drag_window = tk.Toplevel(self.root)
            self.drag_window.overrideredirect(True)
            self.drag_window.attributes('-alpha', 0.8)
            self.drag_window.geometry(f"+{event.x_root}+{event.y_root}")
            
            label = tk.Label(self.drag_window, text=self.fix_text(guest.name), bg=Styles.primary_color, fg="white", 
                           font=Styles.normal_font, padx=10, pady=5, relief="solid", borderwidth=1)
            label.pack()
            self.drag_window.lift()
            return

        # Fallback to table drag
        for tag in tags:
            if tag.startswith("table_"):
                try:
                    self.drag_data["item"] = int(float(tag.split("_")[1]))
                except ValueError:
                    continue
                self.drag_data["type"] = "table"
                self.drag_data["x"] = event.x
                self.drag_data["y"] = event.y
                break

    def on_canvas_drag(self, event):
        z = getattr(self, "zoom_var", tk.DoubleVar(value=1.0)).get()
        if self.drag_data["item"] is not None:
            if self.drag_data["type"] == "table":
                dx = (event.x - self.drag_data["x"]) / z
                dy = (event.y - self.drag_data["y"]) / z
                table = self.seating_plan.tables[self.drag_data["item"]]
                table.x += dx
                table.y += dy
                self.drag_data["x"] = event.x
                self.drag_data["y"] = event.y
                self.refresh_canvas()
            elif self.drag_data["type"] == "seated_guest":
                if hasattr(self, 'drag_window'):
                    self.drag_window.geometry(f"+{event.x_root}+{event.y_root}")

    def on_canvas_release(self, event):
        if self.drag_data["type"] == "seated_guest":
            if hasattr(self, 'drag_window'):
                self.drag_window.destroy()
            
            # Check drop target
            items = self.canvas.find_overlapping(event.x, event.y, event.x, event.y)
            target_table_id = None
            
            for item in items:
                tags = self.canvas.gettags(item)
                for tag in tags:
                    if tag.startswith("table_"):
                        try:
                            target_table_id = int(float(tag.split("_")[1]))
                        except ValueError:
                            continue
                        break
                if target_table_id is not None: break
            
            guest_id = self.drag_data["item"]
            
            if target_table_id:
                # Re-seat to new table
                success = self.seating_plan.assign_guest_to_table(guest_id, target_table_id)
                if not success:
                    messagebox.showwarning("Warning", "Target table is full!")
            else:
                # Dropped in empty space -> Unseat
                self.seating_plan.unseat_guest(guest_id)
            
            self.refresh_canvas()
            self.refresh_guest_list()

        self.drag_data = {"item": None, "x": 0, "y": 0, "type": None}

    def add_guest_dialog(self):
        d = RTLStringDialog(self.root, "Add Guest", "Guest Name:")
        name = d.result
        if name:
            d_cat = RTLStringDialog(self.root, "Add Guest", "Category (optional):")
            category = d_cat.result or "General"
            size = simpledialog.askinteger("Add Guest", "Group Size:", minvalue=1, initialvalue=1) or 1
            self.seating_plan.add_guest(name, category, size)
            self.refresh_guest_list()

    def add_table_dialog(self):
        d = RTLStringDialog(self.root, "Add Table", "Table Name:")
        name = d.result
        if name:
            if self.auto_use_default_capacity:
                self.seating_plan.add_table(name, self.default_table_capacity)
                self.refresh_canvas()
            else:
                capacity = simpledialog.askinteger("Add Table", "Capacity:", minvalue=1, initialvalue=self.default_table_capacity)
                if capacity:
                    self.seating_plan.add_table(name, capacity)
                    self.refresh_canvas()

    def settings_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Settings")
        dialog.geometry("300x200")
        dialog.configure(bg=Styles.bg_color)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Capacity
        ttk.Label(dialog, text="Default Table Capacity:", font=Styles.normal_font).pack(pady=(20, 5))
        cap_var = tk.IntVar(value=self.default_table_capacity)
        ttk.Entry(dialog, textvariable=cap_var, justify="center").pack()
        
        # Auto-use Checkbox
        auto_var = tk.BooleanVar(value=self.auto_use_default_capacity)
        cb = ttk.Checkbutton(dialog, text="Auto-use default capacity", variable=auto_var)
        cb.pack(pady=15)
        
        def save():
            try:
                new_cap = cap_var.get()
            except:
                messagebox.showerror("Error", "Invalid capacity.")
                return

            if new_cap < 1:
                messagebox.showerror("Error", "Capacity must be at least 1.")
                return
            
            self.default_table_capacity = new_cap
            self.auto_use_default_capacity = auto_var.get()
            dialog.destroy()
            messagebox.showinfo("Success", "Settings saved!")
            
        ttk.Button(dialog, text="Save", command=save, style="Primary.TButton").pack(pady=10)

    def save_plan(self):
        filename = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON Files", "*.json")])
        if filename:
            self.seating_plan.save_to_file(filename)
            messagebox.showinfo("Success", "Plan saved successfully!")

    def load_plan(self):
        filename = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if filename:
            self.seating_plan.load_from_file(filename)
            self.refresh_guest_list()
            self.refresh_canvas()
    
    def save_excel(self):
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if filename:
            try:
                ExcelIO.save_to_xlsx(self.seating_plan, filename)
                messagebox.showinfo("Success", "Plan saved to Excel successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save Excel: {e}")

    def load_excel(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if filename:
            try:
                # Ask user if they want to replace or merge
                if getattr(self, "seating_plan", None) and (self.seating_plan.guests or self.seating_plan.tables):
                    answer = messagebox.askyesnocancel("Load Excel", "Do you want to clear the existing plan completely?\n\nYes: Clear and Replace\nNo: Merge into current plan\nCancel: Abort load")
                    if answer is None:
                        return # Cancelled
                    clear_plan = answer
                else:
                    clear_plan = True
                    
                ExcelIO.load_from_xlsx(filename, self.seating_plan, clear=clear_plan)
                self.refresh_guest_list()
                self.refresh_canvas()
                messagebox.showinfo("Success", "Plan loaded from Excel successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load Excel: {e}")

    def import_groups_dialog(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not filename:
            return

        try:
            headers = ExcelIO.get_headers(filename)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file: {e}")
            return

        if not headers:
            messagebox.showerror("Error", "No headers found in the first row.")
            return

        # Dialog for column mapping
        dialog = tk.Toplevel(self.root)
        dialog.title("Import Guest Groups")
        dialog.geometry("400x500")
        dialog.configure(bg=Styles.bg_color)
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Map Columns", font=Styles.header_font).pack(pady=10)

        # Group Name
        ttk.Label(dialog, text="Group Name Column:").pack(pady=5)
        group_var = tk.StringVar(dialog)
        group_combo = ttk.Combobox(dialog, textvariable=group_var, values=[self.fix_text(h) for h in headers], state="readonly")
        group_combo.pack()
        if headers: group_combo.current(0)

        # Guest Count
        ttk.Label(dialog, text="Guest Count Column:").pack(pady=5)
        count_var = tk.StringVar(dialog)
        count_combo = ttk.Combobox(dialog, textvariable=count_var, values=[self.fix_text(h) for h in headers], state="readonly")
        count_combo.pack()
        if len(headers) > 1: count_combo.current(1)

        # Category (Optional)
        ttk.Label(dialog, text="Category Column (Optional):").pack(pady=5)
        cat_var = tk.StringVar(dialog)
        cat_combo = ttk.Combobox(dialog, textvariable=cat_var, values=["(None)"] + [self.fix_text(h) for h in headers], state="readonly")
        cat_combo.pack()
        cat_combo.current(0)

        def do_import():
            # We need to map back the fixed text to the original header for internal logic if we were using it for lookups not by index?
            # Actually import_groups_to_plan uses the string value to find index.
            # But the user sees fixed text.
            # If we pass fixed text to import_groups, it might fail to find it in the file if file headers aren't fixed.
            # However, `get_headers` returns raw strings.
            # Let's use indices or map back. 
            # A simple way is to keep a map of fixed -> original
            header_map = {self.fix_text(h): h for h in headers}
            
            group_col_display = group_var.get()
            count_col_display = count_var.get()
            
            group_col = header_map.get(group_col_display)
            count_col = header_map.get(count_col_display)
            
            cat_col_display = cat_var.get()
            category_col = None
            if cat_col_display and cat_col_display != "(None)":
                category_col = header_map.get(cat_col_display)
            
            if not group_col or not count_col:
                messagebox.showerror("Error", "Please select both columns.")
                return

            try:
                ExcelIO.import_groups_to_plan(filename, group_col, count_col, self.seating_plan, category_col)
                self.refresh_guest_list()
                self.update_stats()
                messagebox.showinfo("Success", "Groups imported successfully!")
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Import failed: {e}")

        ttk.Button(dialog, text="Import", command=do_import, style="Primary.TButton").pack(pady=20)

    def export_to_sheets(self):
        # check for credentials
        creds_file = "credentials.json"
        if not os.path.exists(creds_file):
             messagebox.showerror("Error", "credentials.json not found in the application directory. Please add your Google Service Account credentials to use this feature.")
             return

        sheet_name = simpledialog.askstring("Export to Sheets", "Enter Google Sheet Name or URL:")
        if sheet_name:
            try:
                exporter = GoogleSheetsExporter(creds_file)
                url = exporter.export(self.seating_plan, sheet_name)
                messagebox.showinfo("Success", f"Exported successfully!\nCheck your Google Drive or open:\n{url}")
            except Exception as e:
                messagebox.showerror("Export Failed", str(e))

    # --- Drag and Drop Logic ---

    def on_guest_press(self, event):
        # Identify item under cursor in the guest list tree
        item = self.guest_tree.identify_row(event.y)
        if not item: return

        guest_id = int(item)
        
        self.drag_data["item"] = guest_id
        self.drag_data["type"] = "guest_list"
        self.drag_data["x"] = event.x
        self.drag_data["y"] = event.y
        
        # Create a drag ghost
        guest = self.seating_plan.guests[guest_id]
        self.drag_window = tk.Toplevel(self.root)
        self.drag_window.overrideredirect(True)
        self.drag_window.attributes('-alpha', 0.8)
        self.drag_window.geometry(f"+{event.x_root}+{event.y_root}")
        
        # Styled ghost label
        text = f"{self.fix_text(guest.name)} ({guest.size})"
        label = tk.Label(self.drag_window, text=text, bg=Styles.primary_color, fg="white", 
                       font=Styles.normal_font, padx=10, pady=5, relief="solid", borderwidth=1)
        label.pack()
        
        self.drag_window.lift()

    def on_guest_drag(self, event):
        if self.drag_data["item"] is not None and self.drag_data["type"] == "guest_list":
             if hasattr(self, 'drag_window'):
                self.drag_window.geometry(f"+{event.x_root}+{event.y_root}")

    def on_guest_release(self, event):
        if self.drag_data["item"] is not None and self.drag_data["type"] == "guest_list":
            if hasattr(self, 'drag_window'):
                self.drag_window.destroy()
            
            # Check if dropped on GUI map
            x_root, y_root = event.x_root, event.y_root
            canvas_x = x_root - self.canvas.winfo_rootx()
            canvas_y = y_root - self.canvas.winfo_rooty()

            # Find overlapping items on canvas
            items = self.canvas.find_overlapping(canvas_x, canvas_y, canvas_x, canvas_y)
            for item in items:
                tags = self.canvas.gettags(item)
                for tag in tags:
                    if tag.startswith("table_"):
                        try:
                            table_id = int(float(tag.split("_")[1]))
                        except ValueError:
                            continue
                        success = self.seating_plan.assign_guest_to_table(self.drag_data["item"], table_id)
                        if not success:
                            messagebox.showwarning("Warning", "Table is full!")
                        else:
                            self.refresh_guest_list()
                            self.refresh_canvas()
                        break

        self.drag_data = {"item": None, "x": 0, "y": 0, "type": None}

    def on_canvas_press(self, event):
        items = self.canvas.find_closest(event.x, event.y)
        tags = self.canvas.gettags(items[0]) if items else []
        for tag in tags:
            if tag.startswith("table_"):
                try:
                    self.drag_data["item"] = int(float(tag.split("_")[1]))
                except ValueError:
                    continue
                self.drag_data["type"] = "table"
                self.drag_data["x"] = event.x
                self.drag_data["y"] = event.y
                break

    def on_canvas_drag(self, event):
        z = getattr(self, "zoom_var", tk.DoubleVar(value=1.0)).get()
        if self.drag_data["item"] is not None and self.drag_data["type"] == "table":
            dx = (event.x - self.drag_data["x"]) / z
            dy = (event.y - self.drag_data["y"]) / z
            table = self.seating_plan.tables[self.drag_data["item"]]
            table.x += dx
            table.y += dy
            self.drag_data["x"] = event.x
            self.drag_data["y"] = event.y
            self.refresh_canvas()

    def on_canvas_release(self, event):
        self.drag_data = {"item": None, "x": 0, "y": 0, "type": None}

    def on_table_double_click(self, event):
        # Identify table
        items = self.canvas.find_closest(event.x, event.y)
        tags = self.canvas.gettags(items[0]) if items else []
        for tag in tags:
            if tag.startswith("table_"):
                try:
                    table_id = int(float(tag.split("_")[1]))
                except ValueError:
                    continue
                table = self.seating_plan.tables[table_id]
                self.show_table_details(table)
                break

    def show_table_details(self, table):
        # Show seated guests and allow removal
        detail_window = tk.Toplevel(self.root)
        detail_window.title(f"Table: {table.name}")
        detail_window.configure(bg=Styles.bg_color)
        detail_window.minsize(300, 250)
        
        def refresh_list():
            listbox.delete(0, tk.END)
            # Re-fetch guest ids as they might have changed
            for guest_id in table.guest_ids:
                guest = self.seating_plan.guests[guest_id]
                display = f"{guest.name} ({guest.size})" if guest.size > 1 else guest.name
                listbox.insert(tk.END, self.fix_text(display))
            current_occupancy = sum(self.seating_plan.guests[g_id].size for g_id in table.guest_ids)
            occ_label.config(text=f"Guests ({current_occupancy}/{table.capacity})")

        occ_label = ttk.Label(detail_window, text=f"Guests ({current_occupancy}/{table.capacity})", font=Styles.normal_font)
        occ_label.pack(pady=5)
        
        listbox = tk.Listbox(detail_window, font=Styles.normal_font, borderwidth=1, relief="solid")
        listbox.pack(padx=10, pady=5, expand=True, fill=tk.BOTH)
        
        refresh_list()
        
        def edit_selected_guest(event=None):
            selection = listbox.curselection()
            if selection:
                index = selection[0]
                guest_id = table.guest_ids[index]
                self.edit_guest_properties(guest_id)
                refresh_list()
                
        listbox.bind("<Double-Button-1>", edit_selected_guest)
        
        def remove_selected():
            selection = listbox.curselection()
            if selection:
                index = selection[0]
                guest_id = table.guest_ids[index]
                self.seating_plan.unseat_guest(guest_id)
                refresh_list()
                self.refresh_canvas()
                self.refresh_guest_list()

        ttk.Button(detail_window, text="Edit Guest", command=edit_selected_guest, style="Primary.TButton").pack(pady=5)
        ttk.Button(detail_window, text="Remove Selected", command=remove_selected, style="Secondary.TButton").pack(pady=10)

    def on_canvas_right_click(self, event):
        # Determine target
        items = self.canvas.find_closest(event.x, event.y)
        table_id = None
        
        if items:
            tags = self.canvas.gettags(items[0])
            for tag in tags:
                if tag.startswith("table_"):
                    try:
                        table_id = int(float(tag.split("_")[1]))
                    except ValueError:
                        continue
                    break
        
        if table_id is not None:
            self.show_table_context_menu(event, table_id)
        else:
            self.show_canvas_context_menu(event)

    def show_table_context_menu(self, event, table_id):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Edit Properties (Name/Capacity)", command=lambda: self.edit_table_properties(table_id))
        menu.add_command(label="Edit Table ID", command=lambda: self.edit_table_id(table_id))
        menu.add_command(label="Add Guest to Table", command=lambda: self.add_guest_to_table_dialog(table_id))
        menu.add_separator()
        menu.add_command(label="Delete Table", command=lambda: self.delete_table(table_id), foreground="red")
        menu.post(event.x_root, event.y_root)

    def show_canvas_context_menu(self, event):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Add Table Here", command=lambda: self.add_table_at_pos(event.x, event.y))
        menu.post(event.x_root, event.y_root)

    def edit_table_properties(self, table_id):
        table = self.seating_plan.tables[table_id]
        new_name = simpledialog.askstring("Edit Table", "Table Name:", initialvalue=table.name)
        if new_name:
            new_capacity = simpledialog.askinteger("Edit Table", "Capacity:", initialvalue=table.capacity, minvalue=1)
            if new_capacity:
                table.name = new_name
                table.capacity = new_capacity
                self.refresh_canvas()
                self.update_stats()
                
    def edit_table_id(self, table_id):
        table = self.seating_plan.tables[table_id]
        new_id = simpledialog.askinteger("Edit Table ID", "New Table ID (must be unique):", initialvalue=table.id, minvalue=1)
        if new_id and new_id != table.id:
            if new_id in self.seating_plan.tables:
                messagebox.showerror("Error", f"Table ID {new_id} already exists!")
                return
            
            # Update all guest references
            for guest_id in table.guest_ids:
                self.seating_plan.guests[guest_id].table_id = new_id
                
            table.id = new_id
            self.seating_plan.tables[new_id] = table
            del self.seating_plan.tables[table_id]
            
            if new_id >= self.seating_plan.next_table_id:
                self.seating_plan.next_table_id = new_id + 1
            
            self.refresh_canvas()
            self.update_stats()

    def delete_table(self, table_id):
        if messagebox.askyesno("Delete Table", "Are you sure you want to delete this table?\nGuests will be unseated."):
            self.seating_plan.remove_table(table_id)
            self.refresh_canvas()
            self.refresh_guest_list()

    def add_guest_to_table_dialog(self, table_id):
        # Simple dialog to add a NEW guest directly to this table
        # Or pick from unseated? Let's offer both or just "Add New" for now as "Pick" is drag-drop
        name = simpledialog.askstring("Add Guest to Table", "New Guest Name:")
        if name:
             category = simpledialog.askstring("Add Guest", "Category (optional):") or "General"
             size = simpledialog.askinteger("Add Guest", "Group Size:", minvalue=1, initialvalue=1) or 1
             guest = self.seating_plan.add_guest(name, category, size)
             success = self.seating_plan.assign_guest_to_table(guest.id, table_id)
             if not success:
                 messagebox.showwarning("Warning", "Table is full!")
                 # maybe roll back guest creation? Or leave it unseated. Leaving it unseated is safer.
             self.refresh_canvas()
             self.refresh_guest_list()

    def add_table_at_pos(self, x, y):
        name = simpledialog.askstring("Add Table", "Table Name:")
        if name:
            if self.auto_use_default_capacity:
                table = self.seating_plan.add_table(name, self.default_table_capacity)
                table.x = x
                table.y = y
                self.refresh_canvas()
            else:
                capacity = simpledialog.askinteger("Add Table", "Capacity:", minvalue=1, initialvalue=self.default_table_capacity)
                if capacity:
                    table = self.seating_plan.add_table(name, capacity)
                    table.x = x
                    table.y = y
                    self.refresh_canvas()

    def on_guest_right_click(self, event):
        item = self.guest_tree.identify_row(event.y)
        if not item: return
        
        self.guest_tree.selection_set(item)
        guest_id = int(item)
        
        if guest_id is not None:
            self.show_guest_context_menu(event, guest_id)

    def show_guest_context_menu(self, event, guest_id):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Edit Guest", command=lambda: self.edit_guest_properties(guest_id))
        menu.add_separator()
        menu.add_command(label="Delete Guest", command=lambda: self.delete_guest(guest_id), foreground="red")
        menu.post(event.x_root, event.y_root)

    def edit_guest_properties(self, guest_id):
        guest = self.seating_plan.guests[guest_id]
        new_name = simpledialog.askstring("Edit Guest", "Name:", initialvalue=guest.name)
        if new_name:
            new_category = simpledialog.askstring("Edit Guest", "Category:", initialvalue=guest.category)
            if new_category is not None: # check for None in case of cancel, empty string is valid
                new_size = simpledialog.askinteger("Edit Guest", "Group Size:", initialvalue=guest.size, minvalue=1)
                if new_size:
                    guest.name = new_name
                    guest.category = new_category
                    # Check if size change affects seating
                    old_size = guest.size
                    guest.size = new_size
                    
                    if guest.table_id:
                        # Re-validate capacity
                        table = self.seating_plan.tables[guest.table_id]
                        occupancy = sum(self.seating_plan.guests[g_id].size for g_id in table.guest_ids)
                        # We already updated size, so occupancy reflects new size
                        if occupancy > table.capacity:
                            messagebox.showwarning("Warning", f"New size exceeds table capacity ({occupancy}/{table.capacity}). Guest unseated.")
                            self.seating_plan.unseat_guest(guest.id)

                self.refresh_guest_list() # Re-draw to show changes
                if guest.table_id:
                     self.refresh_canvas() # Update map if seated (initials might change)

    def delete_guest(self, guest_id):
        if messagebox.askyesno("Delete Guest", "Are you sure you want to delete this guest?"):
            self.seating_plan.remove_guest(guest_id)
            self.refresh_guest_list()
            self.refresh_canvas() # In case they were seated, though this menu is for the unseated list (mostly)
            self.update_stats()
