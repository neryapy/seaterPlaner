# Wedding Seating Planner

A modern, intuitive desktop application for planning wedding guest seating arrangements. Design your floor plan, manage guest lists, and export your setup to Excel or Google Sheets.

![App Preview](https://via.placeholder.com/800x450.png?text=Wedding+Seating+Planner+Map)

## ✨ Features

- **Visual Floor Plan**: Drag and drop tables and guests on a dynamic canvas.
- **Smart Drag & Drop**: Seat guests by dragging them from the list onto tables, or move them between tables.
- **Interactive Zoom**: A dedicated zoom bar to view your entire floor plan or focus on specific sections.
- **RTL & Arabic Support**: Fully supports Hebrew and Arabic names with proper reshaping and BiDi layout.
- **Advanced Excel Integration**:
    - **Import/Export**: Save and load your entire plan as `.xlsx`.
    - **Formula-based references**: Tables IDs in Excel are linked via formulas (`=Tables!A2`) for easy external editing.
    - **Merge Mode**: Import multiple Excel files and merge them into one plan with automatic ID collision resolution.
- **Google Sheets Export**: One-click export to your Google Drive for easy sharing.
- **Flexible Editing**: Change table names, capacities, and IDs directly from context menus.

## 🚀 Installation & Running

### Requirements
- Python 3.8 or higher
- All dependencies listed in `requirements.txt`

### From Source
1. Clone the repository.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the application:
   ```bash
   python -m wedding_planner.main
   ```

### Releases (Portable EXE)
You can find the standalone `.exe` version in the [Releases](https://github.com/neryapy/wedding-planner/releases) section. No Python installation is required for the EXE version.

## 📋 Usage Tips

- **Right-Click**: Access powerful context menus on the guest list, table map, or individual tables.
- **Double-Click**: 
    - Double-click a **Table** to see the guest list and edit properties.
    - Double-click a **Guest** inside the table details to edit their name or category.
- **Zoom Bar**: Use the slider in the top right to adjust the scale of the map.
- **Google Sheets**: To use the export feature, place your `credentials.json` service account file in the application directory.

## 🛠️ Requirements

- `openpyxl`: For Excel read/write support.
- `gspread`: For Google Sheets API integration.
- `arabic_reshaper`: For RTL text processing.
- `python-bidi`: For correct RTL display algorithms.

---
Built with ❤️ for your special day.
