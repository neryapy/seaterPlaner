import gspread
import json
from .models import SeatingPlan

class GoogleSheetsExporter:
    def __init__(self, credentials_file: str):
        self.gc = gspread.service_account(filename=credentials_file)

    def export(self, seating_plan: SeatingPlan, sheet_identifier: str):
        try:
            if "docs.google.com/spreadsheets" in sheet_identifier:
                sh = self.gc.open_by_url(sheet_identifier)
            else:
                sh = self.gc.open(sheet_identifier)
        except gspread.exceptions.SpreadsheetNotFound:
            if "docs.google.com" in sheet_identifier:
                raise Exception("Could not access the provided URL. Make sure the Service Account has 'Editor' access.")
            # Create if it doesn't exist (only by name)
            sh = self.gc.create(sheet_identifier)
            print(f"Created new sheet: {sh.url}")

        # Export Guests
        try:
            guest_worksheet = sh.worksheet("Guests")
        except gspread.exceptions.WorksheetNotFound:
            guest_worksheet = sh.add_worksheet(title="Guests", rows=100, cols=10)
        
        guest_worksheet.clear()
        guest_header = ["Guest Name", "Category", "Group Dimension", "Table"]
        guest_data = []
        
        # Sort guests by name
        sorted_guests = sorted(seating_plan.guests.values(), key=lambda x: x.name)
        
        for guest in sorted_guests:
            table_name = "Unseated"
            if guest.table_id is not None and guest.table_id in seating_plan.tables:
                table_name = seating_plan.tables[guest.table_id].name
            guest_data.append([guest.name, guest.category, guest.size, table_name])
            
        guest_worksheet.update([guest_header] + guest_data)

        # Export Tables
        try:
            table_worksheet = sh.worksheet("Tables")
        except gspread.exceptions.WorksheetNotFound:
            table_worksheet = sh.add_worksheet(title="Tables", rows=100, cols=10)
            
        table_worksheet.clear()
        table_header = ["Table Name", "Capacity", "Occupancy", "Status"]
        table_data = []
        
        sorted_tables = sorted(seating_plan.tables.values(), key=lambda x: x.name)
        
        for table in sorted_tables:
            occupancy = len(table.guest_ids)
            status = "Full" if occupancy >= table.capacity else f"{table.capacity - occupancy} seats left"
            table_data.append([table.name, table.capacity, occupancy, status])
            
        table_worksheet.update([table_header] + table_data)

        return sh.url
