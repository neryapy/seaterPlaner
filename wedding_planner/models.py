import json
from dataclasses import dataclass, field
from typing import List, Optional, Dict

@dataclass
class Guest:
    id: int
    name: str
    category: str = "General"
    table_id: Optional[int] = None
    size: int = 1

    def to_dict(self):
        return {
            "id": self.id,
            "name": self.name,
            "category": self.category,
            "table_id": self.table_id,
            "size": self.size
        }

    @classmethod
    def from_dict(cls, data):
        return cls(**data)

@dataclass
class Table:
    id: int
    name: str
    capacity: int
    guest_ids: List[int] = field(default_factory=list)
    x: int = 0
    y: int = 0

    def to_dict(self):
        return {
            "id": self.id,
            "name": self.name,
            "capacity": self.capacity,
            "guest_ids": self.guest_ids,
            "x": self.x,
            "y": self.y
        }

    @classmethod
    def from_dict(cls, data):
        return cls(**data)

class SeatingPlan:
    def __init__(self):
        self.guests: Dict[int, Guest] = {}
        self.tables: Dict[int, Table] = {}
        self.next_guest_id = 1
        self.next_table_id = 1

    def add_guest(self, name: str, category: str = "General", size: int = 1) -> Guest:
        guest = Guest(id=self.next_guest_id, name=name, category=category, size=size)
        self.guests[guest.id] = guest
        self.next_guest_id += 1
        return guest

    def remove_guest(self, guest_id: int):
        if guest_id in self.guests:
            guest = self.guests[guest_id]
            if guest.table_id is not None:
                self.unseat_guest(guest_id)
            del self.guests[guest_id]

    def add_table(self, name: str, capacity: int, x: int = 100, y: int = 100) -> Table:
        table = Table(id=self.next_table_id, name=name, capacity=capacity, x=x, y=y)
        self.tables[table.id] = table
        self.next_table_id += 1
        return table

    def remove_table(self, table_id: int):
        if table_id in self.tables:
            table = self.tables[table_id]
            # Unseat all guests at this table
            for guest_id in list(table.guest_ids):
                self.unseat_guest(guest_id)
            del self.tables[table_id]

    def assign_guest_to_table(self, guest_id: int, table_id: int) -> bool:
        if guest_id not in self.guests or table_id not in self.tables:
            return False
        
        guest = self.guests[guest_id]
        table = self.tables[table_id]

        # Calculate current occupancy
        current_occupancy = sum(self.guests[g_id].size for g_id in table.guest_ids)

        # Check capacity
        if current_occupancy + guest.size > table.capacity:
            return False

        # If guest is already seated, unseat them first
        if guest.table_id is not None:
            self.unseat_guest(guest_id)

        guest.table_id = table_id
        table.guest_ids.append(guest_id)
        return True

    def unseat_guest(self, guest_id: int):
        if guest_id in self.guests:
            guest = self.guests[guest_id]
            if guest.table_id is not None:
                if guest.table_id in self.tables:
                    table = self.tables[guest.table_id]
                    if guest_id in table.guest_ids:
                        table.guest_ids.remove(guest_id)
                guest.table_id = None

    def save_to_file(self, filename: str):
        data = {
            "guests": [g.to_dict() for g in self.guests.values()],
            "tables": [t.to_dict() for t in self.tables.values()],
            "next_guest_id": self.next_guest_id,
            "next_table_id": self.next_table_id
        }
        with open(filename, 'w') as f:
            json.dump(data, f, indent=4)

    def load_from_file(self, filename: str):
        with open(filename, 'r') as f:
            data = json.load(f)
        
        self.guests = {}
        self.tables = {}
        self.next_guest_id = data.get("next_guest_id", 1)
        self.next_table_id = data.get("next_table_id", 1)

        for g_data in data.get("guests", []):
            guest = Guest.from_dict(g_data)
            self.guests[guest.id] = guest
        
        for t_data in data.get("tables", []):
            table = Table.from_dict(t_data)
            self.tables[table.id] = table

