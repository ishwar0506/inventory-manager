import sqlite3
from datetime import date, datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import csv
import re
import os
import sys
try:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import Font, Border, Side, Alignment
    from openpyxl.worksheet.table import Table, TableStyleInfo
    print("openpyxl imported successfully, version:", openpyxl.__version__)
except ImportError as e:
    messagebox.showerror("Missing Module", f"Failed to import openpyxl: {e}\nPlease install openpyxl:\npip install openpyxl")
    exit(1)

try:
    from tkcalendar import DateEntry
except ImportError:
    messagebox.showerror("Missing Module", "Please install tkcalendar:\npip install tkcalendar")
    exit(1)

# Undo/Redo Manager Class
class UndoManager:
    def __init__(self):
        self.undo_stack = []
        self.redo_stack = []
        self.max_size = 50  # Maximum number of actions to store
    
    def add_action(self, action_type, table, action_data):
        """Add an action to the undo stack"""
        self.undo_stack.append({
            'type': action_type,
            'table': table,
            'data': action_data,
            'timestamp': datetime.now()
        })
        # Clear redo stack when new action is added
        self.redo_stack.clear()
        # Limit stack size
        if len(self.undo_stack) > self.max_size:
            self.undo_stack.pop(0)
    
    def can_undo(self):
        return len(self.undo_stack) > 0
    
    def can_redo(self):
        return len(self.redo_stack) > 0
    
    def undo(self):
        if not self.can_undo():
            return None
        action = self.undo_stack.pop()
        self.redo_stack.append(action)
        return action
    
    def redo(self):
        if not self.can_redo():
            return None
        action = self.redo_stack.pop()
        self.undo_stack.append(action)
        return action
    
    def clear(self):
        """Clear all stacks (useful when resetting)"""
        self.undo_stack.clear()
        self.redo_stack.clear()

# Global undo manager instance
undo_manager = UndoManager()

def record_action(action_type, table, action_data):
    """Record an action for undo/redo"""
    undo_manager.add_action(action_type, table, action_data)

def undo_last_action(current_tab, tree_widgets, entry_widgets):
    """Undo the last action based on current tab"""
    action = undo_manager.undo()
    if not action:
        messagebox.showinfo("Undo", "Nothing to undo")
        return False
    
    try:
        action_type = action['type']
        table = action['table']
        data = action['data']
        
        if action_type == 'add_item':
            # Reverse: Delete the added item entry
            item_name = data.get('item_name')
            ref_no = data.get('ref_no')
            entry_date = data.get('date')
            entity = data.get('entity')
            qty = data.get('quantity')
            
            ref_col = 'memo_no' if table == 'inward_inventory' else 'bale_no'
            entity_col = 'merchant' if table == 'inward_inventory' else 'transport'
            
            # Find and delete the added entry
            cursor.execute(f"SELECT id, quantity FROM {table} WHERE item_name = ? AND {ref_col} = ? AND date = ? AND {entity_col} = ?",
                          (item_name, ref_no, entry_date, entity))
            entry = cursor.fetchone()
            
            if entry:
                entry_id, current_qty = entry
                new_qty = current_qty - qty
                
                if new_qty <= 0:
                    cursor.execute(f"DELETE FROM {table} WHERE id = ?", (entry_id,))
                else:
                    cursor.execute(f"UPDATE {table} SET quantity = ? WHERE id = ?", (new_qty, entry_id))
                
                conn.commit()
                messagebox.showinfo("Undo", f"Successfully undid {action_type}")
                return True
                
        elif action_type == 'add_order':
            # Reverse: Delete the added order
            party = data.get('party')
            agent = data.get('agent')
            quality = data.get('quality')
            entry_date = data.get('date')
            qty = data.get('quantity')
            
            cursor.execute("SELECT id, quantity FROM orders WHERE party_name = ? AND agent = ? AND quality_name = ? AND date = ?",
                          (party, agent, quality, entry_date))
            entry = cursor.fetchone()
            
            if entry:
                entry_id, current_qty = entry
                new_qty = current_qty - qty
                
                if new_qty <= 0:
                    cursor.execute("DELETE FROM orders WHERE id = ?", (entry_id,))
                else:
                    cursor.execute("UPDATE orders SET quantity = ? WHERE id = ?", (new_qty, entry_id))
                
                conn.commit()
                messagebox.showinfo("Undo", f"Successfully undid {action_type}")
                return True
                
        elif action_type == 'transfer_stock':
            # Reverse: Reverse the transfer
            source_item = data.get('source_item')
            dest_item = data.get('dest_item')
            transfer_qty = data.get('quantity')
            entry_date = data.get('date')
            
            # Delete the transfer record
            cursor.execute("DELETE FROM stock_transfer WHERE source_item = ? AND dest_item = ? AND quantity = ? AND date = ?",
                          (source_item, dest_item, transfer_qty, entry_date))
            
            # Reverse the inventory changes
            if transfer_qty > 0:
                # Reverse: Add back to source, remove from destination
                cursor.execute("DELETE FROM outward_inventory WHERE item_name = ? AND date = ? AND bale_no LIKE ?",
                              (source_item, entry_date, f"TRANSFER_TO%"))
                cursor.execute("DELETE FROM inward_inventory WHERE item_name = ? AND date = ? AND memo_no LIKE ?",
                              (dest_item, entry_date, f"TRANSFER_FROM%"))
            else:
                abs_qty = abs(transfer_qty)
                cursor.execute("DELETE FROM outward_inventory WHERE item_name = ? AND date = ? AND bale_no LIKE ?",
                              (dest_item, entry_date, f"TRANSFER_TO%"))
                cursor.execute("DELETE FROM inward_inventory WHERE item_name = ? AND date = ? AND memo_no LIKE ?",
                              (source_item, entry_date, f"TRANSFER_FROM%"))
            
            conn.commit()
            messagebox.showinfo("Undo", f"Successfully undid transfer of {abs(transfer_qty)} pcs")
            return True
            
        elif action_type == 'delete_item':
            # Reverse: Restore deleted item (requires storing all records)
            # This is complex, so we'll show a message
            messagebox.showinfo("Info", "Delete operations cannot be undone. Please restore manually if needed.")
            return False
            
        elif action_type == 'edit_item':
            # Reverse: Rename back to original name
            old_name = data.get('old_name')
            new_name = data.get('new_name')
            cursor.execute(f"UPDATE {table} SET item_name = ? WHERE item_name = ?", (old_name, new_name))
            conn.commit()
            messagebox.showinfo("Undo", f"Successfully restored item name to {old_name.title()}")
            return True
        
        return False
        
    except Exception as e:
        messagebox.showerror("Undo Error", f"Failed to undo action: {str(e)}")
        return False

def redo_last_action(current_tab, tree_widgets, entry_widgets):
    """Redo the last undone action"""
    action = undo_manager.redo()
    if not action:
        messagebox.showinfo("Redo", "Nothing to redo")
        return False
    
    try:
        action_type = action['type']
        table = action['table']
        data = action['data']
        
        if action_type == 'add_item':
            # Re-add the item
            item_name = data.get('item_name')
            ref_no = data.get('ref_no')
            entry_date = data.get('date')
            entity = data.get('entity')
            qty = data.get('quantity')
            
            ref_col = 'memo_no' if table == 'inward_inventory' else 'bale_no'
            entity_col = 'merchant' if table == 'inward_inventory' else 'transport'
            
            cursor.execute(f"SELECT id, quantity FROM {table} WHERE item_name = ? AND {ref_col} = ? AND date = ? AND {entity_col} = ?",
                          (item_name, ref_no, entry_date, entity))
            entry = cursor.fetchone()
            
            if entry:
                entry_id, current_qty = entry
                cursor.execute(f"UPDATE {table} SET quantity = ? WHERE id = ?", (current_qty + qty, entry_id))
            else:
                cursor.execute(f"INSERT INTO {table} (item_name, {ref_col}, quantity, date, {entity_col}) VALUES (?, ?, ?, ?, ?)",
                              (item_name, ref_no, qty, entry_date, entity))
            
            conn.commit()
            messagebox.showinfo("Redo", f"Successfully redid {action_type}")
            return True
            
        elif action_type == 'add_order':
            # Re-add the order
            party = data.get('party')
            agent = data.get('agent')
            quality = data.get('quality')
            entry_date = data.get('date')
            qty = data.get('quantity')
            
            cursor.execute("SELECT id, quantity FROM orders WHERE party_name = ? AND agent = ? AND quality_name = ? AND date = ?",
                          (party, agent, quality, entry_date))
            entry = cursor.fetchone()
            
            if entry:
                entry_id, current_qty = entry
                cursor.execute("UPDATE orders SET quantity = ? WHERE id = ?", (current_qty + qty, entry_id))
            else:
                cursor.execute("INSERT INTO orders (party_name, quality_name, quantity, date, agent) VALUES (?, ?, ?, ?, ?)",
                              (party, quality, qty, entry_date, agent))
            
            conn.commit()
            messagebox.showinfo("Redo", f"Successfully redid {action_type}")
            return True
            
        return False
        
    except Exception as e:
        messagebox.showerror("Redo Error", f"Failed to redo action: {str(e)}")
        return False

# Define unit conversion rules
unit_multipliers = {
    'd': 12,    # 1d = 12 pcs
}

conn = sqlite3.connect("basic_inventory.db")
cursor = conn.cursor()

# Create inward_inventory table if not exists
cursor.execute('''
    CREATE TABLE IF NOT EXISTS inward_inventory (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        item_name TEXT NOT NULL,
        memo_no TEXT,
        quantity INTEGER NOT NULL,
        date TEXT NOT NULL,
        merchant TEXT
    )
''')

# Create outward_inventory table if not exists
cursor.execute('''
    CREATE TABLE IF NOT EXISTS outward_inventory (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        item_name TEXT NOT NULL,
        bale_no TEXT,
        quantity INTEGER NOT NULL,
        date TEXT NOT NULL,
        transport TEXT
    )
''')

# Create stock_transfer table if not exists
cursor.execute('''
    CREATE TABLE IF NOT EXISTS stock_transfer (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        source_item TEXT NOT NULL,
        quantity INTEGER NOT NULL,
        dest_item TEXT NOT NULL,
        date TEXT NOT NULL
    )
''')

# Create orders table if not exists
cursor.execute('''
    CREATE TABLE IF NOT EXISTS orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        party_name TEXT NOT NULL,
        quality_name TEXT NOT NULL,
        quantity INTEGER NOT NULL,
        date TEXT NOT NULL
    )
''')

# Check and add columns if missing
cursor.execute("PRAGMA table_info(inward_inventory)")
columns = [col[1] for col in cursor.fetchall()]
if 'merchant' not in columns:
    cursor.execute("ALTER TABLE inward_inventory ADD COLUMN merchant TEXT")
if 'memo_no' not in columns:
    cursor.execute("ALTER TABLE inward_inventory ADD COLUMN memo_no TEXT")

cursor.execute("PRAGMA table_info(outward_inventory)")
columns = [col[1] for col in cursor.fetchall()]
if 'transport' not in columns:
    cursor.execute("ALTER TABLE outward_inventory ADD COLUMN transport TEXT")
if 'bale_no' not in columns:
    cursor.execute("ALTER TABLE outward_inventory ADD COLUMN bale_no TEXT")

cursor.execute("PRAGMA table_info(orders)")
columns = [col[1] for col in cursor.fetchall()]
if 'agent' not in columns:
    cursor.execute("ALTER TABLE orders ADD COLUMN agent TEXT")
conn.commit()

class AutocompleteEntry(tk.Entry):
    def __init__(self, master, fetch_suggestions, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.fetch_suggestions = fetch_suggestions
        self.listbox = None
        self.bind("<KeyRelease>", self.on_keyrelease)
        self.bind("<Down>", self.move_down)
        self.bind("<Up>", self.move_up)
        self.bind("<Return>", self.fill_selection)
        self.bind("<Escape>", lambda e: self.hide_suggestions())
        self.bind("<FocusOut>", self.delayed_hide)

    def on_keyrelease(self, event):
        if event.keysym in ("Down", "Up", "Return", "Tab", "Escape"):
            return
        value = self.get()
        if value == "":
            self.hide_suggestions()
            return
        suggestions = self.fetch_suggestions(value)
        if suggestions:
            if not self.listbox:
                self.listbox = tk.Listbox(self.master, height=5)
                self.listbox.bind("<Return>", self.fill_selection)
                self.listbox.bind("<Double-Button-1>", self.fill_selection)
                self.listbox.bind("<Escape>", lambda e: self.hide_suggestions())
                self.listbox.bind("<Up>", self.listbox_move_up)
                self.listbox.bind("<Down>", self.listbox_move_down)
                x = self.winfo_x()
                y = self.winfo_y() + self.winfo_height()
                self.listbox.place(x=x, y=y, width=self.winfo_width())
            self.listbox.delete(0, tk.END)
            for item in suggestions:
                self.listbox.insert(tk.END, item.title())
        else:
            self.hide_suggestions()

    def move_down(self, event):
        if self.listbox:
            current = self.listbox.curselection()
            if not current:
                self.listbox.selection_set(0)
                self.listbox.activate(0)
            else:
                next_idx = current[0] + 1
                if next_idx < self.listbox.size():
                    self.listbox.selection_clear(0, tk.END)
                    self.listbox.selection_set(next_idx)
                    self.listbox.activate(next_idx)
            self.listbox.focus()
            return "break"

    def move_up(self, event):
        if self.listbox:
            current = self.listbox.curselection()
            if not current:
                self.listbox.selection_set(tk.END)
                self.listbox.activate(tk.END)
            else:
                prev_idx = current[0] - 1
                if prev_idx >= 0:
                    self.listbox.selection_clear(0, tk.END)
                    self.listbox.selection_set(prev_idx)
                    self.listbox.activate(prev_idx)
            self.listbox.focus()
            return "break"

    def listbox_move_down(self, event):
        current = self.listbox.curselection()
        if not current:
            self.listbox.selection_set(0)
            self.listbox.activate(0)
        else:
            next_idx = current[0] + 1
            if next_idx < self.listbox.size():
                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(next_idx)
                self.listbox.activate(next_idx)
        return "break"

    def listbox_move_up(self, event):
        current = self.listbox.curselection()
        if not current:
            self.listbox.selection_set(tk.END)
            self.listbox.activate(tk.END)
        else:
            prev_idx = current[0] - 1
            if prev_idx >= 0:
                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(prev_idx)
                self.listbox.activate(prev_idx)
        return "break"

    def fill_selection(self, event=None):
        if self.listbox:
            selection = self.listbox.curselection()
            if selection:
                value = self.listbox.get(selection[0])
                self.delete(0, tk.END)
                self.insert(0, value)
            self.hide_suggestions()
            self.focus_set()
        return "break"

    def hide_suggestions(self):
        if self.listbox:
            self.listbox.destroy()
            self.listbox = None

    def delayed_hide(self, event=None):
        self.after(100, self._check_focus)

    def _check_focus(self):
        try:
            focused = self.focus_get()
            if focused and focused not in (self, self.listbox):
                self.hide_suggestions()
        except KeyError:
            self.hide_suggestions()

def fetch_suggestions(table, column, text):
    cursor.execute(f"SELECT DISTINCT {column} FROM {table} WHERE {column} LIKE ? ORDER BY {column} LIMIT 10", (text.lower() + '%',))
    return [row[0] for row in cursor.fetchall()]

def save_last_bale_no(bale_no):
    """Save the last bale number to a file."""
    try:
        with open("last_bale_no.txt", "w", encoding='utf-8') as f:
            f.write(bale_no)
    except Exception as e:
        print(f"Failed to save last bale number: {e}")

def load_last_bale_no():
    """Load the last bale number from a file."""
    try:
        if os.path.exists("last_bale_no.txt"):
            with open("last_bale_no.txt", "r", encoding='utf-8') as f:
                return f.read().strip()
        return ""
    except Exception as e:
        print(f"Failed to load last bale number: {e}")
        return ""

def add_item(table, entries, tree):
    item = entries['item'].get().strip().lower()
    ref_no = entries['ref_no'].get().strip()
    qty_input = entries['quantity'].get().strip()
    entry_date = entries['date'].get_date().isoformat()
    entity = entries['entity'].get().strip().lower()

    if item == '*':
        for entry in [entries['ref_no'], entries['entity'], entries['item'], entries['quantity']]:
            entry.delete(0, tk.END)
        if table == 'outward_inventory' and ref_no:
            try:
                match = re.search(r'(\d+)$', ref_no)
                if match:
                    num = int(match.group(1))
                    prefix = ref_no[:match.start()]
                    new_bale_no = f"{prefix}{num + 1:03d}"
                    entries['ref_no'].insert(0, new_bale_no)
                    save_last_bale_no(new_bale_no)
                else:
                    entries['ref_no'].insert(0, ref_no)
                    save_last_bale_no(ref_no)
            except:
                entries['ref_no'].insert(0, ref_no)
                save_last_bale_no(ref_no)
        entries['ref_no'].focus_set()
        return False

    if not qty_input:
        messagebox.showerror("Invalid input", "Quantity required")
        entries['quantity'].focus_set()
        return False

    match = re.fullmatch(r"-?\d+[a-zA-Z]*", qty_input)
    if not match:
        messagebox.showerror("Invalid input", "Quantity must be a number (including 0) or number+unit like '0d' or '-2d'")
        entries['quantity'].focus_set()
        return False

    num = int(qty_input.rstrip('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'))
    unit = qty_input.lstrip('-0123456789').lower()
    if unit and unit not in unit_multipliers:
        messagebox.showerror("Unknown Unit", f"Unit '{unit}' is not defined.")
        entries['quantity'].focus_set()
        return False
    qty = num * unit_multipliers.get(unit, 1)

    ref_col = 'memo_no' if table == 'inward_inventory' else 'bale_no'
    entity_col = 'merchant' if table == 'inward_inventory' else 'transport'
    
    # Record the action for undo before making changes
    action_data = {
        'item_name': item,
        'ref_no': ref_no,
        'quantity': qty,
        'date': entry_date,
        'entity': entity
    }
    
    cursor.execute(f"SELECT id, quantity FROM {table} WHERE item_name = ? AND date = ? AND {entity_col} = ? AND {ref_col} = ?", 
                   (item, entry_date, entity, ref_no))
    entry = cursor.fetchone()

    if entry:
        entry_id, current_qty = entry
        cursor.execute(f"UPDATE {table} SET quantity = ? WHERE id = ?", (current_qty + qty, entry_id))
    else:
        cursor.execute(f"INSERT INTO {table} (item_name, {ref_col}, quantity, date, {entity_col}) VALUES (?, ?, ?, ?, ?)", 
                       (item, ref_no, qty, entry_date, entity))

    if table == 'outward_inventory' and ref_no:
        save_last_bale_no(ref_no)

    conn.commit()
    
    # Record the action after successful commit
    record_action('add_item', table, action_data)
    
    update_tree(table, tree, entries['search'], entries['date_filter'], entries['entity_filter'], entries['ref_filter'], entries['show_date_var'])
    
    if unit:
        messagebox.showinfo("Converted Quantity", f"Interpreted '{qty_input}' as {qty} pcs.")
    
    entries['item'].delete(0, tk.END)
    entries['quantity'].delete(0, tk.END)
    return True

def add_order(entries, tree):
    party = entries['party'].get().strip().lower()
    agent = entries['agent'].get().strip().lower()
    quality = entries['quality'].get().strip().lower()
    qty_input = entries['quantity'].get().strip()
    entry_date = entries['date'].get_date().isoformat()

    if not party or not quality or not qty_input:
        messagebox.showerror("Invalid Input", "Party, quality, and quantity are required")
        return False

    match = re.fullmatch(r"-?\d+[a-zA-Z]*", qty_input)
    if not match:
        messagebox.showerror("Invalid input", "Quantity must be a number (including 0) or number+unit like '0d' or '-2d'")
        entries['quantity'].focus_set()
        return False

    num = int(qty_input.rstrip('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'))
    unit = qty_input.lstrip('-0123456789').lower()
    if unit and unit not in unit_multipliers:
        messagebox.showerror("Unknown Unit", f"Unit '{unit}' is not defined.")
        entries['quantity'].focus_set()
        return False
    qty = num * unit_multipliers.get(unit, 1)

    # Record action for undo
    action_data = {
        'party': party,
        'agent': agent,
        'quality': quality,
        'quantity': qty,
        'date': entry_date
    }

    cursor.execute("SELECT id, quantity FROM orders WHERE party_name = ? AND quality_name = ? AND date = ? AND agent = ?", 
                   (party, quality, entry_date, agent))
    entry = cursor.fetchone()

    if entry:
        entry_id, current_qty = entry
        cursor.execute("UPDATE orders SET quantity = ? WHERE id = ?", (current_qty + qty, entry_id))
    else:
        cursor.execute("INSERT INTO orders (party_name, quality_name, quantity, date, agent) VALUES (?, ?, ?, ?, ?)", 
                       (party, quality, qty, entry_date, agent))

    conn.commit()
    
    # Record the action
    record_action('add_order', 'orders', action_data)
    
    update_orders_tree(tree, entries['search'], entries['date_filter'], entries['show_date_var'])
    
    if unit:
        messagebox.showinfo("Converted Quantity", f"Interpreted '{qty_input}' as {qty} pcs.")
    
    return True

def update_tree(table, tree, search_entry, date_picker, entity_filter, ref_filter, show_date_var=None):
    for row in tree.get_children():
        tree.delete(row)
    query = f"SELECT item_name, SUM(quantity) FROM {table}"
    conditions = []
    params = []
    if show_date_var and show_date_var.get():
        conditions.append("date = ?")
        params.append(date_picker.get_date().isoformat())
    search_term = search_entry.get().strip().lower()
    if search_term:
        conditions.append("item_name LIKE ?")
        params.append('%' + search_term + '%')
    entity_term = entity_filter.get().strip().lower()
    if entity_term:
        entity_col = 'merchant' if table == 'inward_inventory' else 'transport'
        conditions.append(f"{entity_col} LIKE ?")
        params.append('%' + entity_term + '%')
    ref_term = ref_filter.get().strip().lower()
    if ref_term:
        ref_col = 'memo_no' if table == 'inward_inventory' else 'bale_no'
        conditions.append(f"{ref_col} LIKE ?")
        params.append('%' + ref_term + '%')
    if conditions:
        query += " WHERE " + " AND ".join(conditions)
    query += " GROUP BY item_name ORDER BY item_name"
    cursor.execute(query, params)
    for item, total in cursor.fetchall():
        tree.insert('', 'end', values=(item.title(), f"{total} pcs"))

def update_orders_tree(tree, search_entry, date_picker, show_date_var):
    for row in tree.get_children():
        tree.delete(row)
    query = "SELECT party_name, agent, quality_name, SUM(quantity) FROM orders"
    conditions = []
    params = []
    if show_date_var.get():
        conditions.append("date = ?")
        params.append(date_picker.get_date().isoformat())
    search_term = search_entry.get().strip().lower()
    if search_term:
        conditions.append("party_name LIKE ?")
        params.append('%' + search_term + '%')
    if conditions:
        query += " WHERE " + " AND ".join(conditions)
    query += " GROUP BY party_name, agent, quality_name ORDER BY party_name, agent, quality_name"
    cursor.execute(query, params)
    data = cursor.fetchall()
    if not data:
        tree["columns"] = ("Party / Agent", "Total")
        tree.heading("Party / Agent", text="Party / Agent")
        tree.column("Party / Agent", width=250)
        tree.heading("Total", text="Total")
        tree.column("Total", width=100, anchor='center')
        return
    party_agents = sorted(set((row[0], row[1]) for row in data))
    qualities = sorted(set(row[2] for row in data))
    columns = ["Party / Agent"] + [q.title() for q in qualities] + ["Total"]
    tree["columns"] = columns
    tree.heading("Party / Agent", text="Party / Agent")
    tree.column("Party / Agent", width=250)
    for q in qualities:
        tree.heading(q.title(), text=q.title())
        tree.column(q.title(), width=100, anchor='center')
    tree.heading("Total", text="Total")
    tree.column("Total", width=100, anchor='center')
    from collections import defaultdict
    pivot = defaultdict(lambda: defaultdict(int))
    for party, agent, quality, qty in data:
        key = (party, agent)
        pivot[key][quality] = qty
    for pa in party_agents:
        party, agent = pa
        display_party = party.title()
        if agent:
            display_party += f" ({agent.title()})"
        values = [display_party]
        row_total = 0
        for q in qualities:
            qty = pivot[pa][q]
            values.append(f"{qty} pcs" if qty else "")
            row_total += qty
        values.append(f"{row_total} pcs" if row_total else "")
        tree.insert('', 'end', values=values)
    quality_totals = defaultdict(int)
    grand_total = 0
    for pa in party_agents:
        for q in qualities:
            quality_totals[q] += pivot[pa][q]
    total_values = ["Total"]
    for q in qualities:
        q_total = quality_totals[q]
        total_values.append(f"{q_total} pcs" if q_total else "")
        grand_total += q_total
    total_values.append(f"{grand_total} pcs" if grand_total else "")
    tree.insert('', 'end', values=total_values)

def update_transfer_tree(tree, date_picker, show_date_var):
    for row in tree.get_children():
        tree.delete(row)
    query = "SELECT source_item, quantity, dest_item, date FROM stock_transfer"
    params = []
    if show_date_var.get():
        query += " WHERE date = ?"
        params.append(date_picker.get_date().isoformat())
    query += " ORDER BY date DESC, source_item"
    cursor.execute(query, params)
    for source_item, qty, dest_item, transfer_date in cursor.fetchall():
        tree.insert('', 'end', values=(source_item.title(), f"{qty} pcs", dest_item.title(), transfer_date))

def update_summary_tree(tree, search_entry, date_picker, show_date_var):
    for row in tree.get_children():
        tree.delete(row)

    inward_query = """
        SELECT item_name, SUM(quantity)
        FROM inward_inventory
    """

    outward_query = """
        SELECT item_name, SUM(quantity)
        FROM outward_inventory
    """

    inward_conditions = []
    outward_conditions = []
    inward_params = []
    outward_params = []

    if show_date_var.get():
        selected_date = date_picker.get_date().isoformat()
        inward_conditions.append("date = ?")
        outward_conditions.append("date = ?")
        inward_params.append(selected_date)
        outward_params.append(selected_date)

    search_term = search_entry.get().strip().lower()
    if search_term:
        inward_conditions.append("item_name LIKE ?")
        outward_conditions.append("item_name LIKE ?")
        inward_params.append('%' + search_term + '%')
        outward_params.append('%' + search_term + '%')

    if inward_conditions:
        inward_query += " WHERE " + " AND ".join(inward_conditions)

    if outward_conditions:
        outward_query += " WHERE " + " AND ".join(outward_conditions)

    inward_query += " GROUP BY item_name"
    outward_query += " GROUP BY item_name"

    cursor.execute(inward_query, inward_params)
    inward_data = dict(cursor.fetchall())

    cursor.execute(outward_query, outward_params)
    outward_data = dict(cursor.fetchall())

    all_items = set(inward_data.keys()) | set(outward_data.keys())

    for item in sorted(all_items):
        inward_qty = inward_data.get(item, 0)
        outward_qty = outward_data.get(item, 0)
        net_qty = inward_qty - outward_qty
        tree.insert('', 'end', values=(item.title(), f"{net_qty} pcs"))

def delete_selected_item(table, tree, search_entry, date_picker, entity_filter, ref_filter, show_date_var):
    selected = tree.focus()
    if not selected:
        messagebox.showwarning("No selection", "Please select an item to delete.")
        return
    item_name = tree.item(selected)['values'][0].lower()
    confirm = messagebox.askyesno("Confirm Delete", f"Delete all records for '{item_name}'?")
    if confirm:
        cursor.execute(f"DELETE FROM {table} WHERE item_name = ?", (item_name,))
        conn.commit()
        update_tree(table, tree, search_entry, date_picker, entity_filter, ref_filter, show_date_var)

def edit_item_name(table, tree, search_entry, date_picker, entity_filter, ref_filter, show_date_var):
    selected = tree.focus()
    if not selected:
        messagebox.showwarning("No selection", "Please select an item to edit.")
        return
    old_item_name = tree.item(selected)['values'][0].lower()
    new_item_name = simpledialog.askstring("Edit Item Name", f"Enter new name for '{old_item_name.title()}':", initialvalue=old_item_name.title())
    if new_item_name:
        new_item_name = new_item_name.strip().lower()
        if new_item_name and new_item_name != old_item_name:
            cursor.execute(f"UPDATE {table} SET item_name = ? WHERE item_name = ?", (new_item_name, old_item_name))
            conn.commit()
            update_tree(table, tree, search_entry, date_picker, entity_filter, ref_filter, show_date_var)
            messagebox.showinfo("Success", f"Item name updated from '{old_item_name.title()}' to '{new_item_name.title()}'.")

def search_items(table, tree, search_entry, date_picker, entity_filter, ref_filter, show_date_var):
    update_tree(table, tree, search_entry, date_picker, entity_filter, ref_filter, show_date_var)

def search_orders(tree, search_entry, date_picker, show_date_var):
    update_orders_tree(tree, search_entry, date_picker, show_date_var)

def search_summary(tree, search_entry, date_picker, show_date_var):
    update_summary_tree(tree, search_entry, date_picker, show_date_var)

def search_transfers(tree, date_picker, show_date_var):
    """Update transfer tree with optional date filter"""
    for row in tree.get_children():
        tree.delete(row)
    
    query = "SELECT source_item, quantity, dest_item, date FROM stock_transfer"
    params = []
    
    if show_date_var.get():
        query += " WHERE date = ?"
        params.append(date_picker.get_date().isoformat())
    
    query += " ORDER BY date DESC, source_item"
    
    cursor.execute(query, params)
    rows = cursor.fetchall()
    
    if not rows:
        tree.insert('', 'end', values=("No transfers found", "", "", ""))
        return
    
    for source_item, qty, dest_item, transfer_date in rows:
        # Format date for display
        try:
            formatted_date = datetime.strptime(transfer_date, '%Y-%m-%d').strftime('%d/%m/%Y')
        except:
            formatted_date = transfer_date
        
        # Format quantity display
        if qty > 0:
            qty_display = f"+{qty} pcs"
        else:
            qty_display = f"{qty} pcs"
        
        tree.insert('', 'end', values=(source_item.title(), qty_display, dest_item.title(), formatted_date))

def load_selected_item(table, entries, tree):
    selected = tree.focus()
    if not selected:
        return
    item_name = tree.item(selected)['values'][0].lower()
    ref_col = 'memo_no' if table == 'inward_inventory' else 'bale_no'
    entity_col = 'merchant' if table == 'inward_inventory' else 'transport'
    cursor.execute(f"SELECT {ref_col}, {entity_col} FROM {table} WHERE item_name = ? LIMIT 1", (item_name,))
    result = cursor.fetchone()
    if result:
        ref_no, entity = result
        entries['item'].delete(0, tk.END)
        entries['item'].insert(0, item_name.title())
        entries['ref_no'].delete(0, tk.END)
        entries['ref_no'].insert(0, ref_no if ref_no else "")
        entries['entity'].delete(0, tk.END)
        entries['entity'].insert(0, entity.title() if entity else "")
        entries['quantity'].focus_set()

def update_quantity_from_context(table, entries, tree, search_entry, date_picker, entity_filter, ref_filter, show_date_var):
    selected = tree.focus()
    if not selected:
        messagebox.showwarning("No Selection", "Please select an item to update.")
        return

    item_name = tree.item(selected)['values'][0].lower()
    qty_input = entries['quantity'].get().strip()
    ref_no = entries['ref_no'].get().strip()
    entity = entries['entity'].get().strip().lower()
    entry_date = entries['date'].get_date().isoformat()

    if not qty_input:
        messagebox.showerror("Invalid Input", "Please enter a quantity.")
        entries['quantity'].focus_set()
        return

    match = re.fullmatch(r"-?\d+[a-zA-Z]*", qty_input)
    if not match:
        messagebox.showerror("Invalid Input", "Quantity must be a number (including 0) or number+unit like '0d' or '-2d'")
        entries['quantity'].focus_set()
        return

    num = int(qty_input.rstrip('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'))
    unit = qty_input.lstrip('-0123456789').lower()
    if unit and unit not in unit_multipliers:
        messagebox.showerror("Unknown Unit", f"Unit '{unit}' is not defined.")
        entries['quantity'].focus_set()
        return
    qty = num * unit_multipliers.get(unit, 1)

    ref_col = 'memo_no' if table == 'inward_inventory' else 'bale_no'
    entity_col = 'merchant' if table == 'inward_inventory' else 'transport'
    cursor.execute(f"SELECT id FROM {table} WHERE item_name = ? AND {ref_col} = ? AND date = ? AND {entity_col} = ?", 
                   (item_name, ref_no, entry_date, entity))
    result = cursor.fetchone()

    if result:
        item_id = result[0]
        cursor.execute(f"UPDATE {table} SET quantity = ? WHERE id = ?", (qty, item_id))
    else:
        cursor.execute(f"INSERT INTO {table} (item_name, {ref_col}, quantity, date, {entity_col}) VALUES (?, ?, ?, ?, ?)", 
                       (item_name, ref_no, qty, entry_date, entity))

    if table == 'outward_inventory' and ref_no:
        save_last_bale_no(ref_no)

    cursor.execute(f"SELECT SUM(quantity) FROM {table} WHERE item_name = ? AND {entity_col} = ?", (item_name, entity))
    total_qty = cursor.fetchone()[0]

    if total_qty <= 0:
        cursor.execute(f"DELETE FROM {table} WHERE item_name = ? AND {entity_col} = ?", (item_name, entity))
        messagebox.showinfo("Item Removed", f"'{item_name.title()}' from {entity_col} '{entity.title()}' removed due to zero or negative quantity.")

    conn.commit()
    update_tree(table, tree, search_entry, date_picker, entity_filter, ref_filter, show_date_var)
    if unit:
        messagebox.showinfo("Converted Quantity", f"Interpreted '{qty_input}' as {qty} pcs.")
    entries['quantity'].delete(0, tk.END)
    entries['item'].focus_set()

def export_to_xlsx(table, date_format='%d/%m/%y'):
    if table == 'inward_inventory':
        cursor.execute("""
            SELECT date, item_name, quantity, memo_no, merchant 
            FROM inward_inventory 
            ORDER BY memo_no, merchant, item_name
        """)
        rows = cursor.fetchall()
        if not rows:
            messagebox.showinfo("No data", "Inward inventory is empty, nothing to export.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")],
                                                 initialfile=f"inward_export_{date.today().isoformat()}.xlsx")
        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Inward Export"

            bold_font = Font(bold=True)
            border_top = Border(top=Side(style='thin'))
            centered = Alignment(horizontal='center')

            ws.column_dimensions['A'].width = 25
            ws.column_dimensions['B'].width = 12

            ws.page_setup.paperSize = ws.PAPERSIZE_A5
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.page_margins.left = 0.2
            ws.page_margins.right = 0.2
            ws.page_margins.top = 0.2
            ws.page_margins.bottom = 0.2

            current_row = 1
            current_memo = None
            current_merchant = None
            group_rows = []
            total_qty = 0

            for row in rows:
                memo_no = row[3] if row[3] else "No Memo No"
                merchant = row[4] if row[4] else "No Merchant"
                if memo_no != current_memo or merchant != current_merchant:
                    if group_rows:
                        ws.cell(row=current_row, column=1, value=f"inward_memo_{current_memo}_{date.today().isoformat()}")
                        ws.cell(row=current_row, column=1).font = bold_font
                        ws.cell(row=current_row, column=1).alignment = centered
                        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
                        current_row += 1
                        try:
                            parsed_date = datetime.strptime(group_rows[0][0], '%Y-%m-%d')
                            formatted_date = parsed_date.strftime(date_format)
                        except ValueError:
                            formatted_date = group_rows[0][0]
                        ws.cell(row=current_row, column=1, value="Date")
                        ws.cell(row=current_row, column=2, value=formatted_date)
                        current_row += 1
                        ws.cell(row=current_row, column=1, value="Memo No")
                        ws.cell(row=current_row, column=2, value=current_memo)
                        current_row += 1
                        ws.cell(row=current_row, column=1, value="Merchant")
                        ws.cell(row=current_row, column=2, value=current_merchant.title() if current_merchant != "No Merchant" else "")
                        current_row += 1
                        current_row += 1
                        ws.cell(row=current_row, column=1, value="Item Name").font = bold_font
                        ws.cell(row=current_row, column=2, value="Quantity").font = bold_font
                        current_row += 1
                        for group_row in group_rows:
                            ws.cell(row=current_row, column=1, value=group_row[1].title())
                            ws.cell(row=current_row, column=2, value=f"{group_row[2]} pcs")
                            current_row += 1
                        ws.cell(row=current_row, column=2).border = border_top
                        ws.cell(row=current_row, column=1, value="")
                        ws.cell(row=current_row, column=2, value=f"Total  {total_qty} pcs").font = bold_font
                        current_row += 2
                    current_memo = memo_no
                    current_merchant = merchant
                    group_rows = []
                    total_qty = 0
                group_rows.append(row)
                total_qty += row[2]

            if group_rows:
                ws.cell(row=current_row, column=1, value=f"inward_memo_{current_memo}_{date.today().isoformat()}")
                ws.cell(row=current_row, column=1).font = bold_font
                ws.cell(row=current_row, column=1).alignment = centered
                ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
                current_row += 1
                try:
                    parsed_date = datetime.strptime(group_rows[0][0], '%Y-%m-%d')
                    formatted_date = parsed_date.strftime(date_format)
                except ValueError:
                    formatted_date = group_rows[0][0]
                ws.cell(row=current_row, column=1, value="Date")
                ws.cell(row=current_row, column=2, value=formatted_date)
                current_row += 1
                ws.cell(row=current_row, column=1, value="Memo No")
                ws.cell(row=current_row, column=2, value=current_memo)
                current_row += 1
                ws.cell(row=current_row, column=1, value="Merchant")
                ws.cell(row=current_row, column=2, value=current_merchant.title() if current_merchant != "No Merchant" else "")
                current_row += 1
                current_row += 1
                ws.cell(row=current_row, column=1, value="Item Name").font = bold_font
                ws.cell(row=current_row, column=2, value="Quantity").font = bold_font
                current_row += 1
                for group_row in group_rows:
                    ws.cell(row=current_row, column=1, value=group_row[1].title())
                    ws.cell(row=current_row, column=2, value=f"{group_row[2]} pcs")
                    current_row += 1
                ws.cell(row=current_row, column=2).border = border_top
                ws.cell(row=current_row, column=1, value="")
                ws.cell(row=current_row, column=2, value=f"Total  {total_qty} pcs").font = bold_font

            wb.save(file_path)
            messagebox.showinfo("Export Successful", f"Inward inventory exported to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export Excel file:\n{e}")

    elif table == 'stock_transfer':
        cursor.execute("SELECT date, source_item, quantity, dest_item FROM stock_transfer ORDER BY date DESC, source_item")
        rows = cursor.fetchall()
        headers = ['Date', 'Source Item', 'Quantity', 'Destination Item']
        initial_file = f"stock_transfer_export_{date.today().isoformat()}.xlsx"
        if not rows:
            messagebox.showinfo("No data", "Stock Transfer is empty, nothing to export.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")],
                                                 initialfile=initial_file)
        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Stock Transfer"

            bold_font = Font(bold=True)

            ws.column_dimensions['A'].width = 12
            ws.column_dimensions['B'].width = 20
            ws.column_dimensions['C'].width = 12            
            ws.column_dimensions['D'].width = 20

            ws.page_setup.paperSize = ws.PAPERSIZE_A5
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.page_margins.left = 0.2
            ws.page_margins.right = 0.2
            ws.page_margins.top = 0.2
            ws.page_margins.bottom = 0.2

            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = bold_font

            for row_idx, row in enumerate(rows, 2):
                try:
                    parsed_date = datetime.strptime(row[0], '%Y-%m-%d')
                    formatted_date = parsed_date.strftime(date_format)
                except ValueError:
                    formatted_date = row[0]
                ws.cell(row=row_idx, column=1, value=formatted_date)
                ws.cell(row=row_idx, column=2, value=row[1].title() if row[1] else "")
                ws.cell(row=row_idx, column=3, value=row[2] if row[2] is not None else "")
                ws.cell(row=row_idx, column=4, value=row[3].title() if row[3] else "")

            wb.save(file_path)
            messagebox.showinfo("Export Successful", f"Stock Transfer exported to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export Excel file:\n{e}")

    else:
        cursor.execute("SELECT date, item_name, quantity, bale_no, transport FROM outward_inventory ORDER BY date, item_name")
        rows = cursor.fetchall()
        headers = ['Date', 'Item Name', 'Quantity', 'Bale No', 'Transport']
        initial_file = f"outward_export_{date.today().isoformat()}.xlsx"
        if not rows:
            messagebox.showinfo("No data", "Outward inventory is empty, nothing to export.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")],
                                                 initialfile=initial_file)
        if not file_path:
            return

        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Outward Export"

            bold_font = Font(bold=True)
            ws.column_dimensions['A'].width = 12
            ws.column_dimensions['B'].width = 20
            ws.column_dimensions['C'].width = 12
            ws.column_dimensions['D'].width = 12
            ws.column_dimensions['E'].width = 15

            ws.page_setup.paperSize = ws.PAPERSIZE_A5
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.page_margins.left = 0.2
            ws.page_margins.right = 0.2
            ws.page_margins.top = 0.2
            ws.page_margins.bottom = 0.2

            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = bold_font

            for row_idx, row in enumerate(rows, 2):
                try:
                    parsed_date = datetime.strptime(row[0], '%Y-%m-%d')
                    formatted_date = parsed_date.strftime(date_format)
                except ValueError:
                    formatted_date = row[0]
                ws.cell(row=row_idx, column=1, value=formatted_date)
                ws.cell(row=row_idx, column=2, value=row[1].title() if row[1] else "")
                ws.cell(row=row_idx, column=3, value=row[2] if row[2] is not None else "")
                ws.cell(row=row_idx, column=4, value=row[3] if row[3] else "")
                ws.cell(row=row_idx, column=5, value=row[4].title() if row[4] else "")

            wb.save(file_path)
            messagebox.showinfo("Export Successful", f"Outward inventory exported to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export Excel file:\n{e}")

def export_orders_to_xlsx(search_entry, date_picker, show_date_var, date_format='%d/%m/%y'):
    query = "SELECT party_name, agent, quality_name, SUM(quantity) FROM orders"
    conditions = []
    params = []
    if show_date_var.get():
        conditions.append("date = ?")
        params.append(date_picker.get_date().isoformat())
    search_term = search_entry.get().strip().lower()
    if search_term:
        conditions.append("party_name LIKE ?")
        params.append('%' + search_term + '%')
    if conditions:
        query += " WHERE " + " AND ".join(conditions)
    query += " GROUP BY party_name, agent, quality_name ORDER BY party_name, agent, quality_name"
    cursor.execute(query, params)
    data = cursor.fetchall()
    if not data:
        messagebox.showinfo("No data", "Orders are empty, nothing to export.")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")],
                                             initialfile=f"orders_export_{date.today().isoformat()}.xlsx")
    if not file_path:
        return

    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Orders Export"

        bold_font = Font(bold=True)

        party_agents = sorted(set((row[0], row[1]) for row in data))
        qualities = sorted(set(row[2] for row in data))
        headers = ["Party / Agent"] + [q.title() for q in qualities] + ["Total"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = bold_font
        ws.column_dimensions['A'].width = 25
        for col in range(2, len(headers) + 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 15

        from collections import defaultdict
        pivot = defaultdict(lambda: defaultdict(int))
        for party, agent, quality, qty in data:
            key = (party, agent)
            pivot[key][quality] = qty
        current_row = 2
        for pa in party_agents:
            party, agent = pa
            display_party = party.title()
            if agent:
                display_party += f" ({agent.title()})"
            ws.cell(row=current_row, column=1, value=display_party)
            col = 2
            row_total = 0
            for q in qualities:
                qty = pivot[pa][q]
                ws.cell(row=current_row, column=col, value=qty if qty else "")
                row_total += qty
                col += 1
            ws.cell(row=current_row, column=col, value=row_total if row_total else "")
            current_row += 1
        quality_totals = defaultdict(int)
        grand_total = 0
        for pa in party_agents:
            for q in qualities:
                quality_totals[q] += pivot[pa][q]
        ws.cell(row=current_row, column=1, value="Total").font = bold_font
        col = 2
        for q in qualities:
            q_total = quality_totals[q]
            ws.cell(row=current_row, column=col, value=q_total if q_total else "")
            grand_total += q_total
            col += 1
        ws.cell(row=current_row, column=col, value=grand_total if grand_total else "").font = bold_font

        wb.save(file_path)
        messagebox.showinfo("Export Successful", f"Orders exported to:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to export Excel file:\n{e}")

def print_outward_by_specific_bale_no(entries, tree, bale_info_label, date_format='%d/%m/%y'):
    bale_no = entries['bale_print'].get().strip()
    if not bale_no:
        messagebox.showwarning("No Bale No", "Please enter a bale number to print.")
        return
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")],
                                             initialfile=f"outward_bale_{bale_no}_{date.today().isoformat()}.xlsx")
    if not file_path:
        return
    try:
        cursor.execute("""
            SELECT date, item_name, quantity, bale_no, transport 
            FROM outward_inventory 
            WHERE bale_no = ? 
            ORDER BY item_name
        """, (bale_no,))
        rows = cursor.fetchall()
        if not rows:
            messagebox.showinfo("No Data", f"No records found for bale number '{bale_no}'.")
            bale_info_label.config(text="")
            for row in tree.get_children():
                tree.delete(row)
            return

        total_qty = sum(row[2] for row in rows)
        try:
            parsed_date = datetime.strptime(rows[0][0], '%Y-%m-%d')
            formatted_date = parsed_date.strftime(date_format)
        except ValueError:
            formatted_date = rows[0][0]
        transport = rows[0][4].title() if rows[0][4] else ""
        bale_info_label.config(text=f"Date: {formatted_date}\nBale No: {bale_no}\nTransport: {transport}")

        for row in tree.get_children():
            tree.delete(row)
        for row in rows:
            tree.insert('', 'end', values=(row[1].title(), f"{row[2]} pcs"))

        wb = Workbook()
        ws = wb.active
        ws.title = f"Bale {bale_no}"

        bold_font = Font(name='Calibri', size=22, bold=True)
        regular_font = Font(name='Calibri', size=22)
        border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        border_top = Border(top=Side(style='thin'))
        centered = Alignment(horizontal='center')
        left_aligned = Alignment(horizontal='left')

        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 12

        ws.page_setup.paperSize = ws.PAPERSIZE_A5
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_margins.left = 0.2
        ws.page_margins.right = 0.2
        ws.page_margins.top = 0.2
        ws.page_margins.bottom = 0.2

        current_row = 1

        ws.cell(row=current_row, column=1, value="PAWAN WEAVING MILL")
        ws.cell(row=current_row, column=1).font = bold_font
        ws.cell(row=current_row, column=1).alignment = centered
        ws.cell(row=current_row, column=2).font = regular_font
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
        ws.cell(row=current_row, column=1).border = border_thin
        ws.cell(row=current_row, column=2).border = border_thin
        current_row += 1

        ws.cell(row=current_row, column=1, value="223 West Managalwar Peth , Solapur")
        ws.cell(row=current_row, column=1).font = regular_font
        ws.cell(row=current_row, column=1).alignment = centered
        ws.cell(row=current_row, column=2).font = regular_font
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
        ws.cell(row=current_row, column=1).border = border_thin
        ws.cell(row=current_row, column=2).border = border_thin
        current_row += 1

        ws.cell(row=current_row, column=1, value="0217-2733010")
        ws.cell(row=current_row, column=1).font = regular_font
        ws.cell(row=current_row, column=1).alignment = centered
        ws.cell(row=current_row, column=2).font = regular_font
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
        ws.cell(row=current_row, column=1).border = border_thin
        ws.cell(row=current_row, column=2).border = border_thin
        current_row += 1

        ws.cell(row=current_row, column=1, value="")
        ws.cell(row=current_row, column=1).font = regular_font
        ws.cell(row=current_row, column=1).border = border_thin
        ws.cell(row=current_row, column=2, value="")
        ws.cell(row=current_row, column=2).font = regular_font
        ws.cell(row=current_row, column=2).border = border_thin
        current_row += 1

        ws.cell(row=current_row, column=1, value="Date")
        ws.cell(row=current_row, column=1).font = regular_font
        ws.cell(row=current_row, column=1).alignment = left_aligned
        ws.cell(row=current_row, column=1).border = border_thin
        ws.cell(row=current_row, column=2, value=formatted_date)
        ws.cell(row=current_row, column=2).font = regular_font
        ws.cell(row=current_row, column=2).alignment = left_aligned
        ws.cell(row=current_row, column=2).border = border_thin
        current_row += 1

        ws.cell(row=current_row, column=1, value="Bale No")
        ws.cell(row=current_row, column=1).font = regular_font
        ws.cell(row=current_row, column=1).alignment = left_aligned
        ws.cell(row=current_row, column=1).border = border_thin
        ws.cell(row=current_row, column=2, value=bale_no)
        ws.cell(row=current_row, column=2).font = regular_font
        ws.cell(row=current_row, column=2).alignment = left_aligned
        ws.cell(row=current_row, column=2).border = border_thin
        current_row += 1

        ws.cell(row=current_row, column=1, value="Transport")
        ws.cell(row=current_row, column=1).font = regular_font
        ws.cell(row=current_row, column=1).alignment = left_aligned
        ws.cell(row=current_row, column=1).border = border_thin
        ws.cell(row=current_row, column=2, value=transport)
        ws.cell(row=current_row, column=2).font = regular_font
        ws.cell(row=current_row, column=2).alignment = left_aligned
        ws.cell(row=current_row, column=2).border = border_thin
        current_row += 1

        ws.cell(row=current_row, column=1, value="")
        ws.cell(row=current_row, column=1).font = regular_font
        ws.cell(row=current_row, column=1).border = border_thin
        ws.cell(row=current_row, column=2, value="")
        ws.cell(row=current_row, column=2).font = regular_font
        ws.cell(row=current_row, column=2).border = border_thin
        current_row += 1

        ws.cell(row=current_row, column=1, value="Item Name")
        ws.cell(row=current_row, column=1).font = bold_font
        ws.cell(row=current_row, column=1).alignment = left_aligned
        ws.cell(row=current_row, column=1).border = border_thin
        ws.cell(row=current_row, column=2, value="Quantity")
        ws.cell(row=current_row, column=2).font = bold_font
        ws.cell(row=current_row, column=2).alignment = left_aligned
        ws.cell(row=current_row, column=2).border = border_thin
        current_row += 1

        for row in rows:
            ws.cell(row=current_row, column=1, value=row[1].title())
            ws.cell(row=current_row, column=1).font = regular_font
            ws.cell(row=current_row, column=1).alignment = left_aligned
            ws.cell(row=current_row, column=1).border = border_thin
            ws.cell(row=current_row, column=2, value=f"{row[2]} pcs")
            ws.cell(row=current_row, column=2).font = regular_font
            ws.cell(row=current_row, column=2).alignment = left_aligned
            ws.cell(row=current_row, column=2).border = border_thin
            current_row += 1

        ws.cell(row=current_row, column=1).border = border_top
        ws.cell(row=current_row, column=2).border = border_top
        ws.cell(row=current_row, column=1, value="Total")
        ws.cell(row=current_row, column=1).font = bold_font
        ws.cell(row=current_row, column=1).alignment = left_aligned
        ws.cell(row=current_row, column=1).border = border_thin
        ws.cell(row=current_row, column=2, value=f"{total_qty} pcs")
        ws.cell(row=current_row, column=2).font = bold_font
        ws.cell(row=current_row, column=2).alignment = left_aligned
        ws.cell(row=current_row, column=2).border = border_thin

        table_ref = f"A1:B{current_row}"
        table = Table(displayName="BaleTable", ref=table_ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False,
                                             showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        ws.add_table(table)

        wb.save(file_path)
        messagebox.showinfo("Export Successful", f"Outward records for bale number '{bale_no}' exported to:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate Excel file:\n{e}")

def import_from_csv(table, tree, search_entry, date_picker, entity_filter, ref_filter, show_date_var):
    file_path = filedialog.askopenfilename(title="Select CSV file to import", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
    if not file_path:
        return
    try:
        imported_count = 0
        error_count = 0
        errors = []
        ref_col = 'memo_no' if table == 'inward_inventory' else 'bale_no'
        entity_col = 'merchant' if table == 'inward_inventory' else 'transport'
        with open(file_path, mode='r', newline='', encoding='utf-8') as file:
            sample = file.read(1024)
            file.seek(0)
            sniffer = csv.Sniffer()
            has_header = sniffer.has_header(sample)
            reader = csv.reader(file)
            if has_header:
                next(reader)
            for row_num, row in enumerate(reader, start=2 if has_header else 1):
                try:
                    if not any(cell.strip() for cell in row):
                        continue
                    if len(row) < 3:
                        errors.append(f"Row {row_num}: Not enough columns (minimum 3 required)")
                        error_count += 1
                        continue
                    import_date = row[0].strip() if row[0].strip() else date.today().isoformat()
                    item_name = row[1].strip().lower() if len(row) > 1 and row[1].strip() else ""
                    try:
                        quantity = int(row[2].strip()) if len(row) > 2 and row[2].strip() else 0
                    except ValueError:
                        errors.append(f"Row {row_num}: Invalid quantity '{row[2]}'")
                        error_count += 1
                        continue
                    ref_no = row[3].strip() if len(row) > 3 and row[3].strip() else ""
                    entity = row[4].strip().lower() if len(row) > 4 and row[4].strip() else ""
                    try:
                        for fmt in ['%Y-%m-%d', '%d/%m/%y', '%d-%m-%Y', '%m/%d/%Y', '%d/%m/%Y']:
                            try:
                                parsed_date = datetime.strptime(import_date, fmt)
                                import_date = parsed_date.strftime('%Y-%m-%d')
                                break
                            except ValueError:
                                continue
                        else:
                            import_date = date.today().isoformat()
                    except:
                        import_date = date.today().isoformat()
                    if not item_name:
                        errors.append(f"Row {row_num}: Item name is required")
                        error_count += 1
                        continue
                    cursor.execute(f"SELECT id, quantity FROM {table} WHERE item_name = ? AND {ref_col} = ? AND date = ? AND {entity_col} = ?", 
                                   (item_name, ref_no, import_date, entity))
                    existing = cursor.fetchone()
                    if existing:
                        existing_id, existing_qty = existing
                        new_qty = existing_qty + quantity
                        cursor.execute(f"UPDATE {table} SET quantity = ? WHERE id = ?", (new_qty, existing_id))
                    else:
                        cursor.execute(f"INSERT INTO {table} (item_name, {ref_col}, quantity, date, {entity_col}) VALUES (?, ?, ?, ?, ?)",
                                       (item_name, ref_no, quantity, import_date, entity))
                    if table == 'outward_inventory' and ref_no:
                        save_last_bale_no(ref_no)
                    imported_count += 1
                except Exception as e:
                    errors.append(f"Row {row_num}: {str(e)}")
                    error_count += 1
        conn.commit()
        result_message = f"Import completed!\n\nSuccessfully imported: {imported_count} records"
        if error_count > 0:
            result_message += f"\nErrors: {error_count} records"
            if len(errors) <= 10:
                result_message += "\n\nError details:\n" + "\n".join(errors)
            else:
                result_message += f"\n\nShowing first 10 errors:\n" + "\n".join(errors[:10])
                result_message += f"\n... and {len(errors) - 10} more errors"
        messagebox.showinfo("Import Results", result_message)
        if imported_count > 0:
            update_tree(table, tree, search_entry, date_picker, entity_filter, ref_filter, show_date_var)
    except Exception as e:
        messagebox.showerror("Import Error", f"Failed to import CSV file:\n{str(e)}")

def show_import_format_help(table):
    entity = 'Merchant' if table == 'inward_inventory' else 'Transport'
    ref = 'Memo No' if table == 'inward_inventory' else 'Bale No'
    help_text = f"""CSV Import Format Help for {table.replace('_inventory', '').title()}:

Required CSV format:
Date, Item Name, Quantity, {ref}, {entity}

Example:
14/06/25,Rice,50,M001,ABC Store
14/06/25,Wheat,30,M002,XYZ Supplier
15/06/25,Sugar,-20,,ABC Store

Important notes:
- First row can be headers (will be auto-detected)
- Date and Item Name are required
- Date format: DD/MM/YY, YYYY-MM-DD, DD-MM-YYYY, MM/DD/YYYY, or DD/MM/YYYY
- If date is empty or invalid, today's date will be used
- Quantity is required and can be negative (e.g., -20)
- {ref} can be empty
- {entity} can be empty
- If record already exists, quantities will be added together
- Empty rows will be skipped

Supported file encoding: UTF-8"""
    messagebox.showinfo(f"CSV Import Format for {table.replace('_inventory', '').title()}", help_text)

def on_tree_key_press(event, table, tree, search_entry, date_picker, entity_filter, ref_filter, show_date_var, entries):
    selected = tree.focus()
    children = tree.get_children()
    if not children:
        return
    if event.keysym == "Down":
        if selected:
            idx = children.index(selected)
            if idx < len(children) - 1:
                tree.selection_set(children[idx + 1])
                tree.focus(children[idx + 1])
        else:
            tree.selection_set(children[0])
            tree.focus(children[0])
    elif event.keysym == "Up":
        if selected:
            idx = children.index(selected)
            if idx > 0:
                tree.selection_set(children[idx - 1])
                tree.focus(children[idx - 1])
        else:
            tree.selection_set(children[0])
            tree.focus(children[0])
    elif event.keysym == "Return":
        if selected:
            load_selected_item(table, entries, tree)
    elif event.keysym == "Delete":
        delete_selected_item(table, tree, search_entry, date_picker, entity_filter, ref_filter, show_date_var)

def create_tab(tab_frame, table, ref_label, entity_label, inward_tree, outward_tree, summary_tree, inward_entries, outward_entries, summary_entries, date_entries_dict, tab_name):
    entries = {}
    tk.Label(tab_frame, text="Entry Date:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
    entries['date'] = DateEntry(tab_frame, width=18, background='darkblue', foreground='white', borderwidth=2, year=date.today().year, date_pattern='dd/mm/y')
    entries['date'].grid(row=0, column=1, padx=10, pady=5, sticky='w')
    date_entries_dict[tab_name] = entries['date']

    tk.Label(tab_frame, text=f"{ref_label}:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
    entries['ref_no'] = AutocompleteEntry(tab_frame, lambda text: fetch_suggestions(table, 'memo_no' if table == 'inward_inventory' else 'bale_no', text))
    entries['ref_no'].grid(row=1, column=1, padx=10, pady=5, sticky='w')
    entries['ref_no'].bind("<Return>", lambda e: entries['entity'].focus_set())
    if table == 'outward_inventory':
        last_bale_no = load_last_bale_no()
        if last_bale_no:
            entries['ref_no'].delete(0, tk.END)
            entries['ref_no'].insert(0, last_bale_no)

    tk.Label(tab_frame, text=f"{entity_label}:").grid(row=2, column=0, padx=10, pady=5, sticky='e')
    entries['entity'] = AutocompleteEntry(tab_frame, lambda text: fetch_suggestions(table, 'merchant' if table == 'inward_inventory' else 'transport', text))
    entries['entity'].grid(row=2, column=1, padx=10, pady=5, sticky='w')
    entries['entity'].bind("<Return>", lambda e: entries['item'].focus_set())

    tk.Label(tab_frame, text="Item Name:").grid(row=3, column=0, padx=10, pady=5, sticky='e')
    entries['item'] = AutocompleteEntry(tab_frame, lambda text: fetch_suggestions(table, 'item_name', text))
    entries['item'].grid(row=3, column=1, padx=10, pady=5, sticky='w')
    entries['item'].bind("<Return>", lambda e: entries['quantity'].focus_set() if entries['item'].get().strip().lower() != '*' else add_item(table, entries, tree))

    tk.Label(tab_frame, text="Quantity:").grid(row=3, column=2, padx=10, pady=5, sticky='e')
    entries['quantity'] = tk.Entry(tab_frame)
    entries['quantity'].grid(row=3, column=3, padx=10, pady=5, sticky='w')
    entries['quantity'].bind("<Return>", lambda e: [add_item(table, entries, tree), entries['item'].focus_set()] if add_item(table, entries, tree) else None)

    tk.Label(tab_frame, text="Search Item:").grid(row=0, column=4, padx=10, pady=5, sticky='e')
    entries['search'] = tk.Entry(tab_frame)
    entries['search'].grid(row=0, column=5, padx=10, pady=5, sticky='w')
    entries['search'].bind("<KeyRelease>", lambda e: search_items(table, tree, entries['search'], entries['date_filter'], entries['entity_filter'], entries['ref_filter'], entries['show_date_var']))

    tk.Label(tab_frame, text=f"Filter by {entity_label}:").grid(row=1, column=4, padx=10, pady=5, sticky='e')
    entries['entity_filter'] = tk.Entry(tab_frame)
    entries['entity_filter'].grid(row=1, column=5, padx=10, pady=5, sticky='w')
    entries['entity_filter'].bind("<KeyRelease>", lambda e: search_items(table, tree, entries['search'], entries['date_filter'], entries['entity_filter'], entries['ref_filter'], entries['show_date_var']))

    tk.Label(tab_frame, text=f"Search by {ref_label}:").grid(row=2, column=4, padx=10, pady=5, sticky='e')
    entries['ref_filter'] = tk.Entry(tab_frame)
    entries['ref_filter'].grid(row=2, column=5, padx=10, pady=5, sticky='w')
    entries['ref_filter'].bind("<KeyRelease>", lambda e: search_items(table, tree, entries['search'], entries['date_filter'], entries['entity_filter'], entries['ref_filter'], entries['show_date_var']))

    if table == 'outward_inventory':
        tk.Label(tab_frame, text="Print by Bale No:").grid(row=3, column=4, padx=10, pady=5, sticky='e')
        entries['bale_print'] = AutocompleteEntry(tab_frame, lambda text: fetch_suggestions(table, 'bale_no', text))
        entries['bale_print'].grid(row=3, column=5, padx=10, pady=5, sticky='w')
        entries['bale_print'].bind("<Return>", lambda e: print_outward_by_specific_bale_no(entries, tree, entries['bale_info_label']))

    tk.Label(tab_frame, text="Select Date:").grid(row=4, column=0, padx=10, pady=5, sticky='e')
    entries['date_filter'] = DateEntry(tab_frame, width=18, background='darkblue', foreground='white', borderwidth=2, year=date.today().year, date_pattern='dd/mm/y')
    entries['date_filter'].grid(row=4, column=1, padx=10, pady=5, sticky='w')

    entries['show_date_var'] = tk.BooleanVar(value=False)
    tk.Checkbutton(tab_frame, text="Show stock for selected date", variable=entries['show_date_var'], 
                   command=lambda: search_items(table, tree, entries['search'], entries['date_filter'], entries['entity_filter'], entries['ref_filter'], entries['show_date_var'])).grid(row=5, column=1, sticky='w', padx=10)

    def show_today_stock():
        entries['show_date_var'].set(True)
        entries['date_filter'].set_date(date.today())
        search_items(table, tree, entries['search'], entries['date_filter'], entries['entity_filter'], entries['ref_filter'], entries['show_date_var'])

    tk.Button(tab_frame, text="Show Today's Stock", command=show_today_stock).grid(row=5, column=0, pady=5)

    csv_frame = tk.Frame(tab_frame)
    csv_frame.grid(row=6, column=0, columnspan=6, pady=10)
    tk.Button(csv_frame, text="Import from CSV", command=lambda: import_from_csv(table, tree, entries['search'], entries['date_filter'], entries['entity_filter'], entries['ref_filter'], entries['show_date_var']), bg='lightgreen').pack(side=tk.LEFT, padx=5)
    tk.Button(csv_frame, text="Export to Excel", command=lambda: export_to_xlsx(table), bg='lightblue').pack(side=tk.LEFT, padx=5)
    tk.Button(csv_frame, text="CSV Format Help", command=lambda: show_import_format_help(table), bg='lightyellow').pack(side=tk.LEFT, padx=5)

    if table == 'outward_inventory':
        entries['bale_info_label'] = tk.Label(tab_frame, text="", justify=tk.LEFT)
        entries['bale_info_label'].grid(row=7, column=0, columnspan=6, padx=10, pady=5, sticky='w')
        tree_row = 8
    else:
        tree_row = 7

    tree = ttk.Treeview(tab_frame, columns=("Item", "Quantity"), show="headings", selectmode="browse")
    tree.heading("Item", text="Item")
    tree.heading("Quantity", text="Quantity")
    tree.column("Item", width=300)
    tree.column("Quantity", width=150)
    tree.grid(row=tree_row, column=0, columnspan=6, padx=10, pady=10, sticky="nsew")
    tree.bind("<Delete>", lambda e: delete_selected_item(table, tree, entries['search'], entries['date_filter'], entries['entity_filter'], entries['ref_filter'], entries['show_date_var']))
    tree.bind("<Double-1>", lambda e: edit_item_name(table, tree, entries['search'], entries['date_filter'], entries['entity_filter'], entries['ref_filter'], entries['show_date_var']))
    tree.bind("<Key>", lambda e: on_tree_key_press(e, table, tree, entries['search'], entries['date_filter'], entries['entity_filter'], entries['ref_filter'], entries['show_date_var'], entries))

    context_menu = tk.Menu(tab_frame, tearoff=0)
    context_menu.add_command(label="Update Quantity", command=lambda: update_quantity_from_context(table, entries, tree, entries['search'], entries['date_filter'], entries['entity_filter'], entries['ref_filter'], entries['show_date_var']))
    
    def show_context_menu(event):
        selected_item = tree.identify_row(event.y)
        if selected_item:
            tree.selection_set(selected_item)
            context_menu.tk_popup(event.x_root, event.y_root)
    
    tree.bind("<Button-3>", show_context_menu)
    tab_frame.grid_rowconfigure(tree_row, weight=1)
    tab_frame.grid_columnconfigure(1, weight=1)
    tab_frame.grid_columnconfigure(3, weight=1)
    tab_frame.grid_columnconfigure(5, weight=1)

    update_tree(table, tree, entries['search'], entries['date_filter'], entries['entity_filter'], entries['ref_filter'], entries['show_date_var'])
    return tree, entries

def create_orders_tab(tab_frame, date_entries_dict, tab_name):
    entries = {}
    tk.Label(tab_frame, text="Entry Date:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
    entries['date'] = DateEntry(tab_frame, width=18, background='darkblue', foreground='white', borderwidth=2, year=date.today().year, date_pattern='dd/mm/y')
    entries['date'].grid(row=0, column=1, padx=10, pady=5, sticky='w')
    date_entries_dict[tab_name] = entries['date']

    tk.Label(tab_frame, text="Party Name:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
    entries['party'] = AutocompleteEntry(tab_frame, lambda text: fetch_suggestions('orders', 'party_name', text))
    entries['party'].grid(row=1, column=1, padx=10, pady=5, sticky='w')
    entries['party'].bind("<Return>", lambda e: entries['agent'].focus_set())

    tk.Label(tab_frame, text="Agent:").grid(row=2, column=0, padx=10, pady=5, sticky='e')
    entries['agent'] = AutocompleteEntry(tab_frame, lambda text: fetch_suggestions('orders', 'agent', text))
    entries['agent'].grid(row=2, column=1, padx=10, pady=5, sticky='w')
    entries['agent'].bind("<Return>", lambda e: entries['quality'].focus_set())

    tk.Label(tab_frame, text="Quality Name:").grid(row=3, column=0, padx=10, pady=5, sticky='e')
    entries['quality'] = AutocompleteEntry(tab_frame, lambda text: fetch_suggestions('orders', 'quality_name', text))
    entries['quality'].grid(row=3, column=1, padx=10, pady=5, sticky='w')
    def quality_enter(e):
        value = entries['quality'].get().strip().lower()
        if value == '*':
            entries['quality'].delete(0, tk.END)
            entries['quantity'].delete(0, tk.END)
            entries['party'].focus_set()
        else:
            entries['quantity'].focus_set()
    entries['quality'].bind("<Return>", quality_enter)

    tk.Label(tab_frame, text="Quantity:").grid(row=4, column=0, padx=10, pady=5, sticky='e')
    entries['quantity'] = tk.Entry(tab_frame)
    entries['quantity'].grid(row=4, column=1, padx=10, pady=5, sticky='w')
    def quantity_enter(e):
        if add_order(entries, tree):
            entries['quantity'].delete(0, tk.END)
            entries['quality'].focus_set()
    entries['quantity'].bind("<Return>", quantity_enter)

    tk.Label(tab_frame, text="Search by Party:").grid(row=0, column=2, padx=10, pady=5, sticky='e')
    entries['search'] = tk.Entry(tab_frame)
    entries['search'].grid(row=0, column=3, padx=10, pady=5, sticky='w')
    entries['search'].bind("<KeyRelease>", lambda e: search_orders(tree, entries['search'], entries['date_filter'], entries['show_date_var']))

    tk.Label(tab_frame, text="Select Date:").grid(row=1, column=2, padx=10, pady=5, sticky='e')
    entries['date_filter'] = DateEntry(tab_frame, width=18, background='darkblue', foreground='white', borderwidth=2, year=date.today().year, date_pattern='dd/mm/y')
    entries['date_filter'].grid(row=1, column=3, padx=10, pady=5, sticky='w')

    entries['show_date_var'] = tk.BooleanVar(value=False)
    tk.Checkbutton(tab_frame, text="Show orders for selected date", variable=entries['show_date_var'], 
                   command=lambda: search_orders(tree, entries['search'], entries['date_filter'], entries['show_date_var'])).grid(row=2, column=3, sticky='w', padx=10)

    def show_today_orders():
        entries['show_date_var'].set(True)
        entries['date_filter'].set_date(date.today())
        search_orders(tree, entries['search'], entries['date_filter'], entries['show_date_var'])

    tk.Button(tab_frame, text="Show Today's Orders", command=show_today_orders).grid(row=2, column=2, pady=5)

    csv_frame = tk.Frame(tab_frame)
    csv_frame.grid(row=5, column=0, columnspan=4, pady=10)
    tk.Button(csv_frame, text="Export to Excel", command=lambda: export_orders_to_xlsx(entries['search'], entries['date_filter'], entries['show_date_var']), bg='lightblue').pack(side=tk.LEFT, padx=5)

    tree = ttk.Treeview(tab_frame, columns=("Party / Agent",), show="headings", selectmode="browse")
    tree.grid(row=6, column=0, columnspan=4, padx=10, pady=10, sticky="nsew")

    last_clicked_col = None
    def on_tree_click(event):
        nonlocal last_clicked_col
        col = tree.identify_column(event.x)
        if col:
            last_clicked_col = col

    tree.bind("<Button-1>", on_tree_click)

    def on_delete_quality(event):
        if last_clicked_col:
            heading_text = tree.heading(last_clicked_col)['text']
            if heading_text not in ("Party / Agent", "Total"):
                quality = heading_text.lower()
                confirm = messagebox.askyesno("Confirm Delete", f"Delete all records for quality '{heading_text}'?")
                if confirm:
                    cursor.execute("DELETE FROM orders WHERE quality_name = ?", (quality,))
                    conn.commit()
                    update_orders_tree(tree, entries['search'], entries['date_filter'], entries['show_date_var'])

    tree.bind("<Delete>", on_delete_quality)

    tab_frame.grid_rowconfigure(6, weight=1)
    tab_frame.grid_columnconfigure(1, weight=1)
    tab_frame.grid_columnconfigure(3, weight=1)

    update_orders_tree(tree, entries['search'], entries['date_filter'], entries['show_date_var'])
    return tree, entries

def export_summary_to_csv(search_entry, date_picker, show_date_var):
    """Export summary data to CSV file"""
    inward_query = """
        SELECT item_name, SUM(quantity)
        FROM inward_inventory
    """

    outward_query = """
        SELECT item_name, SUM(quantity)
        FROM outward_inventory
    """

    inward_conditions = []
    outward_conditions = []
    inward_params = []
    outward_params = []

    if show_date_var.get():
        selected_date = date_picker.get_date().isoformat()
        inward_conditions.append("date = ?")
        outward_conditions.append("date = ?")
        inward_params.append(selected_date)
        outward_params.append(selected_date)

    search_term = search_entry.get().strip().lower()
    if search_term:
        inward_conditions.append("item_name LIKE ?")
        outward_conditions.append("item_name LIKE ?")
        inward_params.append('%' + search_term + '%')
        outward_params.append('%' + search_term + '%')

    if inward_conditions:
        inward_query += " WHERE " + " AND ".join(inward_conditions)

    if outward_conditions:
        outward_query += " WHERE " + " AND ".join(outward_conditions)

    inward_query += " GROUP BY item_name"
    outward_query += " GROUP BY item_name"

    cursor.execute(inward_query, inward_params)
    inward_data = dict(cursor.fetchall())

    cursor.execute(outward_query, outward_params)
    outward_data = dict(cursor.fetchall())

    all_items = set(inward_data.keys()) | set(outward_data.keys())
    
    if not all_items:
        messagebox.showinfo("No data", "Summary is empty, nothing to export.")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                             filetypes=[("CSV files", "*.csv")],
                                             initialfile=f"summary_export_{date.today().isoformat()}.csv")
    if not file_path:
        return

    try:
        with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Item Name', 'Net Quantity (pcs)'])
            
            for item in sorted(all_items):
                inward_qty = inward_data.get(item, 0)
                outward_qty = outward_data.get(item, 0)
                net_qty = inward_qty - outward_qty
                writer.writerow([item.title(), net_qty])
        
        messagebox.showinfo("Export Successful", f"Summary exported to:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to export CSV file:\n{e}")

def export_summary_to_xlsx(search_entry, date_picker, show_date_var):
    """Export summary data to Excel file"""
    inward_query = """
        SELECT item_name, SUM(quantity)
        FROM inward_inventory
    """

    outward_query = """
        SELECT item_name, SUM(quantity)
        FROM outward_inventory
    """

    inward_conditions = []
    outward_conditions = []
    inward_params = []
    outward_params = []

    if show_date_var.get():
        selected_date = date_picker.get_date().isoformat()
        inward_conditions.append("date = ?")
        outward_conditions.append("date = ?")
        inward_params.append(selected_date)
        outward_params.append(selected_date)

    search_term = search_entry.get().strip().lower()
    if search_term:
        inward_conditions.append("item_name LIKE ?")
        outward_conditions.append("item_name LIKE ?")
        inward_params.append('%' + search_term + '%')
        outward_params.append('%' + search_term + '%')

    if inward_conditions:
        inward_query += " WHERE " + " AND ".join(inward_conditions)

    if outward_conditions:
        outward_query += " WHERE " + " AND ".join(outward_conditions)

    inward_query += " GROUP BY item_name"
    outward_query += " GROUP BY item_name"

    cursor.execute(inward_query, inward_params)
    inward_data = dict(cursor.fetchall())

    cursor.execute(outward_query, outward_params)
    outward_data = dict(cursor.fetchall())

    all_items = set(inward_data.keys()) | set(outward_data.keys())
    
    if not all_items:
        messagebox.showinfo("No data", "Summary is empty, nothing to export.")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")],
                                             initialfile=f"summary_export_{date.today().isoformat()}.xlsx")
    if not file_path:
        return

    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Summary Export"

        bold_font = Font(bold=True)
        
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 20

        ws.cell(row=1, column=1, value="Item Name").font = bold_font
        ws.cell(row=1, column=2, value="Net Quantity (pcs)").font = bold_font

        current_row = 2
        for item in sorted(all_items):
            inward_qty = inward_data.get(item, 0)
            outward_qty = outward_data.get(item, 0)
            net_qty = inward_qty - outward_qty
            ws.cell(row=current_row, column=1, value=item.title())
            ws.cell(row=current_row, column=2, value=net_qty)
            current_row += 1

        wb.save(file_path)
        messagebox.showinfo("Export Successful", f"Summary exported to:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to export Excel file:\n{e}")

def transfer_from_summary(summary_tree, search_entry, date_picker, show_date_var):
    """Transfer stock between items directly from the Summary tab"""
    selected = summary_tree.focus()
    if not selected:
        messagebox.showwarning("No Selection", "Please select an item to transfer from.")
        return
    
    source_item = summary_tree.item(selected)['values'][0].lower()
    
    transfer_dialog = tk.Toplevel()
    transfer_dialog.title(f"Transfer Stock from {source_item.title()}")
    transfer_dialog.geometry("550x550")
    transfer_dialog.transient(summary_tree.master)
    transfer_dialog.grab_set()
    
    # Calculate current net quantity of source item
    cursor.execute("SELECT COALESCE(SUM(quantity), 0) FROM inward_inventory WHERE item_name = ?", (source_item,))
    source_inward = cursor.fetchone()[0]
    cursor.execute("SELECT COALESCE(SUM(quantity), 0) FROM outward_inventory WHERE item_name = ?", (source_item,))
    source_outward = cursor.fetchone()[0]
    source_net = source_inward - source_outward
    
    # Create main frame with scrolling
    main_frame = tk.Frame(transfer_dialog)
    main_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
    # Source information frame
    source_frame = tk.LabelFrame(main_frame, text="Source Item Information", padx=10, pady=5)
    source_frame.pack(fill='x', pady=(0, 10))
    
    tk.Label(source_frame, text="Item Name:", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=5, sticky='e')
    tk.Label(source_frame, text=source_item.title(), font=("Arial", 10, "bold"), fg="blue").grid(row=0, column=1, padx=5, pady=5, sticky='w')
    
    tk.Label(source_frame, text="Current Net Quantity:", font=("Arial", 10)).grid(row=1, column=0, padx=5, pady=5, sticky='e')
    source_net_label = tk.Label(source_frame, text=f"{source_net} pcs", font=("Arial", 10), fg="green" if source_net >= 0 else "red")
    source_net_label.grid(row=1, column=1, padx=5, pady=5, sticky='w')
    
    # Transfer details frame
    transfer_frame = tk.LabelFrame(main_frame, text="Transfer Details", padx=10, pady=5)
    transfer_frame.pack(fill='x', pady=(0, 10))
    
    # Transfer Date
    tk.Label(transfer_frame, text="Transfer Date:", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=10, sticky='e')
    transfer_date = DateEntry(transfer_frame, width=18, background='darkblue', foreground='white', 
                              borderwidth=2, year=date.today().year, date_pattern='dd/mm/y')
    transfer_date.set_date(date.today())
    transfer_date.grid(row=0, column=1, padx=5, pady=10, sticky='w')
    
    # Destination Item
    tk.Label(transfer_frame, text="Destination Item:", font=("Arial", 10, "bold")).grid(row=1, column=0, padx=5, pady=10, sticky='e')
    
    # Get all items from summary for destination
    cursor.execute("""
        SELECT DISTINCT item_name FROM (
            SELECT item_name FROM inward_inventory
            UNION
            SELECT item_name FROM outward_inventory
        ) ORDER BY item_name
    """)
    all_items = [row[0] for row in cursor.fetchall()]
    
    dest_var = tk.StringVar()
    dest_combo = ttk.Combobox(transfer_frame, textvariable=dest_var, values=[item.title() for item in all_items if item != source_item], width=30)
    dest_combo.grid(row=1, column=1, padx=5, pady=10, sticky='w')
    if all_items:
        for item in all_items:
            if item != source_item:
                dest_combo.set(item.title())
                break
    
    # Destination net quantity display
    tk.Label(transfer_frame, text="Destination Net Quantity:", font=("Arial", 10)).grid(row=2, column=0, padx=5, pady=5, sticky='e')
    dest_net_label = tk.Label(transfer_frame, text="", font=("Arial", 10))
    dest_net_label.grid(row=2, column=1, padx=5, pady=5, sticky='w')
    
    def update_dest_net(*args):
        dest_item = dest_var.get().strip().lower()
        if dest_item:
            cursor.execute("SELECT COALESCE(SUM(quantity), 0) FROM inward_inventory WHERE item_name = ?", (dest_item,))
            dest_inward = cursor.fetchone()[0]
            cursor.execute("SELECT COALESCE(SUM(quantity), 0) FROM outward_inventory WHERE item_name = ?", (dest_item,))
            dest_outward = cursor.fetchone()[0]
            dest_net = dest_inward - dest_outward
            dest_net_label.config(text=f"{dest_net} pcs", fg="green" if dest_net >= 0 else "red")
    
    dest_var.trace('w', update_dest_net)
    
    # Quantity input
    tk.Label(transfer_frame, text="Quantity to Transfer:", font=("Arial", 10, "bold")).grid(row=3, column=0, padx=5, pady=10, sticky='e')
    qty_entry = tk.Entry(transfer_frame, width=20)
    qty_entry.grid(row=3, column=1, padx=5, pady=10, sticky='w')
    qty_entry.focus_set()
    
    # Unit selection
    tk.Label(transfer_frame, text="Unit:", font=("Arial", 10)).grid(row=4, column=0, padx=5, pady=5, sticky='e')
    unit_var = tk.StringVar(value="pcs")
    unit_combo = ttk.Combobox(transfer_frame, textvariable=unit_var, values=["pcs", "d (dozen = 12 pcs)"], width=20)
    unit_combo.grid(row=4, column=1, padx=5, pady=5, sticky='w')
    
    # Transfer direction
    tk.Label(transfer_frame, text="Transfer Direction:", font=("Arial", 10, "bold")).grid(row=5, column=0, padx=5, pady=10, sticky='e')
    direction_var = tk.StringVar(value="source_to_dest")
    
    def update_direction_info(*args):
        direction = direction_var.get()
        if direction == "source_to_dest":
            direction_info.config(text=f"Will move quantity FROM {source_item.title()} TO {dest_var.get().title()}")
        else:
            direction_info.config(text=f"Will move quantity FROM {dest_var.get().title()} TO {source_item.title()}")
    
    tk.Radiobutton(transfer_frame, text=f"From {source_item.title()} to Destination", variable=direction_var, 
                   value="source_to_dest", command=update_direction_info).grid(row=5, column=1, padx=5, pady=5, sticky='w')
    tk.Radiobutton(transfer_frame, text=f"From Destination to {source_item.title()}", variable=direction_var, 
                   value="dest_to_source", command=update_direction_info).grid(row=6, column=1, padx=5, pady=5, sticky='w')
    
    direction_info = tk.Label(transfer_frame, text="", font=("Arial", 9), fg="blue")
    direction_info.grid(row=7, column=0, columnspan=2, padx=5, pady=5)
    update_direction_info()
    
    def perform_transfer():
        try:
            dest_item = dest_var.get().strip().lower()
            if not dest_item:
                messagebox.showerror("Invalid Input", "Please select a destination item")
                return
            
            if dest_item == source_item:
                messagebox.showerror("Invalid Input", "Source and destination items cannot be the same")
                return
            
            qty_text = qty_entry.get().strip()
            if not qty_text:
                messagebox.showerror("Invalid Input", "Please enter a quantity")
                return
            
            # Parse quantity with unit
            match = re.fullmatch(r"-?\d+[a-zA-Z]*", qty_text)
            if not match:
                messagebox.showerror("Invalid Input", "Quantity must be a number or number+unit like '10d'")
                return
            
            num = int(qty_text.rstrip('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'))
            unit = qty_text.lstrip('-0123456789').lower()
            if unit == 'd':
                qty_pcs = num * 12
                unit_msg = f" ({num}d = {qty_pcs} pcs)"
            else:
                qty_pcs = num
                unit_msg = ""
            
            direction = direction_var.get()
            entry_date = transfer_date.get_date().isoformat()  # Get the selected date
            
            transfer_qty = abs(qty_pcs)
            
            if direction == "source_to_dest":
                source_change = -transfer_qty
                dest_change = +transfer_qty
                transfer_record_qty = transfer_qty
                transfer_source = source_item
                transfer_dest = dest_item
            else:
                source_change = +transfer_qty
                dest_change = -transfer_qty
                transfer_record_qty = -transfer_qty
                transfer_source = source_item
                transfer_dest = dest_item
            
            # Record the transfer in stock_transfer table with the selected date
            cursor.execute("INSERT INTO stock_transfer (source_item, quantity, dest_item, date) VALUES (?, ?, ?, ?)", 
                           (transfer_source, transfer_record_qty, transfer_dest, entry_date))
            
            # Apply changes to inward/outward inventory
            if source_change > 0:
                cursor.execute("INSERT INTO inward_inventory (item_name, quantity, date, memo_no, merchant) VALUES (?, ?, ?, ?, ?)",
                               (source_item, source_change, entry_date, f"TRANSFER_FROM_{transfer_dest[:20]}", "Transfer"))
            elif source_change < 0:
                cursor.execute("INSERT INTO outward_inventory (item_name, quantity, date, bale_no, transport) VALUES (?, ?, ?, ?, ?)",
                               (source_item, abs(source_change), entry_date, f"TRANSFER_TO_{transfer_dest[:20]}", "Transfer"))
            
            if dest_change > 0:
                cursor.execute("INSERT INTO inward_inventory (item_name, quantity, date, memo_no, merchant) VALUES (?, ?, ?, ?, ?)",
                               (dest_item, dest_change, entry_date, f"TRANSFER_FROM_{transfer_source[:20]}", "Transfer"))
            elif dest_change < 0:
                cursor.execute("INSERT INTO outward_inventory (item_name, quantity, date, bale_no, transport) VALUES (?, ?, ?, ?, ?)",
                               (dest_item, abs(dest_change), entry_date, f"TRANSFER_TO_{transfer_source[:20]}", "Transfer"))
            
            conn.commit()
            
            # Record the transfer action for undo
            action_data = {
                'source_item': transfer_source,
                'dest_item': transfer_dest,
                'quantity': transfer_record_qty,
                'date': entry_date
            }
            record_action('transfer_stock', 'stock_transfer', action_data)
            
            # Refresh the summary tree
            update_summary_tree(summary_tree, search_entry, date_picker, show_date_var)
            
            # Show success message with date
            formatted_date = datetime.strptime(entry_date, '%Y-%m-%d').strftime('%d/%m/%Y')
            
            if direction == "source_to_dest":
                messagebox.showinfo("Transfer Successful", 
                                   f"Successfully transferred {transfer_qty} pcs{unit_msg}\n"
                                   f"From: {source_item.title()}\n"
                                   f"To: {dest_item.title()}\n"
                                   f"Date: {formatted_date}")
            else:
                messagebox.showinfo("Transfer Successful", 
                                   f"Successfully transferred {transfer_qty} pcs{unit_msg}\n"
                                   f"From: {dest_item.title()}\n"
                                   f"To: {source_item.title()}\n"
                                   f"Date: {formatted_date}")
            
            transfer_dialog.destroy()
            
        except Exception as e:
            messagebox.showerror("Transfer Error", f"Failed to transfer stock:\n{str(e)}")
    
    # Buttons
    button_frame = tk.Frame(main_frame)
    button_frame.pack(pady=10)
    tk.Button(button_frame, text="Transfer", command=perform_transfer, bg='lightgreen', font=("Arial", 10, "bold"), padx=20).pack(side=tk.LEFT, padx=10)
    tk.Button(button_frame, text="Cancel", command=transfer_dialog.destroy, bg='lightcoral', padx=20).pack(side=tk.LEFT, padx=10)
    
    # Initial update of destination net
    update_dest_net()

def create_summary_tab(tab_frame, date_entries_dict, tab_name):
    entries = {}
    tk.Label(tab_frame, text="Search Item:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
    entries['search'] = tk.Entry(tab_frame)
    entries['search'].grid(row=0, column=1, padx=10, pady=5, sticky='w')
    entries['search'].bind("<KeyRelease>", lambda e: search_summary(tree, entries['search'], entries['date_filter'], entries['show_date_var']))

    tk.Label(tab_frame, text="Select Date:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
    entries['date_filter'] = DateEntry(tab_frame, width=18, background='darkblue', foreground='white', borderwidth=2, year=date.today().year, date_pattern='dd/mm/y')
    entries['date_filter'].grid(row=1, column=1, padx=10, pady=5, sticky='w')
    date_entries_dict[tab_name] = entries['date_filter']

    entries['show_date_var'] = tk.BooleanVar(value=False)
    tk.Checkbutton(tab_frame, text="Show stock for selected date", variable=entries['show_date_var'], 
                   command=lambda: search_summary(tree, entries['search'], entries['date_filter'], entries['show_date_var'])).grid(row=2, column=1, sticky='w', padx=10)

    def show_today_stock():
        entries['show_date_var'].set(True)
        entries['date_filter'].set_date(date.today())
        search_summary(tree, entries['search'], entries['date_filter'], entries['show_date_var'])

    tk.Button(tab_frame, text="Show Today's Stock", command=show_today_stock).grid(row=2, column=0, pady=5)

    export_frame = tk.Frame(tab_frame)
    export_frame.grid(row=3, column=0, columnspan=2, pady=10)
    tk.Button(export_frame, text="Export to CSV", command=lambda: export_summary_to_csv(entries['search'], entries['date_filter'], entries['show_date_var']), bg='lightgreen').pack(side=tk.LEFT, padx=5)
    tk.Button(export_frame, text="Export to Excel", command=lambda: export_summary_to_xlsx(entries['search'], entries['date_filter'], entries['show_date_var']), bg='lightblue').pack(side=tk.LEFT, padx=5)

    tree = ttk.Treeview(tab_frame, columns=("Item", "Net Quantity"), show="headings", selectmode="browse")
    tree.heading("Item", text="Item")
    tree.heading("Net Quantity", text="Net Quantity")
    tree.column("Item", width=300)
    tree.column("Net Quantity", width=150)
    tree.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
    
    transfer_button = tk.Button(tab_frame, text="Transfer Stock from Selected Item", 
                                command=lambda: transfer_from_summary(tree, entries['search'], entries['date_filter'], entries['show_date_var']),
                                bg='orange', font=("Arial", 10, "bold"), padx=20, pady=5)
    transfer_button.grid(row=5, column=0, columnspan=2, pady=10)
    
    tree.bind("<Double-1>", lambda e: transfer_from_summary(tree, entries['search'], entries['date_filter'], entries['show_date_var']))
    
    context_menu = tk.Menu(tab_frame, tearoff=0)
    context_menu.add_command(label="Transfer Stock", 
                            command=lambda: transfer_from_summary(tree, entries['search'], entries['date_filter'], entries['show_date_var']))
    
    def show_context_menu(event):
        selected_item = tree.identify_row(event.y)
        if selected_item:
            tree.selection_set(selected_item)
            context_menu.tk_popup(event.x_root, event.y_root)
    
    tree.bind("<Button-3>", show_context_menu)

    tab_frame.grid_rowconfigure(4, weight=1)
    tab_frame.grid_columnconfigure(1, weight=1)

    update_summary_tree(tree, entries['search'], entries['date_filter'], entries['show_date_var'])
    return tree, entries

def create_transfer_tab(tab_frame, inward_tree, outward_tree, summary_tree, inward_entries, outward_entries, summary_entries, date_entries_dict, tab_name):
    entries = {}
    
    main_frame = tk.Frame(tab_frame)
    main_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
    # Title
    title_label = tk.Label(main_frame, text="Stock Transfer History", font=("Arial", 14, "bold"))
    title_label.pack(pady=(0, 10))
    
    # Search/Filter Frame
    filter_frame = tk.LabelFrame(main_frame, text="Search & Filter", padx=10, pady=10)
    filter_frame.pack(fill='x', pady=(0, 10))
    
    # Row 0: Date filter
    tk.Label(filter_frame, text="Filter by Date:", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=5, sticky='e')
    entries['date_filter'] = DateEntry(filter_frame, width=18, background='darkblue', foreground='white', 
                                       borderwidth=2, year=date.today().year, date_pattern='dd/mm/y')
    entries['date_filter'].grid(row=0, column=1, padx=5, pady=5, sticky='w')
    
    entries['show_date_var'] = tk.BooleanVar(value=False)
    date_check = tk.Checkbutton(filter_frame, text="Show only selected date", variable=entries['show_date_var'])
    date_check.grid(row=0, column=2, padx=10, pady=5, sticky='w')
    
    # Clear date filter button
    def clear_date_filter():
        entries['show_date_var'].set(False)
        entries['date_filter'].set_date(date.today())
        search_transfers(tree, entries['date_filter'], entries['show_date_var'])
    
    tk.Button(filter_frame, text="Clear Date Filter", command=clear_date_filter, bg='lightyellow').grid(row=0, column=3, padx=5, pady=5)
    
    # Row 1: Item name search
    tk.Label(filter_frame, text="Search by Item Name:", font=("Arial", 10, "bold")).grid(row=1, column=0, padx=5, pady=5, sticky='e')
    entries['item_search'] = tk.Entry(filter_frame, width=30)
    entries['item_search'].grid(row=1, column=1, padx=5, pady=5, sticky='w')
    
    def search_by_item():
        search_term = entries['item_search'].get().strip().lower()
        if search_term:
            # Search in source_item or dest_item
            query = """
                SELECT source_item, quantity, dest_item, date 
                FROM stock_transfer 
                WHERE source_item LIKE ? OR dest_item LIKE ?
                ORDER BY date DESC, source_item
            """
            params = (f'%{search_term}%', f'%{search_term}%')
            
            # Apply date filter if selected
            if entries['show_date_var'].get():
                query = """
                    SELECT source_item, quantity, dest_item, date 
                    FROM stock_transfer 
                    WHERE (source_item LIKE ? OR dest_item LIKE ?) AND date = ?
                    ORDER BY date DESC, source_item
                """
                params = (f'%{search_term}%', f'%{search_term}%', entries['date_filter'].get_date().isoformat())
            
            cursor.execute(query, params)
            rows = cursor.fetchall()
            
            # Clear tree and display results
            for row in tree.get_children():
                tree.delete(row)
            
            if rows:
                for source_item, qty, dest_item, transfer_date in rows:
                    # Format date for display
                    try:
                        formatted_date = datetime.strptime(transfer_date, '%Y-%m-%d').strftime('%d/%m/%Y')
                    except:
                        formatted_date = transfer_date
                    
                    # Format quantity display
                    if qty > 0:
                        qty_display = f"+{qty} pcs"
                    else:
                        qty_display = f"{qty} pcs"
                    
                    tree.insert('', 'end', values=(source_item.title(), qty_display, dest_item.title(), formatted_date))
            else:
                tree.insert('', 'end', values=("No records found", "", "", ""))
        else:
            # If search is empty, show all transfers
            search_transfers(tree, entries['date_filter'], entries['show_date_var'])
    
    def clear_item_search():
        entries['item_search'].delete(0, tk.END)
        search_transfers(tree, entries['date_filter'], entries['show_date_var'])
    
    tk.Button(filter_frame, text="Search", command=search_by_item, bg='lightblue').grid(row=1, column=2, padx=5, pady=5)
    tk.Button(filter_frame, text="Clear Search", command=clear_item_search, bg='lightyellow').grid(row=1, column=3, padx=5, pady=5)
    
    # Row 2: Search info and refresh button
    info_label = tk.Label(filter_frame, text="Tip: Double-click any item to see all transfers involving that item", 
                          font=("Arial", 9), fg="blue")
    info_label.grid(row=2, column=0, columnspan=4, pady=5)
    
    # Transfer History Frame
    history_frame = tk.LabelFrame(main_frame, text="Transfer History", padx=5, pady=5)
    history_frame.pack(fill='both', expand=True)
    
    # Create Treeview with scrollbars
    tree_frame = tk.Frame(history_frame)
    tree_frame.pack(fill='both', expand=True)
    
    # Add vertical scrollbar
    v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
    v_scrollbar.pack(side=tk.RIGHT, fill='y')
    
    # Add horizontal scrollbar
    h_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal")
    h_scrollbar.pack(side=tk.BOTTOM, fill='x')
    
    tree = ttk.Treeview(tree_frame, columns=("Source", "Quantity", "Destination", "Date"), 
                        show="headings", selectmode="browse",
                        yscrollcommand=v_scrollbar.set,
                        xscrollcommand=h_scrollbar.set)
    
    v_scrollbar.config(command=tree.yview)
    h_scrollbar.config(command=tree.xview)
    
    # Configure columns
    tree.heading("Source", text="Source Item")
    tree.heading("Quantity", text="Quantity")
    tree.heading("Destination", text="Destination Item")
    tree.heading("Date", text="Transfer Date")
    
    tree.column("Source", width=200, minwidth=150)
    tree.column("Quantity", width=100, minwidth=80, anchor='center')
    tree.column("Destination", width=200, minwidth=150)
    tree.column("Date", width=120, minwidth=100, anchor='center')
    
    tree.pack(fill='both', expand=True)
    
    # Status label
    status_label = tk.Label(main_frame, text="", font=("Arial", 9), fg="green")
    status_label.pack(pady=5)
    
    # Double-click event to show all transfers for selected item
    def on_item_double_click(event):
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Please select a transfer record to view item details.")
            return
        
        # Get the selected item
        selected_item = tree.item(selected[0])
        source_item = selected_item['values'][0]
        dest_item = selected_item['values'][2]
        
        # Create a new window to show all transfers for the clicked item
        detail_window = tk.Toplevel(tab_frame)
        detail_window.title(f"Transfer History for Items")
        detail_window.geometry("800x500")
        detail_window.transient(tab_frame)
        detail_window.grab_set()
        
        # Create notebook for tabs (source and destination)
        detail_notebook = ttk.Notebook(detail_window)
        detail_notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Function to show transfers for a specific item
        def show_item_transfers(item_name, role):
            # Create frame for this item
            item_frame = ttk.Frame(detail_notebook)
            detail_notebook.add(item_frame, text=f"{role}: {item_name.title()}")
            
            # Search query for this item
            query = """
                SELECT source_item, quantity, dest_item, date 
                FROM stock_transfer 
                WHERE source_item = ? OR dest_item = ?
                ORDER BY date DESC, source_item
            """
            cursor.execute(query, (item_name, item_name))
            rows = cursor.fetchall()
            
            if not rows:
                tk.Label(item_frame, text=f"No transfer records found for {item_name.title()}", 
                        font=("Arial", 12), fg="red").pack(pady=50)
                return
            
            # Create tree for this item
            item_tree_frame = tk.Frame(item_frame)
            item_tree_frame.pack(fill='both', expand=True, padx=10, pady=10)
            
            # Scrollbars
            v_scroll = ttk.Scrollbar(item_tree_frame, orient="vertical")
            v_scroll.pack(side=tk.RIGHT, fill='y')
            h_scroll = ttk.Scrollbar(item_tree_frame, orient="horizontal")
            h_scroll.pack(side=tk.BOTTOM, fill='x')
            
            item_tree = ttk.Treeview(item_tree_frame, columns=("Source", "Quantity", "Destination", "Date"), 
                                     show="headings", selectmode="browse",
                                     yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
            v_scroll.config(command=item_tree.yview)
            h_scroll.config(command=item_tree.xview)
            
            item_tree.heading("Source", text="Source Item")
            item_tree.heading("Quantity", text="Quantity")
            item_tree.heading("Destination", text="Destination Item")
            item_tree.heading("Date", text="Transfer Date")
            
            item_tree.column("Source", width=200)
            item_tree.column("Quantity", width=100, anchor='center')
            item_tree.column("Destination", width=200)
            item_tree.column("Date", width=120, anchor='center')
            
            item_tree.pack(fill='both', expand=True)
            
            # Calculate statistics
            total_in = 0
            total_out = 0
            
            for src, qty, dest, trans_date in rows:
                try:
                    formatted_date = datetime.strptime(trans_date, '%Y-%m-%d').strftime('%d/%m/%Y')
                except:
                    formatted_date = trans_date
                
                # Check if this item is source or destination
                if src == item_name:
                    # Item is source (sending out)
                    qty_display = f"-{abs(qty)} pcs"
                    total_out += abs(qty)
                    item_tree.insert('', 'end', values=(src.title(), qty_display, dest.title(), formatted_date), 
                                    tags=('out',))
                else:
                    # Item is destination (receiving)
                    qty_display = f"+{abs(qty)} pcs"
                    total_in += abs(qty)
                    item_tree.insert('', 'end', values=(src.title(), qty_display, dest.title(), formatted_date),
                                    tags=('in',))
            
            # Color coding
            item_tree.tag_configure('out', foreground='red')
            item_tree.tag_configure('in', foreground='green')
            
            # Add summary frame
            summary_frame = tk.Frame(item_frame)
            summary_frame.pack(fill='x', padx=10, pady=10)
            
            net_change = total_in - total_out
            summary_text = f"Summary for {item_name.title()}:\n"
            summary_text += f"• Total Received: {total_in} pcs\n"
            summary_text += f"• Total Sent: {total_out} pcs\n"
            summary_text += f"• Net Change: {net_change:+d} pcs"
            
            summary_label = tk.Label(summary_frame, text=summary_text, font=("Arial", 10, "bold"),
                                     justify=tk.LEFT, fg="blue" if net_change >= 0 else "red")
            summary_label.pack(anchor='w')
            
            # Export button for this item
            def export_item_transfers():
                file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                         filetypes=[("Excel files", "*.xlsx")],
                                                         initialfile=f"transfers_{item_name}_{date.today().isoformat()}.xlsx")
                if not file_path:
                    return
                
                try:
                    wb = Workbook()
                    ws = wb.active
                    ws.title = f"Transfers - {item_name.title()}"
                    
                    bold_font = Font(bold=True)
                    ws.cell(row=1, column=1, value="Source Item").font = bold_font
                    ws.cell(row=1, column=2, value="Quantity").font = bold_font
                    ws.cell(row=1, column=3, value="Destination Item").font = bold_font
                    ws.cell(row=1, column=4, value="Date").font = bold_font
                    
                    row_num = 2
                    for src, qty, dest, trans_date in rows:
                        try:
                            formatted_date = datetime.strptime(trans_date, '%Y-%m-%d').strftime('%d/%m/%Y')
                        except:
                            formatted_date = trans_date
                        ws.cell(row=row_num, column=1, value=src.title())
                        ws.cell(row=row_num, column=2, value=qty)
                        ws.cell(row=row_num, column=3, value=dest.title())
                        ws.cell(row=row_num, column=4, value=formatted_date)
                        row_num += 1
                    
                    # Add summary
                    ws.cell(row=row_num+1, column=1, value="SUMMARY").font = bold_font
                    ws.cell(row=row_num+2, column=1, value="Total Received:")
                    ws.cell(row=row_num+2, column=2, value=total_in)
                    ws.cell(row=row_num+3, column=1, value="Total Sent:")
                    ws.cell(row=row_num+3, column=2, value=total_out)
                    ws.cell(row=row_num+4, column=1, value="Net Change:")
                    ws.cell(row=row_num+4, column=2, value=net_change)
                    
                    wb.save(file_path)
                    messagebox.showinfo("Export Successful", f"Transfer history exported to:\n{file_path}")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to export: {e}")
            
            export_btn = tk.Button(item_frame, text=f"Export {item_name.title()} Transfers to Excel", 
                                  command=export_item_transfers, bg='lightgreen')
            export_btn.pack(pady=10)
        
        # Show transfers for source and destination items
        if source_item != "No records found":
            show_item_transfers(source_item.lower(), "Item")
        
        # Also show for destination if different
        if dest_item != "No records found" and dest_item.lower() != source_item.lower():
            show_item_transfers(dest_item.lower(), "Item")
    
    tree.bind("<Double-1>", on_item_double_click)
    
    # Right-click context menu
    context_menu = tk.Menu(tab_frame, tearoff=0)
    context_menu.add_command(label="Show All Transfers for This Item", 
                            command=lambda: on_item_double_click(None))
    
    def show_context_menu(event):
        selected_item = tree.identify_row(event.y)
        if selected_item:
            tree.selection_set(selected_item)
            context_menu.tk_popup(event.x_root, event.y_root)
    
    tree.bind("<Button-3>", show_context_menu)
    
    # Refresh button
    def refresh_transfers():
        search_transfers(tree, entries['date_filter'], entries['show_date_var'])
        entries['item_search'].delete(0, tk.END)
        status_label.config(text="Transfer history refreshed", fg="green")
        tab_frame.after(3000, lambda: status_label.config(text=""))
    
    refresh_btn = tk.Button(main_frame, text="Refresh Transfer History", command=refresh_transfers, 
                           bg='lightblue', font=("Arial", 10, "bold"), padx=20, pady=5)
    refresh_btn.pack(pady=5)
    
    # Export all transfers button
    def export_all_transfers():
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")],
                                                 initialfile=f"all_transfers_{date.today().isoformat()}.xlsx")
        if not file_path:
            return
        
        try:
            # Get all transfers based on current filters
            if entries['show_date_var'].get():
                query = "SELECT source_item, quantity, dest_item, date FROM stock_transfer WHERE date = ? ORDER BY date DESC, source_item"
                cursor.execute(query, (entries['date_filter'].get_date().isoformat(),))
            else:
                query = "SELECT source_item, quantity, dest_item, date FROM stock_transfer ORDER BY date DESC, source_item"
                cursor.execute(query)
            
            rows = cursor.fetchall()
            
            if not rows:
                messagebox.showinfo("No Data", "No transfer records to export.")
                return
            
            wb = Workbook()
            ws = wb.active
            ws.title = "All Transfers"
            
            bold_font = Font(bold=True)
            ws.cell(row=1, column=1, value="Source Item").font = bold_font
            ws.cell(row=1, column=2, value="Quantity").font = bold_font
            ws.cell(row=1, column=3, value="Destination Item").font = bold_font
            ws.cell(row=1, column=4, value="Date").font = bold_font
            
            # Set column widths
            ws.column_dimensions['A'].width = 20
            ws.column_dimensions['B'].width = 15
            ws.column_dimensions['C'].width = 20
            ws.column_dimensions['D'].width = 15
            
            row_num = 2
            for src, qty, dest, trans_date in rows:
                try:
                    formatted_date = datetime.strptime(trans_date, '%Y-%m-%d').strftime('%d/%m/%Y')
                except:
                    formatted_date = trans_date
                ws.cell(row=row_num, column=1, value=src.title())
                ws.cell(row=row_num, column=2, value=f"{qty} pcs")
                ws.cell(row=row_num, column=3, value=dest.title())
                ws.cell(row=row_num, column=4, value=formatted_date)
                row_num += 1
            
            wb.save(file_path)
            messagebox.showinfo("Export Successful", f"All transfers exported to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export: {e}")
    
    export_all_btn = tk.Button(main_frame, text="Export All Transfers to Excel", 
                               command=export_all_transfers, bg='lightgreen', padx=20, pady=5)
    export_all_btn.pack(pady=5)
    
    # Statistics label
    stats_label = tk.Label(main_frame, text="", font=("Arial", 9), fg="purple")
    stats_label.pack(pady=5)
    
    def update_statistics():
        cursor.execute("SELECT COUNT(*) FROM stock_transfer")
        total_transfers = cursor.fetchone()[0]
        
        cursor.execute("SELECT COUNT(DISTINCT source_item) FROM stock_transfer")
        unique_sources = cursor.fetchone()[0]
        
        cursor.execute("SELECT COUNT(DISTINCT dest_item) FROM stock_transfer")
        unique_dests = cursor.fetchone()[0]
        
        stats_label.config(text=f"Total Transfers: {total_transfers} | Unique Source Items: {unique_sources} | Unique Destination Items: {unique_dests}")
    
    # Override search_transfers to update statistics
    original_search_transfers = search_transfers
    
    def enhanced_search_transfers(tree, date_picker, show_date_var):
        original_search_transfers(tree, date_picker, show_date_var)
        update_statistics()
    
    # Replace the global search_transfers function
    import sys
    sys.modules[__name__].search_transfers = enhanced_search_transfers
    
    # Initial load
    update_transfer_tree(tree, entries['date_filter'], entries['show_date_var'])
    update_statistics()
    
    # Bind date filter changes
    def on_date_filter_change():
        search_transfers(tree, entries['date_filter'], entries['show_date_var'])
        update_statistics()
    
    entries['show_date_var'].trace('w', lambda *args: on_date_filter_change())
    
    return tree, entries

# Main application window
root = tk.Tk()
root.title("Inventory Manager")
root.geometry("1200x700")
notebook = ttk.Notebook(root)
notebook.pack(pady=10, expand=True, fill='both')

date_entries_dict = {}

inward_frame = ttk.Frame(notebook)
notebook.add(inward_frame, text="Inward")
inward_tree, inward_entries = create_tab(inward_frame, 'inward_inventory', 'Memo No', 'Merchant', None, None, None, None, None, None, date_entries_dict, "Inward")

outward_frame = ttk.Frame(notebook)
notebook.add(outward_frame, text="Outward")
outward_tree, outward_entries = create_tab(outward_frame, 'outward_inventory', 'Bale No', 'Transport', inward_tree, None, None, inward_entries, None, None, date_entries_dict, "Outward")

summary_frame = ttk.Frame(notebook)
notebook.add(summary_frame, text="Summary")
summary_tree, summary_entries = create_summary_tab(summary_frame, date_entries_dict, "Summary")

transfer_frame = ttk.Frame(notebook)
notebook.add(transfer_frame, text="Transfer")
transfer_tree, transfer_entries = create_transfer_tab(transfer_frame, inward_tree, outward_tree, summary_tree, 
                                                       inward_entries, outward_entries, summary_entries, 
                                                       date_entries_dict, "Transfer")

orders_frame = ttk.Frame(notebook)
notebook.add(orders_frame, text="Orders")
orders_tree, orders_entries = create_orders_tab(orders_frame, date_entries_dict, "Orders")

# Add global undo/redo keyboard shortcuts
def on_undo(event=None):
    current_tab = notebook.select()
    tab_name = notebook.tab(current_tab, "text")
    
    # Get the appropriate tree and entries based on current tab
    tree_widget = None
    entry_widgets = None
    
    if tab_name == "Inward":
        tree_widget = inward_tree
        entry_widgets = inward_entries
    elif tab_name == "Outward":
        tree_widget = outward_tree
        entry_widgets = outward_entries
    elif tab_name == "Summary":
        tree_widget = summary_tree
        entry_widgets = summary_entries
    elif tab_name == "Transfer":
        tree_widget = transfer_tree
        entry_widgets = transfer_entries
    elif tab_name == "Orders":
        tree_widget = orders_tree
        entry_widgets = orders_entries
    
    if undo_last_action(tab_name, tree_widget, entry_widgets):
        # Refresh the current tab's view
        if tab_name == "Inward" and inward_entries:
            update_tree('inward_inventory', inward_tree, inward_entries['search'], inward_entries['date_filter'], 
                       inward_entries['entity_filter'], inward_entries['ref_filter'], inward_entries['show_date_var'])
        elif tab_name == "Outward" and outward_entries:
            update_tree('outward_inventory', outward_tree, outward_entries['search'], outward_entries['date_filter'], 
                       outward_entries['entity_filter'], outward_entries['ref_filter'], outward_entries['show_date_var'])
        elif tab_name == "Summary" and summary_entries:
            update_summary_tree(summary_tree, summary_entries['search'], summary_entries['date_filter'], summary_entries['show_date_var'])
        elif tab_name == "Transfer" and transfer_entries:
            update_transfer_tree(transfer_tree, transfer_entries['date_filter'], transfer_entries['show_date_var'])
        elif tab_name == "Orders" and orders_entries:
            update_orders_tree(orders_tree, orders_entries['search'], orders_entries['date_filter'], orders_entries['show_date_var'])

def on_redo(event=None):
    current_tab = notebook.select()
    tab_name = notebook.tab(current_tab, "text")
    
    tree_widget = None
    entry_widgets = None
    
    if tab_name == "Inward":
        tree_widget = inward_tree
        entry_widgets = inward_entries
    elif tab_name == "Outward":
        tree_widget = outward_tree
        entry_widgets = outward_entries
    elif tab_name == "Summary":
        tree_widget = summary_tree
        entry_widgets = summary_entries
    elif tab_name == "Transfer":
        tree_widget = transfer_tree
        entry_widgets = transfer_entries
    elif tab_name == "Orders":
        tree_widget = orders_tree
        entry_widgets = orders_entries
    
    if redo_last_action(tab_name, tree_widget, entry_widgets):
        # Refresh the current tab's view
        if tab_name == "Inward" and inward_entries:
            update_tree('inward_inventory', inward_tree, inward_entries['search'], inward_entries['date_filter'], 
                       inward_entries['entity_filter'], inward_entries['ref_filter'], inward_entries['show_date_var'])
        elif tab_name == "Outward" and outward_entries:
            update_tree('outward_inventory', outward_tree, outward_entries['search'], outward_entries['date_filter'], 
                       outward_entries['entity_filter'], outward_entries['ref_filter'], outward_entries['show_date_var'])
        elif tab_name == "Summary" and summary_entries:
            update_summary_tree(summary_tree, summary_entries['search'], summary_entries['date_filter'], summary_entries['show_date_var'])
        elif tab_name == "Transfer" and transfer_entries:
            update_transfer_tree(transfer_tree, transfer_entries['date_filter'], transfer_entries['show_date_var'])
        elif tab_name == "Orders" and orders_entries:
            update_orders_tree(orders_tree, orders_entries['search'], orders_entries['date_filter'], orders_entries['show_date_var'])

# Bind Ctrl+Z for undo and Ctrl+Y or Ctrl+Shift+Z for redo
root.bind('<Control-z>', on_undo)
root.bind('<Control-Z>', on_undo)  # For shift+ctrl+z
root.bind('<Control-y>', on_redo)
root.bind('<Control-Y>', on_redo)

# Also add a status label to show undo/redo availability
status_bar = tk.Label(root, text="Press Ctrl+Z to undo last action | Ctrl+Y to redo", 
                      bd=1, relief=tk.SUNKEN, anchor=tk.W)
status_bar.pack(side=tk.BOTTOM, fill=tk.X)

def focus_date_entry(event):
    current_tab = notebook.select()
    tab_name = notebook.tab(current_tab, "text")
    if tab_name in date_entries_dict:
        date_entries_dict[tab_name].focus_set()

root.bind("<F2>", focus_date_entry)

root.mainloop()

conn.close()
