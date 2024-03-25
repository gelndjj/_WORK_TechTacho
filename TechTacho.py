import tkinter as tk, sqlite3, os, pandas as pd, ctypes, matplotlib.pyplot as plt, webbrowser, urllib.parse
from tkinter import ttk, messagebox
from tkcalendar import Calendar
from datetime import datetime, timedelta
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from faker import Faker

try:
    # Try to set DPI awareness to make text and elements clear
    ctypes.windll.shcore.SetProcessDpiAwareness(1)  # 1: System DPI aware, 2: Per monitor DPI aware
except AttributeError:
    # Fallback if SetProcessDpiAwareness does not exist (Windows versions < 8.1)
    ctypes.windll.user32.SetProcessDPIAware()


class EquipmentTrackingTab(tk.Frame):
    def __init__(self, parent, bg_color, app):
        super().__init__(parent, background=bg_color)

        self.app = app

        # Initialize the CSV filename as None; it will be set when a database is selected
        self.emails_file = None

        # Initialize the entry frame first
        self.entry_frame = tk.Frame(self, background=bg_color)
        self.entry_frame.grid(row=0, column=0, sticky="nsew")

        # Now that entry_frame is initialized, you can create db_label
        self.db_label = tk.Label(self.entry_frame, text="Select Database:", background=bg_color)
        self.db_label.pack(pady=5)
        self.db_combo = ttk.Combobox(self.entry_frame, postcommand=self.update_db_list)
        self.db_combo.pack(pady=5)
        self.db_combo.bind("<<ComboboxSelected>>", self.combined_database_selection_handler)

        self.email_label = tk.Label(self.entry_frame, text="User Email:", background=bg_color)
        self.email_label.pack(pady=5)
        self.email_combobox = ttk.Combobox(self.entry_frame, width=30)
        self.email_combobox.pack(pady=5)
        self.email_combobox.bind("<<ComboboxSelected>>", self.filter_tree_view_by_email)
        self.email_combobox.bind('<Return>', self.on_combobox_enter)

        self.equipment_label = tk.Label(self.entry_frame, text="Equipment:", background=bg_color)
        self.equipment_label.pack(pady=5)
        self.equipment_combobox = ttk.Combobox(self.entry_frame)
        self.equipment_combobox.pack(pady=5)

        self.due_date_label = tk.Label(self.entry_frame, text="Due Date:", background=bg_color)
        self.due_date_label.pack(pady=5)
        self.calendar = Calendar(self.entry_frame, selectmode='day', year=datetime.now().year,
                                 month=datetime.now().month, day=datetime.now().day)
        self.calendar.pack(pady=5)

        self.submit_button = ttk.Button(self.entry_frame, text="Submit", command=self.add_entry)
        self.submit_button.pack(pady=10)

        # Define the columns for the Treeview including the 'Status' column
        self.tree_frame = tk.Frame(self, background=bg_color)
        self.tree_frame.grid(row=0, column=1, sticky="nsew")
        self.tree_view = ttk.Treeview(self.tree_frame,
                                      columns=("ID", "Date", "Email", "Equipment", "Due Date", "Status"),
                                      show="headings")

        # Define the column headings including the 'Status' column
        self.tree_view.heading("ID", text="ID", anchor="center")
        self.sort_reverse = False
        self.tree_view.heading("Date", text="Date", anchor="center", command=lambda: self.sort_by_date(self.sort_reverse))
        self.tree_view.heading("Email", text="User Email", anchor="center", command=lambda: self.sort_by_column("Email", False))
        self.tree_view.heading("Equipment", text="Equipment", anchor="center", command=lambda: self.sort_by_column("Email", False))
        self.tree_view.heading("Due Date", text="Due Date", anchor="center", command=lambda: self.sort_by_date(self.sort_reverse))
        self.tree_view.heading("Status", text="Status", anchor="center", command=lambda: self.sort_by_column("Status", False))

        # Configure the columns, including the 'Status' column
        self.tree_view.column("ID", anchor='center', width=50)
        self.tree_view.column("Date", anchor='center', width=100)
        self.tree_view.column("Email", anchor='center', width=300)
        self.tree_view.column("Equipment", anchor='center', width=220)
        self.tree_view.column("Due Date", anchor='center', width=100)
        self.tree_view.column("Status", anchor='center', width=100)

        self.tree_view.pack(expand=True, fill='both')

        self.tree_view.bind("<Double-1>", self.handle_double_click)
        self.tree_view.bind("<Button-3>", self.show_context_menu)
        self.tree_view.bind("<<TreeviewSelect>>", self.on_tree_select)

        self.setup_tags()

        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(label="Edit Date", command=self.set_custom_date)
        self.context_menu.add_command(label="Edit Due Date", command=self.edit_due_date)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Edit Return Date", command=self.edit_return_date)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Not Returned", command=self.mark_as_not_returned)
        self.context_menu.add_command(label="Returned", command=self.mark_as_returned)

        self.bind_all("<Control-c>", lambda e: self.create_new_db())
        self.tree_view.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.tree_view.bind('<s>', self.on_item_double_click)
        self.equipment_combobox.bind("<<ComboboxSelected>>", self.filter_tree_view_by_equipment)
        self.email_combobox.bind("<<ComboboxSelected>>", self.filter_tree_view_by_email)
        self.equipment_combobox.bind('<Return>', self.on_combobox_enter)
        self.email_combobox.bind('<KeyRelease>', self.on_email_keyrelease)

        # Bind Ctrl+Shift+D to generate_fake_database
        self.bind_all("<Control-Shift-D>", lambda e: self.generate_fake_data())

        # Load emails into the combobox
        self.load_emails_into_combobox()

        # Load initial data
        self.on_database_selected()

    def generate_fake_data(self):
        fake = Faker()
        number_of_emails = 1000

        db_index = 0
        while os.path.exists(f'fake_db_{db_index}.db'):
            db_index += 1

        db_filename = f'fake_db_{db_index}.db'
        csv_filename = f'fake_db_{db_index}_users.csv'

        fake_emails = [fake.email() for _ in range(number_of_emails)]
        emails_df = pd.DataFrame(fake_emails, columns=['User Email'])
        emails_df.to_csv(csv_filename, index=False)

        equipment_names = ['Laptop', 'Projector', 'Camera', 'Microphone', 'Speaker', 'Mouse', 'Keyboard', 'Screen', 'Smartphone']
        dates, emails, equipments, due_dates, statuses = [], [], [], [], []

        for _ in range(number_of_emails):
            email = fake.random_element(elements=fake_emails)
            equipment = fake.random_element(elements=equipment_names)
            borrow_date = fake.date_between(start_date="-2y", end_date="today")
            due_date, status = self.calculate_due_date_and_return_status(borrow_date, fake)

            dates.append(borrow_date.strftime('%Y-%m-%d'))
            emails.append(email)
            equipments.append(equipment)
            due_dates.append(due_date.strftime('%Y-%m-%d'))
            statuses.append(status)

        fake_db_df = pd.DataFrame({
            'Date': dates,
            'Email': emails,
            'Equipment': equipments,
            'DueDate': due_dates,
            'Status': statuses
        })

        conn = sqlite3.connect(db_filename)
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS equipment (
                            ID INTEGER PRIMARY KEY AUTOINCREMENT,
                            Date TEXT,
                            Email TEXT,
                            Equipment TEXT,
                            DueDate TEXT,
                            Status TEXT)''')
        conn.commit()

        fake_db_df.to_sql('equipment', conn, if_exists='append', index=False)
        conn.close()

        messagebox.showinfo("Info", f"Generated a fake database with {number_of_emails} entries.")
        self.app.refresh_pie_charts()
        self.app.load_last_selected_db()

    def calculate_due_date_and_return_status(self, borrow_date, fake):
        due_date = borrow_date + timedelta(days=fake.random_int(min=1, max=60))
        days_overdue = fake.random_int(min=1, max=10)  # Random number of overdue days

        # Randomly decide if the item is returned
        if fake.boolean(chance_of_getting_true=75):  # 75% chance the item is returned
            # If returned, decide if it's late
            if fake.boolean(chance_of_getting_true=20):  # 20% chance the item is returned late
                return due_date, f'Returned +{days_overdue}'
            else:
                return due_date, 'Returned'
        else:
            # If not returned, decide if it's overdue
            if due_date < datetime.now().date():
                return due_date, f'+{days_overdue}'
            else:
                return due_date, 'Not Returned'

    def on_item_double_click(self, event):
        # Get the selected item
        item = self.tree_view.selection()[0]
        item_values = self.tree_view.item(item, 'values')

        # Extract email and equipment name assuming they are in known columns
        email = item_values[2]  # Change index based on your setup
        equipment_name = item_values[3]  # Change index based on your setup

        # Define the email subject and body
        subject = f"[IT DESK BRUGES] Kind Reminder: Please Return the Equipment ({equipment_name})"
        body = (f"Hello,\n\nThis is a kind reminder to please return the {equipment_name} you borrowed. "
                "This email is generated automatically. If you have already returned the item, please disregard this message.\n\nThank you.")

        # URL-encode the subject and body
        encoded_subject = urllib.parse.quote(subject)
        encoded_body = urllib.parse.quote(body)

        # Construct the mailto URL
        mailto_url = f"mailto:{email}?subject={encoded_subject}&body={encoded_body}"

        # Open the default mail client
        webbrowser.open(mailto_url)

    def mark_as_returned(self):
        selected_items = self.tree_view.selection()
        for item in selected_items:
            item_values = self.tree_view.item(item, 'values')
            current_status = item_values[-1]  # Assuming Status is the last column

            # Determine the new status based on the existing status
            if current_status.startswith('+'):
                new_status = "Returned " + current_status
            else:
                new_status = "Returned"

            # Update the TreeView item with the new status
            new_values = item_values[:-1] + (new_status,)
            self.tree_view.item(item, values=new_values)

            # Update the status in the database
            self.update_status_in_db(item, new_status)

            # Update the Treeview item color immediately
            if new_status.startswith("Returned"):
                self.tree_view.item(item, tags=('returned',))
            else:
                self.tree_view.item(item, tags=('default',))

            # Refresh the treeview and the colors
            self.tree_view.update_idletasks()
            self.update_row_colors()

        # Optionally, refresh the entire TreeView to reflect changes
        self.load_selected_db(None)
        self.app.refresh_pie_charts()

    def mark_as_not_returned(self):
        selected_items = self.tree_view.selection()
        for item in selected_items:
            item_values = self.tree_view.item(item, 'values')
            new_values = item_values[:-1] + ('Not Returned',)
            self.tree_view.item(item, values=new_values)
            self.update_item_color_and_status(item, 'Not Returned')
        self.app.refresh_pie_charts()

    def setup_tags(self):
        # Using a lighter shade of green
        self.light_green = '#90ee90'  # This is a hexadecimal color code for a light green, similar to 'lightgreen'
        self.tree_view.tag_configure('returned', background=self.light_green)
        self.tree_view.tag_configure('default', background='')
        self.tree_view.tag_configure('overdue', background='red')

    def show_context_menu(self, event):
        # Show the menu only if there are selected items
        if self.tree_view.selection():
            self.context_menu.post(event.x_root, event.y_root)

    def scan_for_databases(self):
        # Scan the current directory (or a specific path) for .db files
        db_files = []
        for file in os.listdir('.'):  # Adjust the path as necessary
            if file.endswith('.db'):
                db_files.append(file)
        return db_files

    def load_equipment_entries(self):
        db_file = self.db_combo.get()
        if db_file:
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()

            try:
                cursor.execute("SELECT DISTINCT Equipment FROM equipment")
                equipment_list = [row[0] for row in cursor.fetchall()]
                self.equipment_combobox['values'] = equipment_list
            except sqlite3.Error as e:
                print("Database error:", e)
            finally:
                conn.close()

    def load_email_entries(self):
        db_file = self.db_combo.get()

        if db_file:
            # Extract the base name from the database file and construct the CSV filename
            base_name = db_file.split('.db')[0]
            csv_filename = f"{base_name}_users.csv"

            try:
                if os.path.exists(csv_filename):
                    df = pd.read_csv(csv_filename)
                    emails = df['User Email'].dropna().unique().tolist()
                    self.email_combobox['values'] = emails
                else:
                    print("CSV file not found. Ensure the file exists in the specified path.")
                    self.email_combobox['values'] = []  # Clear the combobox if the file doesn't exist
            except Exception as e:
                print(f"An error occurred while loading the CSV: {e}")
                self.email_combobox['values'] = []  # Clear the combobox in case of an error
        else:
            print("No database selected.")
            self.email_combobox['values'] = []  # Clear the combobox if no database is selected

    def add_entry(self):
        current_date = datetime.now().strftime("%Y-%m-%d")
        email = self.email_combobox.get()
        equipment = self.equipment_combobox.get()  # Get value from the combobox
        due_date = self.calendar.get_date()
        status = 'Not Returned'  # Default status for new entries

        try:
            formatted_due_date = datetime.strptime(due_date, "%m/%d/%y").strftime("%Y-%m-%d")
        except ValueError:
            print("Error: Incorrect date format from calendar.")
            return

        db_file = self.db_combo.get()
        if db_file:
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()

            try:
                cursor.execute(
                    "INSERT INTO equipment (Date, Email, Equipment, DueDate, Status) VALUES (?, ?, ?, ?, ?)",
                    (current_date, email, equipment, formatted_due_date, status))

                conn.commit()

                # Fetch the last row id
                last_id = cursor.lastrowid  # This should be the actual ID of the inserted row

                # Insert the new entry into the Treeview with the correct ID and include the status
                self.tree_view.insert("", "end",
                                      values=(last_id, current_date, email, equipment, formatted_due_date, status))

                # Assuming you want to add the new equipment to the combobox values if it's not already there
                existing_values = self.equipment_combobox['values']
                if equipment not in existing_values:
                    self.equipment_combobox['values'] = existing_values + (equipment,)

            except sqlite3.Error as e:
                print("Database error:", e)
            except Exception as e:
                print("Exception in _query:", e)
            finally:
                conn.close()

            # Clear the equipment combobox selection
            self.email_combobox.set('')
            self.load_email_entries()
            self.equipment_combobox.set('')
            self.load_equipment_entries()
            self.sort_by_date(reverse=True)
            # After adding a new record to the database, update the CSV file
            self.update_emails_file(email)

            self.app.refresh_pie_charts()
            self.app.load_last_selected_db()

    def update_emails_file(self, new_email):
        # Load existing data
        df = pd.read_csv(self.emails_file)

        # Append new email if it doesn't exist
        if new_email not in df['User Email'].values:
            df.loc[len(df)] = [new_email]
            df.to_csv(self.emails_file, index=False)

        # Reload the combobox values
        self.load_emails_into_combobox()

    def load_emails_into_combobox(self):
        db_file = self.db_combo.get()

        if db_file:
            # Extract the base name from the database file and construct the CSV filename
            base_name = os.path.splitext(os.path.basename(db_file))[0]
            csv_filename = f"{base_name}_users.csv"

            try:
                if os.path.exists(csv_filename):
                    df = pd.read_csv(csv_filename)
                    self.all_emails = df['User Email'].dropna().unique().tolist()
                    self.email_combobox['values'] = self.all_emails
                else:
                    print("CSV file not found. Ensure the file exists in the specified path.")
                    self.all_emails = []  # Clear the list if the file doesn't exist
            except Exception as e:
                print(f"An error occurred while loading the CSV: {e}")
                self.all_emails = []  # Clear the list in case of an error
        else:
            print("No database selected.")
            self.all_emails = []  # Clear the list if no database is selected
            self.email_combobox['values'] = []

    def on_email_keyrelease(self, event):
        # Get the current entry value
        value = self.email_combobox.get()

        if value == '':
            self.email_combobox['values'] = self.all_emails
            self.email_combobox.set('')
        else:
            # Filter the email list
            filtered_data = [email for email in self.all_emails if email.lower().startswith(value.lower())]

            # Update the combobox dropdown values
            self.email_combobox['values'] = filtered_data

            # Set the entry to the typed value & restore the cursor position
            self.email_combobox.set(value)
            self.email_combobox.icursor(len(value))

        # If the dropdown is not displayed, display it
        if not self.email_combobox.master.winfo_ismapped():
            self.email_combobox.event_generate('<Down>')

    def edit_due_date(self):
        selected_items = self.tree_view.selection()
        if not selected_items:
            messagebox.showinfo("Info", "No item selected.")
            return

        def on_date_selected():
            selected_date = cal.selection_get()
            for item in selected_items:
                self.update_due_date(item, selected_date.strftime('%Y-%m-%d'))
            date_window.destroy()

        # Create a new top-level window
        date_window = tk.Toplevel(self)
        date_window.title("Select Date")

        # Create a Calendar widget in the new window
        cal = Calendar(date_window, selectmode='day', date_pattern='yyyy-mm-dd')
        cal.pack(pady=10)

        # Confirmation button to use the selected date
        tk.Button(date_window, text="Ok", command=on_date_selected).pack(pady=10)

    def edit_return_date(self):
        selected_items = self.tree_view.selection()
        if not selected_items:
            messagebox.showinfo("Info", "No item selected.")
            return

        def on_date_selected():
            pseudo_current_date = cal.selection_get()
            for item in selected_items:
                self.update_status_based_on_date(item, pseudo_current_date)
            status_window.destroy()

        # Create a new window for date selection
        status_window = tk.Toplevel(self)
        status_window.title("Select Date")

        # Calendar widget
        cal = Calendar(status_window, selectmode='day', date_pattern='yyyy-mm-dd')
        cal.pack(pady=10)

        # Confirmation button to use the selected date
        tk.Button(status_window, text="Ok", command=on_date_selected).pack(pady=10)

    def on_tree_select(self, event):
        selected_item = self.tree_view.selection()

        if selected_item:
            item = selected_item[0]
            row = self.tree_view.item(item, 'values')

            if len(row) >= 6:  # Ensure there are enough values, including the new 'Status' column
                # Extracting the Email and Equipment based on their positions in the row
                email = row[2]  # Assuming Email is the third column
                equipment = row[3]  # Assuming Equipment is the fourth column

                # Update the Email Entry and Equipment Combobox with the selected values
                self.email_combobox.set(email)
                self.equipment_combobox.set(equipment)

    def filter_tree_view_by_equipment(self, event=None):
        selected_equipment = self.equipment_combobox.get()

        # Clear the current TreeView
        for item in self.tree_view.get_children():
            self.tree_view.delete(item)

        # Connect to the database
        db_file = self.db_combo.get()
        if db_file:
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()

            try:
                if selected_equipment:
                    # Load only the entries that match the selected equipment, including the Status
                    cursor.execute(
                        "SELECT ID, Date, Email, Equipment, DueDate, Status FROM equipment WHERE Equipment = ?",
                        (selected_equipment,))
                else:
                    # If no equipment is selected, load all entries, including the Status
                    cursor.execute("SELECT ID, Date, Email, Equipment, DueDate, Status FROM equipment")

                rows = cursor.fetchall()

                # Populate the TreeView with the fetched rows
                for row in rows:
                    # Ensure that the order of the values corresponds to the TreeView columns
                    self.tree_view.insert("", "end", values=row)

                # Update the row colors based on the Status and Due Date
                self.update_row_colors()

            except sqlite3.Error as e:
                print("Database error:", e)
            finally:
                conn.close()

    def filter_tree_view_by_email(self, event=None):
        selected_email = self.email_combobox.get()

        # Clear the current TreeView
        for item in self.tree_view.get_children():
            self.tree_view.delete(item)

        # Connect to the database and fetch filtered data
        db_file = self.db_combo.get()
        if db_file:
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()
            try:
                if selected_email:
                    cursor.execute(
                        "SELECT ID, Date, Email, Equipment, DueDate, Status FROM equipment WHERE Email = ?",
                        (selected_email,))
                else:
                    cursor.execute("SELECT ID, Date, Email, Equipment, DueDate, Status FROM equipment")

                for row in cursor.fetchall():
                    self.tree_view.insert("", "end", values=row)

                # Update the row colors based on the Status and Due Date
                self.update_row_colors()

            except sqlite3.Error as e:
                print("Database error:", e)
            finally:
                conn.close()

    def set_custom_date(self):
        selected_items = self.tree_view.selection()
        if not selected_items:
            messagebox.showinfo("Info", "No item selected.")
            return

        def on_date_selected():
            chosen_date = cal.selection_get()
            for item in selected_items:
                self.update_item_date(item, chosen_date)
            date_window.destroy()

        # Create a new window for date selection
        date_window = tk.Toplevel(self)
        date_window.title("Select Date")

        # Calendar widget
        cal = Calendar(date_window, selectmode='day', date_pattern='yyyy-mm-dd')
        cal.pack(pady=10)

        # Confirmation button
        tk.Button(date_window, text="Ok", command=on_date_selected).pack(pady=10)

        self.sort_by_date(reverse=True)

    def clear_combobox_selection(self):
        self.equipment_combobox.set('')
        self.filter_tree_view_by_equipment()

    def on_combobox_enter(self, event):
        if not self.equipment_combobox.get().strip():  # Check if the combobox is empty
            self.clear_combobox_selection()
        if not self.email_combobox.get().strip():  # If the combobox is empty
            self.filter_tree_view_by_email()

    def handle_double_click(self, event):
        selected_items = self.tree_view.selection()
        for item in selected_items:
            item_values = self.tree_view.item(item, 'values')
            current_status = item_values[-1]

            # Check if the status is a digit preceded by '+', and update it to 'Returned +digit'
            if current_status.startswith('+') and current_status[1:].isdigit():
                new_status = 'Returned ' + current_status
                new_values = item_values[:-1] + (new_status,)
                self.tree_view.item(item, values=new_values)
                self.update_status_in_db(item, new_status)
                self.tree_view.item(item, tags=('returned',))
            elif not current_status.startswith('Returned'):
                # If the status is not starting with 'Returned', change it to 'Returned'
                new_status = 'Returned'
                new_values = item_values[:-1] + (new_status,)
                self.tree_view.item(item, values=new_values)
                self.update_status_in_db(item, new_status)
                self.tree_view.item(item, tags=('returned',))
            else:
                # If the status already starts with 'Returned', do nothing
                pass

        # Refresh the treeview to reflect the changes if any were made
        self.tree_view.update_idletasks()

    def process_date(self, date_str):
        if date_str is None:
            return None  # or some default value or handling as per your requirement
        try:
            return datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError as e:
            print(f"Error processing date: {date_str} - {e}")
            return None  # or some error handling

    def update_row_colors(self):
        for child in self.tree_view.get_children():
            item = self.tree_view.item(child)
            item_values = item['values']
            due_date_str = item_values[4]  # Adjust the index based on where the Due Date is in your values
            status = str(item_values[-1])  # Convert status to string in case it's not

            # Continue to the next iteration if the due date string is None or can't be parsed
            if not due_date_str or due_date_str.lower() == 'none':
                continue

            try:
                due_date = datetime.strptime(due_date_str, "%Y-%m-%d").date()
                current_date = datetime.now().date()
                days_diff = (current_date - due_date).days

                if status == "Not Returned" and due_date < current_date:
                    new_status = f"+{days_diff}"
                    new_values = tuple(item_values[:-1]) + (new_status,)
                    self.tree_view.item(child, values=new_values, tags=('overdue',))
                    self.tree_view.tag_configure('overdue', background='red')
                elif status.startswith('Returned'):
                    self.tree_view.item(child, tags=('returned',))
                    self.tree_view.tag_configure('returned', background=self.light_green)
                elif status.startswith('+') and due_date >= current_date:
                    new_values = tuple(item_values[:-1]) + ('Not Returned',)
                    self.tree_view.item(child, values=new_values, tags=('default',))
                    self.tree_view.tag_configure('default', background='')
                else:
                    self.tree_view.item(child, tags=('default',))
                    self.tree_view.tag_configure('default', background='')

            except ValueError as e:
                # Skip the row if the date can't be processed
                print(f"Skipped processing due to invalid date format: {due_date_str}")
                continue

    def update_item_date(self, item, new_date):
        item_values = self.tree_view.item(item, 'values')
        new_values = (item_values[0], new_date.strftime('%Y-%m-%d')) + item_values[2:]  # Update the date
        self.tree_view.item(item, values=new_values)  # Update the Treeview

        # Update the database
        id = item_values[0]  # Ensure this is the correct column for your ID
        db_file = self.db_combo.get()
        if db_file:
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()
            cursor.execute("UPDATE equipment SET Date = ? WHERE ID = ?", (new_date.strftime('%Y-%m-%d'), id))
            conn.commit()
            conn.close()

    def update_status_in_db(self, item, new_status):
        item_values = self.tree_view.item(item, 'values')
        id = item_values[0]  # Assuming the ID is in the first column
        db_file = self.db_combo.get()
        if db_file:
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()
            try:
                cursor.execute("UPDATE equipment SET Status = ? WHERE ID = ?", (new_status, id))
                conn.commit()
            except sqlite3.Error as e:
                print(f"Database error: {e}")
            finally:
                conn.close()
        self.app.refresh_pie_charts()

    def update_item_color_and_status(self, item, status):
        id = self.tree_view.item(item, 'values')[0]  # Assuming the ID is the first column
        db_file = self.db_combo.get()
        if db_file:
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()
            cursor.execute("UPDATE equipment SET Status = ? WHERE ID = ?", (status, id))
            conn.commit()
            conn.close()

        # Update the item color
        if status == 'Returned':
            self.tree_view.item(item, tags=('returned',))
        else:
            self.tree_view.item(item, tags=('default',))

        self.update_row_colors()  # Update all row colors

    def update_status_based_on_date(self, item, pseudo_current_date):
        item_values = self.tree_view.item(item, 'values')
        id = item_values[0]  # Assuming the ID is the first value
        due_date_str = item_values[4]  # Adjust the index as necessary
        due_date = datetime.strptime(due_date_str, "%Y-%m-%d").date()

        # Calculate the difference in days
        day_difference = (pseudo_current_date - due_date).days
        if day_difference > 0:
            new_status = f"Returned +{day_difference}"
        else:
            new_status = "Returned"

        # Update the item in the Treeview
        new_values = item_values[:-1] + (new_status,)
        self.tree_view.item(item, values=new_values)

        # Update the status in the database
        self.update_status_in_db(item, new_status)

        self.app.refresh_pie_charts()

    def update_db_list(self):
        # Update the ComboBox with available .db files
        self.db_combo['values'] = self.scan_for_databases()

    def combined_database_selection_handler(self, event):
        self.load_selected_db(event)
        self.on_database_selected(event)

    def load_selected_db(self, event=None):
        # Clear the existing TreeView entries
        for item in self.tree_view.get_children():
            self.tree_view.delete(item)

        # Get the selected database file
        db_file = self.db_combo.get()
        if not db_file:
            return  # No database selected

        # Construct the CSV filename based on the selected database name
        csv_file = db_file.replace('.db', '_users.csv')
        self.emails_file = csv_file  # Update the attribute to the new CSV file

        if not os.path.exists(csv_file):
            # If the CSV file does not exist, create an empty one
            pd.DataFrame(columns=['User Email']).to_csv(csv_file, index=False)

        self.load_emails_into_combobox()  # This method will load emails from the new CSV into the combobox

        # Connect to the SQLite database
        conn = sqlite3.connect(db_file)
        cursor = conn.cursor()

        try:
            # Fetch data from the database
            cursor.execute("SELECT ID, Date, Email, Equipment, DueDate, Status FROM equipment")
            rows = cursor.fetchall()

            # Populate the TreeView with the database data
            for row in rows:
                self.tree_view.insert("", "end", values=row)

            # Update row colors based on the status
            self.update_row_colors()

        except sqlite3.Error as e:
            print(f"Database error: {e}")
        finally:
            conn.close()

        self.load_equipment_entries()
        self.sort_by_date(reverse=True)

    def update_due_date(self, item, new_due_date):
        item_values = self.tree_view.item(item, 'values')
        current_date = datetime.now().date()
        new_due_date_obj = datetime.strptime(new_due_date, "%Y-%m-%d").date()

        # Calculate the status based on the new due date
        if new_due_date_obj < current_date:
            # If the new due date is past, calculate how many days overdue
            days_overdue = (current_date - new_due_date_obj).days
            new_status = f"+{days_overdue}"
        else:
            # If the new due date is today or in the future, keep or set the status to 'Not Returned'
            new_status = 'Not Returned' if 'Returned' not in item_values[-1] else item_values[-1]

        # Update the treeview with new due date and status
        new_values = item_values[:4] + (new_due_date,) + (new_status,)
        self.tree_view.item(item, values=new_values)

        # Update the database
        id = item_values[0]
        db_file = self.db_combo.get()
        if db_file:
            with sqlite3.connect(db_file) as conn:
                cursor = conn.cursor()
                cursor.execute("UPDATE equipment SET DueDate = ?, Status = ? WHERE ID = ?",
                               (new_due_date, new_status, id))
                conn.commit()

        # Update the colors and refresh the Treeview
        self.update_row_colors()
        self.tree_view.update_idletasks()

    def create_new_db(self):
        self.new_db_window = tk.Toplevel(self)
        self.new_db_window.title("Create New Database")
        self.new_db_window.geometry("270x120+100+150")
        background_color = "#d9d9d9"
        self.new_db_window.configure(background=background_color)
        ttk.Label(self.new_db_window, text="Enter base name for Database:").pack(pady=5)
        self.new_db_name_entry = tk.Entry(self.new_db_window)
        self.new_db_name_entry.pack(pady=5)
        ttk.Button(self.new_db_window, text="Create", command=self.add_new_db).pack(pady=10)

    def add_new_db(self):
        base_db_name = self.new_db_name_entry.get().strip()

        # Basic validation for the base database name
        if not base_db_name:
            messagebox.showerror("Error", "The base database name cannot be empty.")
            return

        # Find the next database index by checking existing files
        db_index = 0
        while os.path.exists(f'{base_db_name}_{db_index}.db'):
            db_index += 1

        # Database and CSV file names
        db_filename = f'{base_db_name}_{db_index}.db'
        csv_filename = f'{base_db_name}_{db_index}_users.csv'

        # Create the new database and its table
        conn = sqlite3.connect(db_filename)
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS equipment (
                              ID INTEGER PRIMARY KEY AUTOINCREMENT,
                              Date TEXT,
                              Email TEXT,
                              Equipment TEXT,
                              DueDate TEXT,
                              Status TEXT DEFAULT 'Not Returned')''')
        conn.commit()
        conn.close()

        # Create the corresponding CSV file
        if not os.path.exists(csv_filename):
            pd.DataFrame(columns=['User Email']).to_csv(csv_filename, index=False)

        # Update the ComboBox with the new database list and select the new database
        self.update_db_list()
        self.db_combo.set(db_filename)

        # Optionally, refresh the TreeView if needed
        self.load_selected_db(None)

        self.new_db_window.destroy()
        messagebox.showinfo("Info", f"Created new database and corresponding CSV: {db_filename}")

    def on_app_close(self):
        # Save the currently selected database name
        with open('last_db.txt', 'w') as f:
            f.write(self.db_combo.get())
        self.destroy()

    def sort_by_date(self, reverse=False):
        l = [(self.tree_view.set(k, "Date"), k) for k in self.tree_view.get_children('')]
        l.sort(reverse=reverse)  # True for descending order (newest to oldest), False for ascending

        # Rearrange items in sorted positions
        for index, (_, k) in enumerate(l):
            self.tree_view.move(k, '', index)

        # Switch the order for the next sort
        self.sort_reverse = not reverse

    def sort_by_column(self, col, reverse):
        l = [(self.tree_view.set(k, col), k) for k in self.tree_view.get_children('')]
        l.sort(reverse=reverse)  # True for descending, False for ascending

        # rearrange items in sorted positions
        for index, (val, k) in enumerate(l):
            self.tree_view.move(k, '', index)

        # reverse sort next time
        reverse = not reverse
        # Change the heading so that it will sort in the opposite direction
        self.tree_view.heading(col, command=lambda: self.sort_by_column(col, reverse))

    def on_database_selected(self, event=None):
        # Get the selected database file from the combobox
        db_file = self.db_combo.get()

        if db_file:
            # Extract the base filename without the extension and append '_users.csv' for the CSV filename
            base_name = os.path.splitext(os.path.basename(db_file))[0]
            self.emails_file = f"{base_name}_users.csv"

            # Ensure the CSV file exists; if not, create it with the appropriate column header
            if not os.path.exists(self.emails_file):
                df = pd.DataFrame(columns=['User Email'])
                df.to_csv(self.emails_file, index=False)

            # Load email addresses from the corresponding CSV file into the combobox
            self.load_emails_into_combobox()

            # Load the data from the selected database into the TreeView
            self.load_selected_db()

            # Trigger the update in the ConfidenceIndexTab
            if self.app.confidence_index_tab:
                self.app.confidence_index_tab.update_for_new_database(db_file, self.emails_file)
        else:
            # Optionally, handle the case when no database is selected
            print("No database selected")
            self.email_combobox['values'] = []  # Clear the email combobox values
            self.tree_view.delete(*self.tree_view.get_children())  # Clear the treeview entries

class ConfidenceIndexTab(tk.Frame):
    def __init__(self, parent, bg_color, equipment_tab):  # Add equipment_tab parameter here
        super().__init__(parent, background=bg_color)
        self.equipment_tab = equipment_tab  # Store the reference
        self.all_emails = []  # Initialize the attribute
        self.create_widgets()
        self.update_overall_chart()  # Call this to update the lower chart immediately
        self.populate_user_listbox()

        self.bind("<Visibility>", self.on_visibility)

    def on_visibility(self, event):
        if event.widget == self:
            self.update_overall_chart()

    def create_widgets(self):
        # Main container frame
        main_container = ttk.Frame(self)
        main_container.pack(fill=tk.BOTH, expand=True)

        # Frame for the listbox on the left
        self.user_listbox_frame = ttk.Frame(main_container)
        self.user_listbox_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Create and pack the listbox
        self.user_listbox = tk.Listbox(self.user_listbox_frame, exportselection=False)
        self.user_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.user_listbox.bind("<<ListboxSelect>>", self.on_user_select)

        # Create a scrollbar for the listbox and pack it
        scrollbar = ttk.Scrollbar(self.user_listbox_frame, orient='vertical', command=self.user_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill='y')
        self.user_listbox.config(yscrollcommand=scrollbar.set)

        # Container for the charts
        charts_container = ttk.Frame(main_container)
        charts_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Left chart container for the user-specific pie chart
        left_chart_frame = ttk.Frame(charts_container)
        left_chart_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Right chart container for the overall stats pie chart
        right_chart_frame = ttk.Frame(charts_container)
        right_chart_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Setup for the first pie chart (User-specific chart)
        self.figure1 = plt.Figure(figsize=(3, 2), dpi=100)
        self.ax1 = self.figure1.add_subplot(111)
        self.canvas1 = FigureCanvasTkAgg(self.figure1, left_chart_frame)
        self.canvas1.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.figure1.patch.set_facecolor('none')
        self.ax1.patch.set_facecolor('none')

        # Setup for the second pie chart (Overall stats chart)
        self.figure2 = plt.Figure(figsize=(3, 2), dpi=100)
        self.figure2.tight_layout()
        self.ax2 = self.figure2.add_subplot(111)
        self.canvas2 = FigureCanvasTkAgg(self.figure2, right_chart_frame)
        self.canvas2.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self.figure2.patch.set_facecolor('none')
        self.ax2.patch.set_facecolor('none')

        # Load user emails into the listbox
        self.load_user_emails()

        # Call the update methods for both charts during initialization
        self.update_user_chart(None)  # Assuming a None email parameter for a default message
        self.update_overall_chart()

    def load_user_emails(self, emails_file=None):
        if not emails_file:
            # If no file is provided, clear the list and return
            self.user_listbox.delete(0, tk.END)
            return

        try:
            df = pd.read_csv(emails_file)
            emails = df['User Email'].dropna().unique().tolist()
            self.user_listbox.delete(0, tk.END)
            for email in emails:
                self.user_listbox.insert(tk.END, email)
        except Exception as e:
            print(f"An error occurred while loading the emails: {e}")

    def update_for_new_database(self, db_file, emails_file):
        # Load the new set of emails into the listbox
        self.load_user_emails(emails_file)

        # Update the pie charts to reflect the new data
        self.populate_user_listbox()
        self.update_overall_chart()
        self.update_user_chart(None)  # Or pass the first email if needed

    def populate_user_listbox(self):
        user_emails = self.get_borrower_emails_from_db()
        self.user_listbox.delete(0, tk.END)
        for email in user_emails:
            self.user_listbox.insert(tk.END, email)

    def on_user_select(self, event):
        # Initialize the email variable
        email = None

        # Check if the event is None
        if event is not None:
            selection = event.widget.curselection()
            if selection:
                index = selection[0]
                email = event.widget.get(index)
        else:
            # If event is None, you might want to handle it differently
            # For example, select the first item in the listbox if it's not empty
            if self.user_listbox.size() > 0:
                self.user_listbox.selection_set(0)
                email = self.user_listbox.get(0)

        # Update the user-specific chart if an email is selected
        if email:
            self.update_user_chart(email)
        else:
            # Handle the case where no email is selected or listbox is empty
            print("No email selected or available.")

    # Update the upper pie chart with user-specific data
    def update_user_chart(self, email):
        self.ax1.clear()
        self.ax1.set_title('Trust Index', loc='center', fontweight='bold')  # Set the title for the pie chart

        if email is None:
            # If no user is selected, display the message and hide the axes
            self.ax1.axis('off')  # This hides the axes
            self.ax1.text(0.5, 0.5, 'Select a user to view trust index', horizontalalignment='center',
                          verticalalignment='center', transform=self.ax1.transAxes)
            self.canvas1.draw()
            return

        if email:
            user_identifier = email.split('@')[0]  # This gets the substring before '@'
            self.ax1.set_xlabel(f'{user_identifier}', fontsize=10, fontstyle='italic')
        else:
            self.ax1.set_xlabel('')

        # Initialize the count variables
        returned_on_time = 0
        total_items = 0

        db_file = self.equipment_tab.db_combo.get()

        if not os.path.exists(db_file):
            self.ax1.axis('off')  # Hide the axes if no database file is found
            self.ax1.text(0.5, 0.5, 'Database file not found.', horizontalalignment='center',
                          verticalalignment='center', transform=self.ax1.transAxes)
            self.canvas1.draw()
            return

        try:
            with sqlite3.connect(db_file) as conn:
                cursor = conn.cursor()

                # Count items returned on time by the user
                cursor.execute("SELECT COUNT(*) FROM equipment WHERE Email = ? AND Status = 'Returned'", (email,))
                returned_on_time = cursor.fetchone()[0]

                # Count total items taken by the user
                cursor.execute("SELECT COUNT(*) FROM equipment WHERE Email = ?", (email,))
                total_items = cursor.fetchone()[0]

        except sqlite3.Error as e:
            print(f"Database error: {e}")
            self.ax1.axis('off')  # Hide the axes if there's a database error
            self.ax1.text(0.5, 0.5, 'Database error.', horizontalalignment='center',
                          verticalalignment='center', transform=self.ax1.transAxes)
            self.canvas1.draw()
            return
        except Exception as e:
            print(f"An error occurred: {e}")
            self.ax1.axis('off')  # Hide the axes if there's a general error
            self.ax1.text(0.5, 0.5, 'An error occurred.', horizontalalignment='center',
                          verticalalignment='center', transform=self.ax1.transAxes)
            self.canvas1.draw()
            return

        trust_index = (returned_on_time / total_items * 100) if total_items > 0 else 0
        trust_data = [trust_index, 100 - trust_index]  # The trust score and the remaining percentage to 100

        # Only create the pie chart if we have valid data
        if total_items > 0:
            explode = (0.05, 0)  # only "explode" the first slice
            self.ax1.pie(trust_data, labels=['', ''], autopct='%1.1f%%', startangle=90, explode=explode, textprops={'fontsize': 10, 'color': 'white', 'weight': 'bold'}, shadow=True,
                         colors=['green', 'red'])
            self.ax1.axis('equal')
        else:
            self.ax1.axis('off')
            self.ax1.text(0.5, 0.5, 'No items to display', horizontalalignment='center',
                          verticalalignment='center', transform=self.ax1.transAxes)

        self.canvas1.draw()

    def update_overall_chart(self):
        self.ax2.clear()

        # Define 'labels' for the pie chart
        labels = ['Pending', 'Returned On Time', 'Returned Late', 'Currently Late']
        sizes = [0, 0, 0, 0]  # Initialize the sizes for each category

        db_file = self.equipment_tab.db_combo.get()
        if not db_file or not os.path.exists(db_file):
            self.display_message_on_chart(self.ax2, 'Database file not found.\nPlease check the database settings.')
            return

        try:
            with sqlite3.connect(db_file) as conn:
                cursor = conn.cursor()

                # Count items that are pending (not returned but not overdue)
                cursor.execute("SELECT COUNT(*) FROM equipment WHERE Status = 'Not Returned'")
                sizes[0] = cursor.fetchone()[0]

                # Count items that have been returned on time
                cursor.execute("SELECT COUNT(*) FROM equipment WHERE Status = 'Returned'")
                sizes[1] = cursor.fetchone()[0]

                # Count items that have been returned late
                cursor.execute("SELECT COUNT(*) FROM equipment WHERE Status LIKE 'Returned +%'")
                sizes[2] = cursor.fetchone()[0]

                # Count items that are late (not returned and overdue)
                cursor.execute("SELECT COUNT(*) FROM equipment WHERE Status LIKE '+%'")
                sizes[3] = cursor.fetchone()[0]

        except sqlite3.Error as e:
            print(f"Database error: {e}")
            self.display_message_on_chart(self.ax2, 'Database error.\nPlease check the database integrity.')
            return
        except Exception as e:
            print(f"An error occurred: {e}")
            self.display_message_on_chart(self.ax2, 'An error occurred.')
            return

        if sum(sizes) == 0:
            self.display_message_on_chart(self.ax2, 'No data available')
            return

        slice_colors = ['grey', 'green', 'red', 'orange']  # Define slice colors
        wedges, texts, autotexts = self.ax2.pie(
            sizes,
            autopct='%1.1f%%',
            startangle=90,
            colors=slice_colors,
            wedgeprops={'edgecolor': 'white', 'linewidth': 1.0, 'width': 0.3},
            textprops={'fontsize': 9, 'color': 'black', 'weight': 'bold'},
            shadow=True
        )

        # Improve the autopct positioning
        for autotext in autotexts:
            autotext.set_color('black')
            autotext.set_fontsize('10')

        # Place the legend below the pie chart
        self.ax2.legend(
            wedges,
            labels,
            title="",
            loc='upper center',
            bbox_to_anchor=(0.5, -0.1),
            frameon=False
        )

        # Equal aspect ratio ensures that pie is drawn as a circle
        self.ax2.axis('equal')
        self.ax2.set_title('Overall Equipment Status', loc='center', fontweight='bold')

        # Adjust layout to make room for the legend
        self.figure2.tight_layout()
        self.canvas2.draw()

    def get_borrower_emails_from_db(self):
        db_file = self.equipment_tab.db_combo.get()

        if not db_file or not os.path.exists(db_file):
            return []

        emails = set()
        try:
            conn = sqlite3.connect(db_file)
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT Email FROM equipment")
            for row in cursor.fetchall():
                emails.add(row[0])
        except sqlite3.Error as e:
            pass  # Optionally, handle the error e.g., logging it or displaying a user-friendly message
        finally:
            if conn:
                conn.close()

        return list(emails)

    def display_message_on_chart(self, axis, message):
        """ Helper function to display a message on a given chart axis. """
        axis.clear()
        axis.axis('off')
        axis.text(0.5, 0.5, message, horizontalalignment='center', verticalalignment='center', transform=axis.transAxes)
        self.canvas2.draw()

class TechTachoApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TechTacho - IT Equipment Tracker")

        self.style = ttk.Style()
        self.style.theme_use("clam")
        background_color = "#d9d9d9"
        self.configure(background=background_color)
        self.iconbitmap("icons.ico")

        self.tab_control = ttk.Notebook(self)

        # Existing tab
        self.equipment_tab = EquipmentTrackingTab(self.tab_control, background_color, self)
        self.tab_control.add(self.equipment_tab, text='Equipment Tracking')

        # New tab for Confidence Index
        self.confidence_index_tab = ConfidenceIndexTab(self.tab_control, background_color, self.equipment_tab)
        self.tab_control.add(self.confidence_index_tab, text='Confidence Index')
        # Select the tab to make it active
        self.tab_control.select(self.equipment_tab)
        self.tab_control.pack(expand=1, fill="both")

        # Ensure the 'on_app_close' method is set to handle the window close event
        self.protocol("WM_DELETE_WINDOW", self.on_app_close)

        # Load the last selected database after all initializations
        self.load_last_selected_db()
        self.confidence_index_tab.update_overall_chart()  # Update the overall chart

        self.tab_control = ttk.Notebook(self)
        self.tab_control.bind("<<NotebookTabChanged>>", self.on_tab_changed)

    def on_tab_changed(self, event):
        selected_tab = event.widget.select()
        tab_text = event.widget.tab(selected_tab, "text")
        if tab_text == "Confidence Index":
            self.confidence_index_tab.update_overall_chart()

    def refresh_pie_charts(self):
        self.confidence_index_tab.update_overall_chart()

        # Get the first email from the list or a default
        if self.confidence_index_tab.all_emails:
            default_email = self.confidence_index_tab.all_emails[0]
            self.confidence_index_tab.update_user_chart(default_email)
        else:
            # Handle the case where there are no emails, perhaps clearing the chart or showing default info
            self.confidence_index_tab.update_user_chart(None)

    def on_app_close(self):
        with open('last_db.txt', 'w') as f:
            f.write(self.equipment_tab.db_combo.get())
        self.destroy()

    def load_last_selected_db(self):
        if os.path.exists('last_db.txt'):
            with open('last_db.txt', 'r') as f:
                last_db = f.read().strip()
                if last_db and os.path.exists(last_db):
                    self.equipment_tab.db_combo.set(last_db)
                    self.equipment_tab.load_selected_db(None)
                    self.after(500, self.confidence_index_tab.populate_user_listbox)  # Delayed population
                else:
                    print("The last database file was not found.")

if __name__ == "__main__":
    app = TechTachoApp()
    app.mainloop()
