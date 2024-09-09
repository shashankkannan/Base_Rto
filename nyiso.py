import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from datetime import datetime


def log_save_operation(filename, num_rows, counties, zone, fuel, mwv):
    """Log the save operation details in 'save_log.txt'."""
    log_filename = "save_log.txt"
    log_data = {
        "RTO": "NYISO",
        "Time of Save": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Filename": filename,
        "Number of Rows": num_rows,
        "States & Counties": str(counties),
        "Zone": zone,
        "Fuel Types": fuel,
        "Megawatt Value": mwv,
    }

    # Write the log to the log file
    with open(log_filename, "a") as log_file:
        log_file.write(str(log_data) + "\n")


class MultiPhaseApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MultiPhase Excel App")

        # Initialize main frame and variables
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill="both", expand=True)
        self.selected_file = tk.StringVar()
        self.phase = 1
        self.files_list = []
        self.state_county_selection = {}
        self.sc = {}
        self.zone = None
        self.avails = None
        self.fu = None
        self.mw = None
        self.firstfilter = None

        self.create_phase_1()

    def create_phase_1(self):
        """Phase 1: File Selection"""
        self.clear_frame()
        label = tk.Label(self.main_frame, text="Select an Excel file containing 'nyiso':")
        label.pack(pady=10)

        # Get all excel files containing "nyiso" (case-insensitive)
        self.files_list = [f for f in os.listdir() if f.lower().endswith('.xlsx') and 'nyiso' in f.lower()]
        if not self.files_list:
            messagebox.showwarning("No Files", "No 'nyiso' Excel files found!")
            self.root.quit()

        # Dropdown to choose file
        dropdown = ttk.Combobox(self.main_frame, textvariable=self.selected_file, values=self.files_list,
                                state="readonly")
        dropdown.pack(pady=10)

        # Navigation buttons
        next_button = tk.Button(self.main_frame, text="Next", command=self.create_phase_2)
        next_button.pack(side="right", padx=10, pady=20)

    def create_phase_2(self):
        """Phase 2: Choose List Type"""
        self.clear_frame()

        label = tk.Label(self.main_frame, text="Choose the list type:")
        label.pack(pady=10)

        # Radio buttons to choose between Queue List and Withdrawn List
        self.list_choice = tk.StringVar(value="Queue List")
        tk.Radiobutton(self.main_frame, text="Queue List", variable=self.list_choice, value="Queue List").pack(pady=5)
        tk.Radiobutton(self.main_frame, text="Withdrawn List", variable=self.list_choice, value="Withdrawn List").pack(
            pady=5)

        # Navigation buttons
        back_button = tk.Button(self.main_frame, text="Back", command=self.create_phase_1)
        back_button.pack(side="left", padx=10, pady=20)
        next_button = tk.Button(self.main_frame, text="Next", command=self.create_phase_3)
        next_button.pack(side="right", padx=10, pady=20)

    def create_phase_3(self):

        """Phase 3: Month and Year Selection and Filtering"""
        self.clear_frame()
        if self.list_choice.get() == "Withdrawn List":
            self.save_withdrawn_list()
            return  # Exit after saving

        label = tk.Label(self.main_frame, text="Select Date Range (Month and Year):")
        label.pack(pady=10)

        # Month selection
        tk.Label(self.main_frame, text="From Month:").pack(pady=5)
        self.from_month = ttk.Combobox(self.main_frame, values=[f'{i:02d}' for i in range(1, 13)], state="readonly")
        self.from_month.pack(pady=5)

        tk.Label(self.main_frame, text="From Year:").pack(pady=5)
        self.from_year = ttk.Combobox(self.main_frame, values=[str(i) for i in range(2000, datetime.now().year + 1)],
                                      state="readonly")
        self.from_year.pack(pady=5)

        # To Date selection
        tk.Label(self.main_frame, text="To Month:").pack(pady=5)
        self.to_month = ttk.Combobox(self.main_frame, values=[f'{i:02d}' for i in range(1, 13)], state="readonly")
        self.to_month.pack(pady=5)

        tk.Label(self.main_frame, text="To Year:").pack(pady=5)
        self.to_year = ttk.Combobox(self.main_frame, values=[str(i) for i in range(2000, datetime.now().year + 1)],
                                    state="readonly")
        self.to_year.pack(pady=5)

        # Navigation buttons
        back_button = tk.Button(self.main_frame, text="Back", command=self.gobacktofirst)
        back_button.pack(side="left", padx=10, pady=20)
        next_button = tk.Button(self.main_frame, text="Next", command=self.filter_dataframe)
        next_button.pack(side="right", padx=10, pady=20)

    def save_withdrawn_list(self):
        """Save the Withdrawn List to an Excel file and exit the program"""
        # Ask the user for the filename to save the Withdrawn list
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            try:
                # Assuming 'withdrawn_df' is the DataFrame you want to save
                withdrawn_df = pd.read_excel(self.selected_file.get(), sheet_name="Withdrawn")  # Replace with your actual DataFrame for the Withdrawn List

                # Save the DataFrame to Excel
                withdrawn_df.to_excel(save_path, index=False)

                print(f"Withdrawn list successfully saved to {save_path}. Exiting the program...")

                # Exit the program after saving
                sys.exit()

            except Exception as e:
                print(f"Failed to save the Withdrawn list: {e}")

    def filter_dataframe(self):
        """Filter the DataFrame based on selected month and year."""
        # Get selected from/to month and year
        from_year = self.from_year.get()
        from_month = self.from_month.get()
        to_year = self.to_year.get()
        to_month = self.to_month.get()

        # Create 'from_date' and 'to_date' with year and month only
        from_date = f"{from_year}-{from_month}"
        to_date = f"{to_year}-{to_month}"

        # Convert the 'from_date' and 'to_date' to datetime in '%Y-%m' format
        from_date = pd.to_datetime(from_date, format="%Y-%m")
        to_date = pd.to_datetime(to_date, format="%Y-%m")

        # Load selected Excel file and sheet based on user selection
        file_path = self.selected_file.get()
        sheet_name = "Interconnection Queue" if self.list_choice.get() == "Queue List" else "Withdrawn"

        # Read the Excel file
        try:
            self.df = pd.read_excel(file_path, sheet_name=sheet_name)
            self.firstfilter = self.df
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sheet: {e}")
            return

        # Filter out rows with 'Proposed COD' values that are 'NA', 'N/A', 'I/S', or NaN
        invalid_values = ['NA', 'N/A', 'I/S']
        self.df = self.df[~self.df['Proposed COD'].isin(invalid_values)].dropna(subset=['Proposed COD'])

        # Convert 'Proposed COD' to datetime using '%m/%Y' format
        self.df['Proposed COD'] = pd.to_datetime(self.df['Proposed COD'], errors='coerce', format="%m/%Y")

        # Filter rows where 'Proposed COD' falls within the date range
        self.filtered_df = self.df[(self.df['Proposed COD'] >= from_date) & (self.df['Proposed COD'] <= to_date)]

        if self.filtered_df.empty:
            messagebox.showinfo("No Data", "No data found for the selected date range.")
        else:
            self.create_phase_4()

    def create_phase_4(self):
        """Phase 4: Choose State/County or Zone"""
        self.clear_frame()

        label = tk.Label(self.main_frame, text="Choose Filter Type:")
        label.pack(pady=10)

        # Radio buttons for State/County or Zone
        self.filter_choice = tk.StringVar(value="State/County")
        tk.Radiobutton(self.main_frame, text="State/County", variable=self.filter_choice, value="State/County").pack(
            pady=5)
        tk.Radiobutton(self.main_frame, text="Zone", variable=self.filter_choice, value="Zone").pack(pady=5)

        # Navigation buttons
        back_button = tk.Button(self.main_frame, text="Back", command=self.gobacktofirst)
        back_button.pack(side="left", padx=10, pady=20)
        next_button = tk.Button(self.main_frame, text="Next", command=self.process_next_phase)
        next_button.pack(side="right", padx=10, pady=20)

    def process_next_phase(self):
        """Process the selection based on the chosen phase (State/County or Zone)."""
        phase = self.filter_choice.get()
        if phase == "State/County":
            self.create_phase_5()
        elif phase == "Zone":
            self.create_zone_phase()

    def create_zone_phase(self):
        """Phase 6: Zone Selection"""
        self.clear_frame()

        # Extract unique zones from filtered data
        unique_zones = self.filtered_df['Z'].unique()

        label = tk.Label(self.main_frame, text="Select Zones:")
        label.pack(pady=10)

        # Frame for zone checkboxes
        zone_frame = tk.Frame(self.main_frame)
        zone_frame.pack(fill="both", expand=True)

        self.zone_vars = {}
        for zone in unique_zones:
            zone_var = tk.BooleanVar()
            self.zone_vars[zone] = zone_var
            tk.Checkbutton(zone_frame, text=zone, variable=zone_var).pack(anchor="w")

        # Navigation buttons
        back_button = tk.Button(self.main_frame, text="Back", command=self.gobacktofirst)
        back_button.pack(side="left", padx=10, pady=20)
        next_button = tk.Button(self.main_frame, text="Next", command=self.process_zone_selection)
        next_button.pack(side="right", padx=10, pady=20)

    def process_zone_selection(self):
        """Process the selected zones and save the filtered data."""
        selected_zones = [zone for zone, var in self.zone_vars.items() if var.get()]

        if selected_zones:
            # Filter dataframe based on selected zones
            zone_filtered_df = self.filtered_df[self.filtered_df['Z'].isin(selected_zones)]
            self.filtered_df = zone_filtered_df

            if not self.filtered_df.empty:
                self.zone = selected_zones
                self.create_availability_phase()
            # Ask the user where to save the filtered data
            # save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            # if save_path:
            #     zone_filtered_df.to_excel(save_path, index=False)
            #     messagebox.showinfo("Success", "Filtered data saved successfully!")
        else:
            messagebox.showwarning("No Selection", "No zones selected.")

    def create_phase_5(self):
        """Phase 5: State and County Selection (if State/County is chosen)"""
        if self.filter_choice.get() == "State/County":
            self.clear_frame()

            label = tk.Label(self.main_frame, text="Select States and Counties:")
            label.pack(pady=10)

            # Create a canvas for the scrollable area
            canvas = tk.Canvas(self.main_frame)
            canvas.pack(side="left", fill="both", expand=True)

            # Create a scrollbar and link it to the canvas
            scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=canvas.yview)
            scrollbar.pack(side="right", fill="y")

            # Create a frame inside the canvas for checkboxes
            checkbox_frame = tk.Frame(canvas)
            checkbox_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            # Add the frame to the canvas
            canvas.create_window((0, 0), window=checkbox_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            # Create dictionary for State -> Counties
            state_county_dict = {}
            unique_states = self.filtered_df['State'].unique()
            for state in unique_states:
                counties = self.filtered_df[self.filtered_df['State'] == state]['County'].unique()
                state_county_dict[state] = counties

            # Display States and Counties with checkboxes
            self.state_vars = {}
            self.county_vars = {}
            for state, counties in state_county_dict.items():
                state_var = tk.BooleanVar()
                state_checkbox = tk.Checkbutton(checkbox_frame, text=state, variable=state_var,
                                                command=lambda s=state: self.select_state(s))
                state_checkbox.pack(anchor='w')
                self.state_vars[state] = state_var

                self.county_vars[state] = {}
                for county in counties:
                    county_var = tk.BooleanVar()
                    county_checkbox = tk.Checkbutton(checkbox_frame, text=f"  {county}", variable=county_var,
                                                     command=lambda s=state, c=county: self.select_county(s, c))
                    county_checkbox.pack(anchor='w')
                    self.county_vars[state][county] = county_var

            # Navigation buttons
            back_button = tk.Button(self.main_frame, text="Back", command=self.gobacktofirst)
            back_button.pack(side="left", padx=10, pady=20)
            next_button = tk.Button(self.main_frame, text="Next", command=self.process_state_county_selection)
            next_button.pack(side="right", padx=10, pady=20)

    def select_state(self, state):
        """Select all counties if a state checkbox is selected."""
        state_selected = self.state_vars[state].get()
        for county, var in self.county_vars[state].items():
            var.set(state_selected)

    def select_county(self, state, county):
        """Ensure if specific counties are selected, the state checkbox remains unchecked."""
        # If not all counties are selected, uncheck the state checkbox
        if not all(var.get() for var in self.county_vars[state].values()):
            self.state_vars[state].set(False)
        # If all counties are selected, check the state checkbox
        elif all(var.get() for var in self.county_vars[state].values()):
            self.state_vars[state].set(True)

        # Update the dictionary every time a county is selected
        self.update_state_county_dict()

    def update_state_county_dict(self):
        """Update the state-county dictionary based on the current selection."""
        self.state_county_dict = {}  # Initialize an empty dictionary

        # Loop through the state checkboxes
        for state, state_var in self.state_vars.items():
            selected_counties = [
                county for county, county_var in self.county_vars[state].items() if county_var.get()
            ]
            if selected_counties:
                # Add the state with selected counties to the dictionary
                self.state_county_dict[state] = selected_counties

        # Print the current state-county dictionary for verification
        print("Current State-County Selection Dictionary:", self.state_county_dict)

    def process_state_county_selection(self):
        """Process the selected states and counties and save the filtered data."""
        selected_data = []

        # Collect selected states and counties from the dictionary
        for state, counties in self.state_county_dict.items():
            # Filter dataframe based on selected state and counties
            state_filtered_df = self.filtered_df[(self.filtered_df['State'] == state) &
                                                 (self.filtered_df['County'].isin(counties))]
            selected_data.append(state_filtered_df)

        if selected_data:
            # Combine all the selected data
            combined_filtered_df = pd.concat(selected_data)
            self.sc = self.state_county_dict
            self.filtered_df = combined_filtered_df

            if not self.filtered_df.empty:
                self.create_availability_phase()
            # Ask the user where to save the filtered data
            # save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            # if save_path:
            #     combined_filtered_df.to_excel(save_path, index=False)
            #     messagebox.showinfo("Success", "Filtered data saved successfully!")
        else:
            messagebox.showwarning("No Selection", "No states or counties selected.")

    def create_availability_phase(self):
        """Phase 7: Availability of Studies Selection"""
        self.clear_frame()

        # Extract unique Availability of Studies values from filtered data
        unique_availability = self.filtered_df['Availability of Studies'].unique()

        label = tk.Label(self.main_frame, text="Select Availability of Studies:")
        label.pack(pady=10)

        # Frame for Availability of Studies checkboxes
        availability_frame = tk.Frame(self.main_frame)
        availability_frame.pack(fill="both", expand=True)

        self.availability_vars = {}
        for availability in unique_availability:
            availability_var = tk.BooleanVar()
            self.availability_vars[availability] = availability_var
            tk.Checkbutton(availability_frame, text=availability, variable=availability_var).pack(anchor="w")

        # Navigation buttons
        back_button = tk.Button(self.main_frame, text="Back", command=self.gobacktofirst)
        back_button.pack(side="left", padx=10, pady=20)
        next_button = tk.Button(self.main_frame, text="Next", command=self.process_availability_selection)
        next_button.pack(side="right", padx=10, pady=20)

    def process_availability_selection(self):
        """Filter the DataFrame based on selected Availability of Studies and proceed to the next phase."""
        selected_availabilities = [availability for availability, var in self.availability_vars.items() if var.get()]
        if not selected_availabilities:
            messagebox.showwarning("No Selection", "No Availability of Studies selected!")
            return
        self.avails = selected_availabilities
        self.filtered_df = self.filtered_df[self.filtered_df['Availability of Studies'].isin(selected_availabilities)]
        self.create_type_fuel_phase()

    def create_type_fuel_phase(self):
        """Phase 8: Type/Fuel Selection"""
        self.clear_frame()

        # Extract unique Type/Fuel values from filtered data
        unique_types_fuel = self.filtered_df['Type/ Fuel'].unique()

        label = tk.Label(self.main_frame, text="Select Type/Fuel:")
        label.pack(pady=10)

        # Frame for Type/Fuel checkboxes
        type_fuel_frame = tk.Frame(self.main_frame)
        type_fuel_frame.pack(fill="both", expand=True)

        self.type_fuel_vars = {}
        for type_fuel in unique_types_fuel:
            type_fuel_var = tk.BooleanVar()
            self.type_fuel_vars[type_fuel] = type_fuel_var
            tk.Checkbutton(type_fuel_frame, text=type_fuel, variable=type_fuel_var).pack(anchor="w")

        # Navigation buttons
        back_button = tk.Button(self.main_frame, text="Back", command=self.gobacktofirst)
        back_button.pack(side="left", padx=10, pady=20)
        next_button = tk.Button(self.main_frame, text="Next", command=self.process_type_fuel_selection)
        next_button.pack(side="right", padx=10, pady=20)

    def process_type_fuel_selection(self):
        """Filter the DataFrame based on selected Type/Fuel."""
        selected_types_fuel = [type_fuel for type_fuel, var in self.type_fuel_vars.items() if var.get()]
        if not selected_types_fuel:
            messagebox.showwarning("No Selection", "No Type/Fuel selected!")
            return

        self.filtered_df = self.filtered_df[self.filtered_df['Type/ Fuel'].isin(selected_types_fuel)]

        if not self.filtered_df.empty:
            self.fu = selected_types_fuel
            self.create_megawatt_phase()
        else:
            messagebox.showinfo("No Data", "No data available after filtering by Type/Fuel.")

    def create_megawatt_phase(self):
        """Phase 8: Megawatt Value Input"""
        self.clear_frame()

        label = tk.Label(self.main_frame, text="Enter the minimum Megawatt value (SP (MW)):")
        label.pack(pady=10)

        self.megawatt_entry = tk.Entry(self.main_frame)
        self.megawatt_entry.pack(pady=5)

        # Navigation buttons
        back_button = tk.Button(self.main_frame, text="Back", command=self.gobacktofirst)
        back_button.pack(side="left", padx=10, pady=20)
        next_button = tk.Button(self.main_frame, text="Next", command=self.process_megawatt_selection)
        next_button.pack(side="right", padx=10, pady=20)

    def process_megawatt_selection(self):
        """Filter the DataFrame based on the Megawatt input and show the final data."""
        megawatt_value = self.megawatt_entry.get()
        try:
            megawatt_value = float(megawatt_value)
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter a valid number for Megawatt value.")
            return

        # Filter the DataFrame based on 'SP (MW)' column
        self.filtered_df = self.filtered_df[self.filtered_df['SP (MW)'] >= megawatt_value]

        if not self.filtered_df.empty:
            self.mw = megawatt_value
            self.show_filtered_data()
        else:
            messagebox.showinfo("No Data", "No data available after filtering by Megawatt value.")

    def show_filtered_data(self):
        """Show the filtered data with options to save or cancel."""
        self.clear_frame()

        # Create a text widget to display the DataFrame
        text = tk.Text(self.main_frame, wrap="none")
        text.insert("1.0", self.filtered_df.to_string(index=False))
        text.pack(expand=True, fill="both")

        # Scrollbars for the text widget
        x_scroll = tk.Scrollbar(self.main_frame, orient='horizontal', command=text.xview)
        x_scroll.pack(side='bottom', fill='x')
        y_scroll = tk.Scrollbar(self.main_frame, orient='vertical', command=text.yview)
        y_scroll.pack(side='right', fill='y')
        text.configure(xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set)

        # Save and Cancel buttons
        save_button = tk.Button(self.main_frame, text="Save", command=self.save_filtered_data)
        save_button.pack(side="left", padx=10, pady=20)

        cancel_button = tk.Button(self.main_frame, text="Cancel", command=self.root.quit)
        cancel_button.pack(side="right", padx=10, pady=20)

    def save_filtered_data(self):
        """Save the filtered DataFrame to an Excel file."""
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            try:
                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    # Save the filtered DataFrame to the default sheet
                    self.filtered_df.to_excel(writer, sheet_name='Queue', index=False)

                    # Create a new DataFrame for the 'Position' column
                    bus_info_df = self.filtered_df[['Queue Pos.']].copy()

                    # Save the 'Position' column to a new sheet named 'BusInfo'
                    bus_info_df.to_excel(writer, sheet_name='BusInfo', index=False)
                num_rows = len(self.filtered_df)
                log_save_operation(save_path, num_rows, self.sc, self.zone, self.fu, self.mw)
                messagebox.showinfo("Success", "Filtered data saved successfully!")
                sys.exit()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {e}")

    def clear_frame(self):
        """Clear all widgets from the frame."""
        for widget in self.main_frame.winfo_children():
            widget.destroy()

    def gobacktofirst(self):
        print("go back first is called")
        self.filtered_df = self.firstfilter
        self.create_phase_2()

if __name__ == "__main__":
    root = tk.Tk()
    app = MultiPhaseApp(root)
    root.mainloop()
