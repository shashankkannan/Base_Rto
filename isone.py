import json
import sys
import tkinter as tk
from datetime import datetime
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from tkcalendar import DateEntry
import os


class MultiPhaseApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("NYISO Excel Processor")
        self.geometry("400x300")
        self.selected_file = None
        self.option = None
        self.start_date = None
        self.end_date = None
        self.filtered_data = None
        self.init_phase_one()
        self.selected_dict = {}
        self.state_county_dict = {}
        self.first_filter = None
        self.statescounty = {}
        self.zone = None
        self.fuel = None
        self.mwv = None
        self.core = None
        self.sis = None

    def clear_frame(self):
        for widget in self.winfo_children():
            widget.destroy()

    def init_phase_one(self):
        self.clear_frame()
        self.label = tk.Label(self, text="Select an Excel file:")
        self.label.pack(pady=10)

        self.file_var = tk.StringVar()
        self.file_dropdown = ttk.Combobox(self, textvariable=self.file_var, width=50)

        files = [f for f in os.listdir('.') if os.path.isfile(f) and 'isone' in f.lower() and f.endswith('.xlsx')]
        self.file_dropdown['values'] = files
        self.file_dropdown.pack(pady=10)

        self.next_button = tk.Button(self, text="Next", command=self.phase_two)
        self.next_button.pack(pady=20)

    def phase_two(self):
        if not self.file_var.get():
            messagebox.showerror("Error", "Please select a file.")
            return

        self.selected_file = self.file_var.get()
        self.clear_frame()

        self.label = tk.Label(self, text="Select the status:")
        self.label.pack(pady=10)

        self.option_var = tk.StringVar(value="Active")
        self.active_radio = tk.Radiobutton(self, text="Active", variable=self.option_var, value="Active")
        self.withdrawn_radio = tk.Radiobutton(self, text="Withdrawn", variable=self.option_var, value="Withdrawn")
        self.active_radio.pack(pady=5)
        self.withdrawn_radio.pack(pady=5)

        self.back_button = tk.Button(self, text="Back", command=self.init_phase_one)
        self.back_button.pack(side=tk.LEFT, padx=10, pady=20)

        self.next_button = tk.Button(self, text="Next", command=self.phase_three)
        self.next_button.pack(side=tk.RIGHT, padx=10, pady=20)

    def phase_three(self):
        self.option = self.option_var.get()
        self.clear_frame()

        self.label = tk.Label(self, text="Select start and end dates:")
        self.label.pack(pady=10)

        self.start_label = tk.Label(self, text="Start Date:")
        self.start_label.pack(pady=5)
        self.start_date_entry = DateEntry(self, date_pattern='y-mm-dd')
        self.start_date_entry.pack(pady=5)

        self.end_label = tk.Label(self, text="End Date:")
        self.end_label.pack(pady=5)
        self.end_date_entry = DateEntry(self, date_pattern='y-mm-dd')
        self.end_date_entry.pack(pady=5)

        self.back_button = tk.Button(self, text="Back", command=self.phase_two)
        self.back_button.pack(side=tk.LEFT, padx=10, pady=20)

        self.next_button = tk.Button(self, text="Next", command=self.process_data)
        self.next_button.pack(side=tk.RIGHT, padx=10, pady=20)

    def process_data(self):
        self.start_date = pd.to_datetime(self.start_date_entry.get_date())
        self.end_date = pd.to_datetime(self.end_date_entry.get_date())

        try:
            df = pd.read_excel(self.selected_file, sheet_name='Queue', skiprows=4)

            if self.option == "Active":
                df_filtered = df[df['W/ D Date'].notna()]
            else:
                df_filtered = df[df['W/ D Date'].isna()]

            df_filtered.loc[:, 'Op Date'] = pd.to_datetime(df_filtered['Op Date'], errors='coerce')
            self.start_date = pd.Timestamp(self.start_date)
            self.end_date = pd.Timestamp(self.end_date)

            df_filtered = df_filtered[
                (df_filtered['Op Date'] >= self.start_date) & (df_filtered['Op Date'] <= self.end_date)
                ]

            self.filtered_data = df_filtered  # Save the filtered data for the next phase
            self.phase_four()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def phase_four(self):
        self.clear_frame()

        self.label = tk.Label(self, text="Select an option:")
        self.label.pack(pady=10)

        self.option_var = tk.StringVar(value="State/County")
        self.state_county_radio = tk.Radiobutton(self, text="State/County", variable=self.option_var,
                                                 value="State/County")
        self.zone_radio = tk.Radiobutton(self, text="Zone", variable=self.option_var, value="Zone")
        self.state_county_radio.pack(pady=5)
        self.zone_radio.pack(pady=5)

        self.back_button = tk.Button(self, text="Back", command=self.back_from_phase_four)
        self.back_button.pack(side=tk.LEFT, padx=10, pady=20)

        self.next_button = tk.Button(self, text="Next", command=self.process_state_county_zone)
        self.next_button.pack(side=tk.RIGHT, padx=10, pady=20)

    def back_from_phase_four(self):
        self.phase_three()

    def process_state_county_zone(self):
        print("entered process stte zone")
        selected_option = self.option_var.get()
        self.first_filter = self.filtered_data
        if selected_option == "State/County":
            self.process_state_county()
        elif selected_option == "Zone":
            self.process_zone()
            # Implement Zone handling if needed
            pass

    def process_zone(self):
        # Get unique zones from the filtered data
        unique_zones = self.filtered_data['Zone'].dropna().unique()

        self.clear_frame()

        self.label = tk.Label(self, text="Select Zones:")
        self.label.pack(pady=10)

        # Create a canvas and scrollbar for scrolling
        canvas = tk.Canvas(self)
        scrollbar = tk.Scrollbar(self, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Dictionary to hold zone checkboxes
        self.zone_vars = {}

        for zone in unique_zones:
            zone_var = tk.BooleanVar()
            zone_cb = tk.Checkbutton(scrollable_frame, text=zone, variable=zone_var)
            zone_cb.pack(anchor=tk.W)
            self.zone_vars[zone] = zone_var

        self.back_button = tk.Button(self, text="Back", command=lambda: self.gobacktofirst())
        self.back_button.pack(side=tk.LEFT, padx=10, pady=20)

        self.next_button = tk.Button(self, text="Filter and Save", command=self.filter_and_save_zone)
        self.next_button.pack(side=tk.RIGHT, padx=10, pady=20)

    def filter_and_save_zone(self):
        selected_zones = [zone for zone, var in self.zone_vars.items() if var.get()]

        if selected_zones:
            self.zone = selected_zones
            filtered_data = self.filtered_data[self.filtered_data['Zone'].isin(selected_zones)]
            self.filtered_data = filtered_data
            # Print selected zones for debugging
            print("Selected Zones:", selected_zones)
            print("Filtered Data:\n", filtered_data)
            self.common_phase()

            # save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
            #                                          filetypes=[("Excel files", "*.xlsx")],
            #                                          title="Save the filtered data")
            # if save_path:
            #     filtered_data.to_excel(save_path, index=False)
            #     messagebox.showinfo("Success", f"Filtered data saved to {save_path}")
        else:
            messagebox.showwarning("No Selection", "No zones selected.")


    def process_state_county(self):
        # Process the filtered data to create a dictionary of states and counties
        self.state_county_dict = {}  # Store the state-county mapping as an instance variable
        for index, row in self.filtered_data.iterrows():
            state = row['State']
            counties = str(row['County'])

            if pd.isna(state) or pd.isna(counties):
                continue

            counties = counties.replace(' & ', '/').split('/')

            if state not in self.state_county_dict:
                self.state_county_dict[state] = []

            for county in counties:
                county = county.strip()
                if county not in ['NA', 'N/A', 'nan', '']:
                    if county not in self.state_county_dict[state]:
                        self.state_county_dict[state].append(county)

        self.show_state_county_result(self.state_county_dict)

    def show_state_county_result(self, state_county_dict):
        self.clear_frame()

        self.label = tk.Label(self, text="State-County Mapping:")
        self.label.pack(pady=10)

        # Create a canvas and scrollbar for scrolling
        canvas = tk.Canvas(self)
        scrollbar = tk.Scrollbar(self, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Dictionaries to hold state and county checkboxes
        self.state_vars = {}
        self.county_vars = {}

        for state, counties in state_county_dict.items():
            # State checkbox
            state_var = tk.BooleanVar()
            state_cb = tk.Checkbutton(scrollable_frame, text=state, font=('Arial', 10, 'bold'), variable=state_var,
                                      command=lambda s=state: self.toggle_state_selection(s))
            state_cb.pack(anchor=tk.W)
            self.state_vars[state] = state_var

            # County checkboxes
            county_frame = tk.Frame(scrollable_frame, padx=20)
            county_frame.pack(anchor=tk.W)

            county_vars_list = []
            for county in counties:
                county_var = tk.BooleanVar()
                county_cb = tk.Checkbutton(county_frame, text=county, variable=county_var,
                                           command=lambda s=state, c=county: self.toggle_county_selection(s, c))
                county_cb.pack(anchor=tk.W)
                county_vars_list.append(county_var)

            self.county_vars[state] = county_vars_list

        self.back_button = tk.Button(self, text="Back", command=lambda: self.gobacktofirst())
        self.back_button.pack(side=tk.LEFT, padx=10, pady=20)

        self.next_button = tk.Button(self, text="Filter and Save", command=self.filter_and_save)
        self.next_button.pack(side=tk.RIGHT, padx=10, pady=20)

    def toggle_state_selection(self, state):
        """Toggle selection of all counties in the state."""
        state_selected = self.state_vars[state].get()
        for county_var in self.county_vars[state]:
            county_var.set(state_selected)

    def toggle_county_selection(self, state, county):
        """Do not automatically select the entire state. Only ensure selected counties remain checked."""
        # No need to change the state selection here.

    def filter_and_save(self):
        selected_data = []
        selected_state_county_dict = {}  # Dictionary to hold the selected state-county pairs

        print("State selection status:", self.state_vars)

        for state, state_var in self.state_vars.items():
            if state_var.get():  # If the state is selected
                # Add all rows for this state and add all counties to the dictionary
                selected_data.append(self.filtered_data[self.filtered_data['State'] == state])
                selected_state_county_dict[state] = self.state_county_dict[state]
                print(f"Selected all counties in state: {state}")
            else:
                selected_counties = []
                for i, county_var in enumerate(self.county_vars[state]):
                    if county_var.get():  # If the county is selected
                        county_name = self.state_county_dict[state][i]
                        selected_counties.append(county_name)
                        selected_data.append(self.filtered_data[
                                                 (self.filtered_data['State'] == state) &
                                                 (self.filtered_data['County'].str.contains(county_name))
                                                 ])
                if selected_counties:
                    selected_state_county_dict[state] = selected_counties
                    print(f"Selected counties in {state}: {selected_counties}")

        # Print the final selected state-county dictionary
        print("Final selected state-county dictionary:")
        print(selected_state_county_dict)

        if selected_data:
            self.statescounty = selected_state_county_dict
            final_filtered_data = pd.concat(selected_data).drop_duplicates()
            self.filtered_data = final_filtered_data
            self.common_phase()
            # save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
            #                                          filetypes=[("Excel files", "*.xlsx")],
            #                                          title="Save the filtered data")
            # if save_path:
            #     final_filtered_data.to_excel(save_path, index=False)
            #     messagebox.showinfo("Success", f"Filtered data saved to {save_path}")
        else:
            messagebox.showwarning("No Selection", "No states or counties selected.")

    def common_phase(self):  # Store the filtered data for use in the next steps

        self.clear_frame()

        self.label = tk.Label(self, text="Choose Option:")
        self.label.pack(pady=10)

        # Radio buttons for Energy or Capacity
        self.option_var = tk.StringVar(value="Energy")
        self.energy_rb = tk.Radiobutton(self, text="Energy", variable=self.option_var, value="Energy")
        self.energy_rb.pack(anchor=tk.W)

        self.capacity_rb = tk.Radiobutton(self, text="Capacity", variable=self.option_var, value="Capacity")
        self.capacity_rb.pack(anchor=tk.W)

        # Input field for Megawatt value
        self.mw_label = tk.Label(self, text="Enter Megawatt Value:")
        self.mw_label.pack(pady=10)

        self.mw_entry = tk.Entry(self)
        self.mw_entry.pack(pady=10)

        # Buttons for Back and Filter & Save
        self.back_button = tk.Button(self, text="Back", command=lambda: self.gobacktofirst())
        self.back_button.pack(side=tk.LEFT, padx=10, pady=20)

        self.next_button = tk.Button(self, text="Filter and Save", command=self.filter_and_save_final)
        self.next_button.pack(side=tk.RIGHT, padx=10, pady=20)

    def filter_and_save_final(self):
        selected_option = self.option_var.get()
        try:
            megawatt_value = float(self.mw_entry.get())
        except ValueError:
            messagebox.showerror("Input Error", "Please enter a valid megawatt value.")
            return

        filtered_data = self.filtered_data.copy()

        # Filter based on the selected option
        if selected_option == "Capacity":
            self.core = "Capacity"
            filtered_data = filtered_data[filtered_data['Serv'] == 'CNR']
        else:
            self.core = "Energy"

        # Filter based on megawatt value
        filtered_data = filtered_data[filtered_data['Net MW'] >= megawatt_value]
        self.mwv = megawatt_value
        self.filtered_data = filtered_data
        self.select_fuel_types()
        # Print filtered data for debugging
        # print("Filtered Data based on option and megawatt value:\n", filtered_data)
        #
        # # Save the final filtered data to an Excel file
        # save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
        #                                          filetypes=[("Excel files", "*.xlsx")],
        #                                          title="Save the filtered data")
        # if save_path:
        #     filtered_data.to_excel(save_path, index=False)
        #     messagebox.showinfo("Success", f"Filtered data saved to {save_path}")

    def select_fuel_types(self):

        self.clear_frame()

        self.label = tk.Label(self, text="Select Fuel Types:")
        self.label.pack(pady=10)

        # Get unique fuel types
        unique_fuel_types = self.filtered_data['Fuel Type'].dropna().unique()

        # Scrollable form setup
        canvas = tk.Canvas(self)
        scrollbar = tk.Scrollbar(self, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Dictionary to hold the fuel type checkboxes
        self.fuel_type_vars = {}

        # Create checkboxes for each fuel type
        for fuel_type in unique_fuel_types:
            var = tk.BooleanVar()
            fuel_cb = tk.Checkbutton(scrollable_frame, text=fuel_type, variable=var)
            fuel_cb.pack(anchor=tk.W)
            self.fuel_type_vars[fuel_type] = var

        # Buttons for Back and Filter & Save
        self.back_button = tk.Button(self, text="Back", command=lambda: self.gobacktofirst())
        self.back_button.pack(side=tk.LEFT, padx=10, pady=20)

        self.next_button = tk.Button(self, text="Filter and Save", command=self.filter_and_save_fuel_type)
        self.next_button.pack(side=tk.RIGHT, padx=10, pady=20)

    def filter_and_save_fuel_type(self):
        # Get the selected fuel types
        selected_fuel_types = [fuel_type for fuel_type, var in self.fuel_type_vars.items() if var.get()]

        if selected_fuel_types:
            # Filter the data by selected fuel types
            filtered_data = self.filtered_data[self.filtered_data['Fuel Type'].isin(selected_fuel_types)]
            self.fuel = selected_fuel_types
            self.filtered_data = filtered_data
            self.select_sis_complete()
            # Print filtered data for debugging
        #     print("Filtered Data by selected Fuel Types:\n", filtered_data)
        #
        #     # Save the final filtered data to an Excel file
        #     save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
        #                                              filetypes=[("Excel files", "*.xlsx")],
        #                                              title="Save the filtered data")
        #     if save_path:
        #         filtered_data.to_excel(save_path, index=False)
        #         messagebox.showinfo("Success", f"Filtered data saved to {save_path}")
        # else:
        #     messagebox.showwarning("No Selection", "No fuel types selected.")

    def select_sis_complete(self):

        self.clear_frame()

        self.label = tk.Label(self, text="Select SIS Complete Status:")
        self.label.pack(pady=10)

        # Get unique SIS Complete values
        unique_sis_complete = self.filtered_data['SIS Complete'].dropna().unique()

        # Scrollable form setup
        canvas = tk.Canvas(self)
        scrollbar = tk.Scrollbar(self, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Dictionary to hold the SIS Complete checkboxes
        self.sis_complete_vars = {}

        # Create checkboxes for each SIS Complete option
        for sis_complete in unique_sis_complete:
            var = tk.BooleanVar()
            sis_complete_cb = tk.Checkbutton(scrollable_frame, text=sis_complete, variable=var)
            sis_complete_cb.pack(anchor=tk.W)
            self.sis_complete_vars[sis_complete] = var

        # Buttons for Back and Filter & Save
        self.back_button = tk.Button(self, text="Back", command=lambda: self.gobacktofirst())
        self.back_button.pack(side=tk.LEFT, padx=10, pady=20)

        self.next_button = tk.Button(self, text="Filter and Save", command=self.preview_filtered_data)
        self.next_button.pack(side=tk.RIGHT, padx=10, pady=20)

    def preview_filtered_data(self):
        # Get the selected SIS Complete statuses
        selected_sis_complete = [sis_complete for sis_complete, var in self.sis_complete_vars.items() if var.get()]

        if selected_sis_complete:
            # Filter the data by selected SIS Complete statuses
            self.sis = selected_sis_complete
            filtered_data = self.filtered_data[self.filtered_data['SIS Complete'].isin(selected_sis_complete)]
            self.filtered_data = filtered_data

            # Create a preview window to display filtered data
            self.preview_window = tk.Toplevel(self)
            self.preview_window.title("Preview Filtered Data")
            self.preview_window.geometry("800x400")

            frame = tk.Frame(self.preview_window)
            frame.pack(fill=tk.BOTH, expand=True)

            # Create a treeview to display the data
            tree = ttk.Treeview(frame)
            tree.pack(fill=tk.BOTH, expand=True)

            # Define columns
            columns = list(self.filtered_data.columns)
            tree["columns"] = columns
            tree["show"] = "headings"

            for col in columns:
                tree.heading(col, text=col)
                tree.column(col, width=100)

            # Insert data into the treeview
            for index, row in self.filtered_data.iterrows():
                tree.insert("", "end", values=list(row))

            # Add Save and Cancel buttons
            button_frame = tk.Frame(self.preview_window)
            button_frame.pack(pady=10)

            save_button = tk.Button(button_frame, text="Save", command=self.save_filtered_data)
            save_button.pack(side=tk.LEFT, padx=10)

            cancel_button = tk.Button(button_frame, text="Cancel", command=self.preview_window.destroy)
            cancel_button.pack(side=tk.RIGHT, padx=10)

        else:
            messagebox.showwarning("No Selection", "No SIS Complete status selected.")

    def save_filtered_data(self):
        # Ask the user for a file name to save the data
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")],
                                                 title="Save the filtered data")
        if save_path:
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                # Save the filtered data to the default sheet
                self.filtered_data.to_excel(writer, sheet_name='Queue', index=False)

                # Create a new dataframe for the 'Position' column
                bus_info_df = self.filtered_data[['Position']].copy()

                # Save the 'Position' data to a new sheet named 'BusInfo'
                bus_info_df.to_excel(writer, sheet_name='BusInfo', index=False)

                # Show success message
            messagebox.showinfo("Success", f"Filtered data saved to {save_path}")
            self.log_save_operation(
                save_path,
                len(self.filtered_data),
                self.statescounty if hasattr(self, 'statescounty') else None,
                self.zone if hasattr(self, 'zone') else None,
                self.fuel,
                self.mwv,
                self.core,
                self.sis
            )

        # Close the preview window after saving
        self.preview_window.destroy()
        sys.exit()

    def log_save_operation(self, filename, num_rows, counties, zone, fuel, mwv, core, sis):
        log_filename = "save_log.txt"
        log_data = {
            "RTO": "ISONE",
            "Time of Save": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Filename": filename,
            "Number of Rows": num_rows,
            "States & Counties": str(counties),
            "Zone": zone,
            "Fuel Types": fuel,
            "Megawatt Value": mwv,
            "Capacity or Energy": core,
            "SIS Status": sis
        }

        with open(log_filename, "a") as log_file:
            log_file.write(str(log_data) + "\n")

    # def filter_and_save_sis_complete(self):
    #     # Get the selected SIS Complete statuses
    #     selected_sis_complete = [sis_complete for sis_complete, var in self.sis_complete_vars.items() if var.get()]
    #
    #     if selected_sis_complete:
    #         # Filter the data by selected SIS Complete statuses
    #         filtered_data = self.filtered_data[self.filtered_data['SIS Complete'].isin(selected_sis_complete)]
    #
    #         # Print filtered data for debugging
    #         print("Filtered Data by selected SIS Complete values:\n", filtered_data)
    #
    #         # Save the final filtered data to an Excel file
    #         save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
    #                                                  filetypes=[("Excel files", "*.xlsx")],
    #                                                  title="Save the filtered data")
    #         if save_path:
    #             filtered_data.to_excel(save_path, index=False)
    #             messagebox.showinfo("Success", f"Filtered data saved to {save_path}")
    #     else:
    #         messagebox.showwarning("No Selection", "No SIS Complete status selected.")

    def gobacktofirst(self):
        print("go back first is called")
        self.filtered_data = self.first_filter
        self.phase_four()

if __name__ == "__main__":
    app = MultiPhaseApp()
    app.mainloop()
