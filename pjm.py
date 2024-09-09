import subprocess
import sys
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, simpledialog
from tkcalendar import DateEntry
import pandas as pd
from tkinter import ttk
import os


class ScrollableFrame(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)

        # Create a canvas and a scrollbar
        self.canvas = tk.Canvas(self)
        self.scroll_y = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scroll_x = tk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)

        # Create a frame for the content
        self.content_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.content_frame, anchor="nw")

        # Configure the canvas to use the scrollbars
        self.canvas.configure(yscrollcommand=self.scroll_y.set, xscrollcommand=self.scroll_x.set)
        self.scroll_y.pack(side="right", fill="y")
        self.scroll_x.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)

        # Update the scrollregion to encompass the content frame
        self.content_frame.bind("<Configure>", self.on_frame_configure)

    def on_frame_configure(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))


# Step 1: File Selection
class Step1(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        tk.Label(self, text="Step 1: Select Excel File").pack(pady=10)

        # Dropdown menu for file selection
        self.file_options = tk.StringVar(self)
        self.file_options.set("Select File")  # Default value
        self.file_menu = tk.OptionMenu(self, self.file_options, *self.get_pjm_files())
        self.file_menu.pack(pady=10)

        self.btn_next = tk.Button(
            self,
            text="Next",
            bg="darkblue",
            fg="white",
            activebackground="navy",
            activeforeground="white",
            relief="flat",
            borderwidth=0,
            highlightthickness=0,
            state=tk.DISABLED,
            command=lambda: controller.show_frame(Step2),
            font=("Helvetica", 16, "bold"),  # Increase the font size and make it bold
            padx=20,  # Increase horizontal padding
            pady=10  # Increase vertical padding
        )
        self.btn_next.pack(pady=20)

        # Bind the dropdown selection event to a function
        self.file_options.trace('w', self.on_file_select)

    def get_pjm_files(self):
        # Directory where Excel files are located
        directory = "."  # or specify the directory where your files are located
        # List all files that contain 'pjm' (case-insensitive) in their name
        return [f for f in os.listdir(directory) if 'pjm' in f.lower() and f.lower().endswith(('.xlsx', '.xls'))]

    def on_file_select(self, *args):
        file_name = self.file_options.get()
        if file_name != "Select File":
            file_path = os.path.join(".", file_name)  # Update directory if needed
            try:
                self.controller.df = pd.read_excel(file_path, sheet_name="Data")
                print(self.controller.df)
                messagebox.showinfo("Success", "Excel file loaded successfully!")
                self.btn_next.config(state=tk.NORMAL)
            except Exception as e:
                messagebox.showwarning("Error", f"Failed to load file. Please try again.\n{e}")


# Step 2: Date Selection
class Step2(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        tk.Label(self, text="Step 2: Select Start and End Dates").pack(pady=10)

        self.date_frame = tk.Frame(self)
        self.date_frame.pack(pady=10)

        tk.Label(self.date_frame, text="Start Date:").grid(row=0, column=0, padx=10, pady=5)
        self.start_date_entry = DateEntry(self.date_frame, width=12, background='darkblue', foreground='white',
                                          borderwidth=2)
        self.start_date_entry.grid(row=0, column=1, padx=10, pady=5)

        tk.Label(self.date_frame, text="End Date:").grid(row=1, column=0, padx=10, pady=5)
        self.end_date_entry = DateEntry(self.date_frame, width=12, background='darkblue', foreground='white',
                                        borderwidth=2)
        self.end_date_entry.grid(row=1, column=1, padx=10, pady=5)

        self.btn_next = tk.Button(self, text="Next", command=self.filter_dates)
        self.btn_next.pack(pady=20)

        self.btn_back = tk.Button(self, text="Back", command=lambda: controller.show_frame(Step1))
        self.btn_back.pack()

    def filter_dates(self):
        if self.controller.df is None:
            messagebox.showwarning("Warning", "Please load an Excel file first.")
            return

        start_date = self.start_date_entry.get()
        end_date = self.end_date_entry.get()

        if start_date >= end_date:
            messagebox.showwarning("Warning", "End date must be greater than start date.")
            return

        self.controller.df_filtered_by_date = filter_by_dates(self.controller.df, start_date, end_date)

        if self.controller.df_filtered_by_date.empty:
            messagebox.showinfo("No Data", "No data found for the given date range.")
        else:
            messagebox.showinfo("Success", "Data filtered by dates successfully!")
            self.controller.show_frame(Step3)


# Step 3: Choose Filter Method
class Step3(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        tk.Label(self, text="Step 3: Choose Filtering Method").pack(pady=10)

        self.radio_var = tk.StringVar(value="state")
        tk.Radiobutton(self, text="Filter by State & County", variable=self.radio_var, value="state").pack(pady=5)
        tk.Radiobutton(self, text="Filter by Transmission Owner", variable=self.radio_var, value="transmission").pack(
            pady=5)

        self.btn_next = tk.Button(self, text="Next", command=self.apply_filter)
        self.btn_next.pack(pady=20)

        self.btn_back = tk.Button(self, text="Back", command=lambda: controller.show_frame(Step2))
        self.btn_back.pack()

    def apply_filter(self):
        filter_choice = self.radio_var.get()

        if filter_choice == 'state':
            self.controller.show_frame(Step4)
        elif filter_choice == 'transmission':
            self.controller.show_frame(Step5)


# Step 4: State and County Filtering


class Step4(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        tk.Label(self, text="Step 4: Select States and Counties").pack(pady=10)

        # Container for state frames
        # self.state_frame = tk.Frame(self)
        # self.state_frame.pack(pady=10, fill=tk.BOTH, expand=True)
        self.scrollable_frame = ScrollableFrame(self)
        self.scrollable_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        # Add buttons
        self.btn_back = tk.Button(self, text="Back", command=self.go_back)
        self.btn_back.pack(pady=10)

        self.btn_submit = tk.Button(self, text="Submit", command=self.submit_selection)
        self.btn_submit.pack(pady=10)

        # Dictionary to keep track of checkbutton states
        self.check_vars = {}
        self.state_vars = {}

    def load_state_county_options(self):
        # Clear any existing frames
        for widget in self.scrollable_frame.content_frame.winfo_children():
            widget.destroy()

        if self.controller.df_filtered_by_date is None:
            messagebox.showwarning("Warning", "Please filter by dates first.")
            return

        state_county_dict = {}
        unique_states = self.controller.df_filtered_by_date["State"].unique()

        for state in unique_states:
            counties = self.controller.df_filtered_by_date[self.controller.df_filtered_by_date["State"] == state][
                "County"].unique()
            state_county_dict[state] = counties

        # Create a frame for each state
        for state, counties in state_county_dict.items():
            # Frame for state and its checkbuttons
            state_frame = tk.Frame(self.scrollable_frame.content_frame)
            state_frame.pack(pady=5, fill=tk.X)

            # State checkbox
            state_var = tk.BooleanVar()
            state_cb = tk.Checkbutton(state_frame, variable=state_var,
                                      command=lambda s=state, var=state_var: self.toggle_counties(s, var))
            state_cb.grid(row=0, column=0, sticky="w")

            # State label
            state_label = tk.Label(state_frame, text=state, font=("Arial", 14, "bold"))
            state_label.grid(row=0, column=1, sticky="w")

            # Select all counties label
            select_all_label = tk.Label(state_frame, text="Select all counties", font=("Arial", 8))
            select_all_label.grid(row=0, column=2, sticky="w")

            # Frame to contain county checkbuttons
            county_frame = tk.Frame(self.scrollable_frame.content_frame)
            county_frame.pack(pady=5, fill=tk.X)

            # Create checkbuttons for each county
            for county in counties:
                county_var = tk.BooleanVar()
                cb = tk.Checkbutton(county_frame, text=county, variable=county_var)
                cb.pack(anchor="w")
                self.check_vars[(state, county)] = county_var

            # Store state variable
            self.state_vars[state] = state_var

    def toggle_counties(self, state, state_var):
        # Check or uncheck all county checkbuttons based on the state checkbox
        for (s, county), county_var in self.check_vars.items():
            if s == state:
                county_var.set(state_var.get())

    def submit_selection(self):
        # Create a dictionary to store selected states and counties
        selected_dict = {}

        # Collect selected states and counties
        for (state, county), county_var in self.check_vars.items():
            if county_var.get():  # Check if the county checkbox is selected
                if state not in selected_dict:
                    selected_dict[state] = []
                selected_dict[state].append(county)

        # Remove duplicates
        for state in selected_dict:
            selected_dict[state] = list(set(selected_dict[state]))

        # For demonstration, print the dictionary
        print("Selected Data Dictionary:")
        print(selected_dict)

        # Filter the DataFrame based on the selected states and counties
        if self.controller.df_filtered_by_date is not None:
            self.controller.cands = str(selected_dict)
            filtered_df = filter_by_states_counties(self.controller.df_filtered_by_date, selected_dict)
            self.controller.old_df = self.controller.df_filtered_by_date
            self.controller.old_s4 = self.controller.df_filtered_by_date
            self.controller.df_filtered_by_date = filtered_df
            self.controller.check_step_4 = True
            if not self.controller.check_step_5:
                self.controller.show_frame(Step5)
            else:
                self.controller.last_step = 4
                self.controller.show_frame(Step6)
                # Save the filtered DataFrame to an Excel file
                # self.controller.df_filtered_by_date.to_excel('output_sc.xlsx', index=False)
                # print("Filtered data has been saved to 'output_sc.xlsx'")
        else:
            messagebox.showwarning("Warning", "DataFrame is not loaded. Please filter by dates first.")

    def go_back(self):
        # Check if step 4 or step 5 was completed
        if self.controller.check_step_4 or self.controller.check_step_5:
            # Restore the original dataframe if step 4 or 5 was completed
            self.controller.df_filtered_by_date = self.controller.old_df
            self.controller.check_step_4 = False
            self.controller.check_step_5 = False

        # Navigate to Step 3
        self.controller.show_frame(Step3)

# Step 5: Transmission Owner Filtering
class Step5(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        tk.Label(self, text="Step 5: Select Transmission Owners").pack(pady=10)

        self.selected_transmission_owners = []

        self.owner_var_dict = {}
        self.checkbuttons = []

        # Create a canvas widget for scrolling
        self.canvas = tk.Canvas(self)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Create a scrollbar widget
        self.scrollbar = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Create a frame inside the canvas that will contain the checkbuttons
        self.owner_frame = tk.Frame(self.canvas)
        self.owner_frame.bind("<Configure>", self.on_frame_configure)  # Update scroll region when frame size changes

        # Add the frame to the canvas
        self.canvas.create_window((0, 0), window=self.owner_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.btn_filter = tk.Button(self, text="Apply Transmission Owner Filter",
                                    command=self.filter_by_transmission_owner)
        self.btn_filter.pack(pady=20)

        # Back button
        self.btn_back = tk.Button(self, text="Back", command=self.go_back)
        self.btn_back.pack()

    def on_frame_configure(self, event):
        # Update scroll region to encompass the frame
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def load_transmission_owner_options(self):
        for widget in self.owner_frame.winfo_children():
            widget.destroy()

        unique_transmission_owners = self.controller.df_filtered_by_date["Transmission Owner"].unique()

        for owner in unique_transmission_owners:
            var = tk.IntVar(value=0)
            self.owner_var_dict[owner] = var
            chk = tk.Checkbutton(self.owner_frame, text=owner, variable=var)
            chk.pack(anchor="w")
            self.checkbuttons.append(chk)

    def go_back(self):
        # Check if step 4 or step 5 was completed
        if self.controller.check_step_4 or self.controller.check_step_5:
            # Restore the original dataframe if step 4 or 5 was completed
            self.controller.df_filtered_by_date = self.controller.old_df
            self.controller.check_step_4 = False
            self.controller.check_step_5 = False

        # Navigate to Step 3
        self.controller.show_frame(Step3)

    def filter_by_transmission_owner(self):
        selected_transmission_owners = [owner for owner, var in self.owner_var_dict.items() if var.get() == 1]

        if not selected_transmission_owners:
            messagebox.showwarning("Warning", "No transmission owner selected.")
            return

        self.controller.df_final_filtered = filter_by_transmission_owner(self.controller.df_filtered_by_date,
                                                                         selected_transmission_owners)
        self.controller.old_df = self.controller.df_filtered_by_date
        self.controller.old_s5 = self.controller.df_filtered_by_date
        self.controller.df_filtered_by_date = self.controller.df_final_filtered

        if not self.controller.df_final_filtered.empty:
            # Define the output file path
            output_file_path = "output_tr.xlsx"
            self.controller.check_step_5 = True
            try:
                if not self.controller.check_step_4:
                    self.controller.show_frame(Step4)
                else:
                    # Save the filtered DataFrame to an Excel file
                    self.controller.last_step = 5
                    self.controller.show_frame(Step6)
                    # self.controller.df_filtered_by_date.to_excel(output_file_path, index=False)
                    # messagebox.showinfo("Success",
                    #                     f"Data filtered by Transmission Owner successfully! Saved to {output_file_path}.")
                    # print( self.controller.df_filtered_by_date)
            except Exception as e:
                messagebox.showwarning("Error", f"Failed to save file. Error: {e}")
        else:
            messagebox.showinfo("No Data", "No data found with the selected Transmission Owner filters.")

class Step6(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        tk.Label(self, text="Step 6: Choose Status List").pack(pady=10)

        self.radio_var = tk.StringVar(value="withdrawn")
        tk.Radiobutton(self, text="Withdrawn List", variable=self.radio_var, value="withdrawn").pack(pady=5)
        tk.Radiobutton(self, text="Queue List", variable=self.radio_var, value="queuelist").pack(pady=5)

        self.btn_next = tk.Button(self, text="Next", command=self.apply_status_filter)
        self.btn_next.pack(pady=20)

        # self.btn_back = tk.Button(self, text="Back", command=lambda: controller.show_frame(Step3))
        self.btn_back = tk.Button(self, text="Back", command=self.go_back)
        self.btn_back.pack()

    def apply_status_filter(self):
        status_choice = self.radio_var.get()
        status_values_withdrawn = [
            "Deactivated", "Withdrawn", "Retracted", "Suspended", "Partially in Service - Under Construction", "Under Construction"
        ]
        status_values_queue = [
            "Active", "Confirmed", "Engineering and Procurement", "In Service"
        ]

        if status_choice == 'withdrawn':
            self.controller.df_status_filtered = self.controller.df_filtered_by_date[
                self.controller.df_filtered_by_date["Status"].isin(status_values_withdrawn)
            ]
        elif status_choice == 'queuelist':
            self.controller.df_status_filtered = self.controller.df_filtered_by_date[
                self.controller.df_filtered_by_date["Status"].isin(status_values_queue)
            ]

        if not self.controller.df_status_filtered.empty:
            self.controller.show_frame(Step7)
            self.controller.sclog = status_choice
            # Proceed with the next step or operation
            # self.controller.df_status_filtered.to_excel('output_status_filtered.xlsx', index=False)  # Optional: Save to file
            # messagebox.showinfo("Success", "Data filtered by status successfully!")
            print(self.controller.df_status_filtered)
            # Optionally, transition to another step if needed
        else:
            messagebox.showinfo("No Data", "No data found with the selected Status filters.")

    def go_back(self):
        # Check if step 4 or step 5 was completed
        if self.controller.check_step_4 or self.controller.check_step_5:
            self.controller.check_step_4 = False
            self.controller.check_step_5 = False
            if self.controller.last_step == 4:
                self.controller.df_filtered_by_date = self.controller.old_s5
                self.controller.show_frame(Step3)
            elif self.controller.last_step == 5:
                self.controller.df_filtered_by_date = self.controller.old_s4
                self.controller.show_frame(Step3)

class Step7(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        self.label = tk.Label(self, text="Select an Option:")
        self.label.pack(pady=10)

        # Radio buttons for Capacity and Energy
        self.choice_var = tk.StringVar(value="Capacity")
        self.radio_capacity = tk.Radiobutton(self, text="Capacity", variable=self.choice_var, value="Capacity")
        self.radio_energy = tk.Radiobutton(self, text="Energy", variable=self.choice_var, value="Energy")
        self.radio_capacity.pack()
        self.radio_energy.pack()

        # Entry for Megawatt value
        self.label_mw = tk.Label(self, text="Enter Megawatt Value:")
        self.label_mw.pack(pady=10)
        self.mw_entry = tk.Entry(self)
        self.mw_entry.pack()

        # Next and Back buttons
        self.next_button = tk.Button(self, text="Next", command=self.process_data)
        self.next_button.pack(pady=10)
        self.back_button = tk.Button(self, text="Back", command=self.go_back)
        self.back_button.pack(pady=10)

    def process_data(self):
        # Get the selected option and MW value
        selected_option = self.choice_var.get()
        try:
            mw_value = float(self.mw_entry.get())
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter a valid number for megawatt value.")
            return

        # Filter data based on selection
        if selected_option == "Capacity":
            filtered_data = self.controller.df_status_filtered[
                self.controller.df_status_filtered['Capacity or Energy'] == 'Capacity'
            ]
            # Compare MW Capacity with the user input value
            filtered_data = filtered_data[filtered_data['MW Capacity'] >= mw_value]
        else:
            # Compare MW Energy with the user input value
            filtered_data = self.controller.df_status_filtered[
                self.controller.df_status_filtered['MW Energy'] >= mw_value
            ]

        # Save the filtered data to an Excel file
        if not filtered_data.empty:
            self.controller.df_filtered_by_mw = filtered_data
            self.controller.sop = selected_option
            self.controller.mwv = mw_value
            # Initialize Step8 now that df_filtered_by_mw is ready
            self.controller.frames["Step8"] = Step8(parent=self.controller.frames["Step3"].master,
                                                    controller=self.controller)
            self.controller.frames["Step8"].grid(row=0, column=0, sticky="nsew")
            self.controller.show_frame(Step8)
            # filtered_data.to_excel('output_filtered_mw.xlsx', index=False)
            messagebox.showinfo("Success", "Filtered data saved successfully!")
        else:
            messagebox.showinfo("No Data", "No data found with the selected Status filters.")

    def go_back(self):
        # Check if step 4 or step 5 was completed
        if self.controller.check_step_4 or self.controller.check_step_5:
            self.controller.check_step_4 = False
            self.controller.check_step_5 = False
            if self.controller.last_step == 4:
                self.controller.df_filtered_by_date = self.controller.old_s5
                self.controller.show_frame(Step3)
            elif self.controller.last_step == 5:
                self.controller.df_filtered_by_date = self.controller.old_s4
                self.controller.show_frame(Step3)

class Step8(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        self.label = tk.Label(self, text="Select Fuel Options:")
        self.label.pack(pady=10)

        # Select All checkbox
        self.select_all_var = tk.IntVar()
        self.select_all_checkbox = tk.Checkbutton(self, text="Select All", variable=self.select_all_var, command=self.toggle_select_all)
        self.select_all_checkbox.pack()

        # Frame to hold the fuel checkboxes
        self.fuel_frame = tk.Frame(self)
        self.fuel_frame.pack()

        # Dynamically generate fuel checkboxes
        self.fuel_vars = {}
        self.generate_fuel_checkboxes()

        # Next and Back buttons
        self.next_button = tk.Button(self, text="Next", command=self.process_data)
        self.next_button.pack(pady=10)
        self.back_button = tk.Button(self, text="Back", command=self.go_back)
        self.back_button.pack(pady=10)

    def generate_fuel_checkboxes(self):
        # Get unique fuel types from the filtered data
        unique_fuels = self.controller.df_filtered_by_mw['Fuel'].unique()

        for fuel in unique_fuels:
            var = tk.IntVar()
            checkbox = tk.Checkbutton(self.fuel_frame, text=fuel, variable=var)
            checkbox.pack(anchor="w")
            self.fuel_vars[fuel] = var

    def toggle_select_all(self):
        # Toggle all checkboxes based on the select all option
        for var in self.fuel_vars.values():
            var.set(self.select_all_var.get())

    def process_data(self):
        selected_fuels = [fuel for fuel, var in self.fuel_vars.items() if var.get() == 1]

        if self.select_all_var.get() == 1:
            selected_fuels = list(self.fuel_vars.keys())

        if not selected_fuels:
            messagebox.showerror("No Selection", "Please select at least one fuel type.")
            return

        # Filter the data based on selected fuels
        filtered_data = self.controller.df_filtered_by_mw[
            self.controller.df_filtered_by_mw['Fuel'].isin(selected_fuels)
        ]

        if not filtered_data.empty:
            self.controller.df_filtered_by_mw = filtered_data
            self.controller.fuel = selected_fuels
            # Initialize Step8 now that df_filtered_by_mw is ready
            self.controller.frames["Step9"] = Step9(parent=self.controller.frames["Step3"].master,
                                                    controller=self.controller)
            self.controller.frames["Step9"].grid(row=0, column=0, sticky="nsew")
            self.controller.show_frame(Step9)
            # filtered_data.to_excel('output_filtered_mw.xlsx', index=False)
            messagebox.showinfo("Success", "Filtered data saved successfully!")
        else:
            messagebox.showinfo("No Data", "No data found with the selected Status filters.")

        # Save the filtered data to an Excel file
        # filtered_data.to_excel('output_filtered_fuel.xlsx', index=False)
        # messagebox.showinfo("Success", "Filtered data saved successfully!")

    def go_back(self):
        # Check if step 4 or step 5 was completed
        if self.controller.check_step_4 or self.controller.check_step_5:
            self.controller.check_step_4 = False
            self.controller.check_step_5 = False
            if self.controller.last_step == 4:
                self.controller.df_filtered_by_date = self.controller.old_s5
                self.controller.show_frame(Step3)
            elif self.controller.last_step == 5:
                self.controller.df_filtered_by_date = self.controller.old_s4
                self.controller.show_frame(Step3)


class Step9(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        label = tk.Label(self, text="Step 9: Filter by System Impact Study Status")
        label.pack(pady=10, padx=10)

        # Create the "All" checkbox
        self.all_var = tk.BooleanVar()
        self.all_checkbox = tk.Checkbutton(self, text="All", variable=self.all_var, command=self.toggle_all_checkboxes)
        self.all_checkbox.pack(anchor="w")

        # Create checkboxes for unique System Impact Study Status values
        self.checkbox_vars = []
        unique_statuses = self.controller.df_filtered_by_mw['System Impact Study Status'].unique()

        for status in unique_statuses:
            var = tk.BooleanVar()
            checkbox = tk.Checkbutton(self, text=status, variable=var, command=self.check_all_checkbox_state)
            checkbox.pack(anchor="w")
            self.checkbox_vars.append((status, var))

        # Store unique statuses for logging
        self.controller.unique_statuses_log = unique_statuses

        # Create Next and Back buttons
        next_button = tk.Button(self, text="Next", command=self.process_data)
        next_button.pack(pady=20)

        back_button = tk.Button(self, text="Back", command=self.go_back)
        back_button.pack(pady=10)

    def toggle_all_checkboxes(self):
        # Set all checkboxes to the state of the "All" checkbox
        all_selected = self.all_var.get()
        for status, var in self.checkbox_vars:
            var.set(all_selected)

    def check_all_checkbox_state(self):
        # Check if all individual checkboxes are selected or not
        all_selected = all(var.get() for _, var in self.checkbox_vars)
        self.all_var.set(all_selected)

    def process_data(self):
        if self.all_var.get():
            # If "All" is selected, use all rows
            filtered_data = self.controller.df_filtered_by_mw
            self.controller.selected_status_log = str(self.controller.unique_statuses_log)
        else:
            # Otherwise, filter based on the selected checkboxes
            selected_statuses = [status for status, var in self.checkbox_vars if var.get()]
            filtered_data = self.controller.df_filtered_by_mw[
                self.controller.df_filtered_by_mw['System Impact Study Status'].isin(selected_statuses)
            ]
            self.controller.selected_status_log = str(selected_statuses)

        # Show the filtered data in a pop-up window for review
        self.show_dataframe_popup(filtered_data)

    def show_dataframe_popup(self, df):
        top = tk.Toplevel(self)
        top.title("Review Filtered Data")

        # Create a frame for the DataFrame display
        frame = tk.Frame(top)
        frame.pack(fill=tk.BOTH, expand=True)

        # Convert DataFrame to a string format
        df_string = df.to_string()

        # Create a Text widget to display the DataFrame
        text = tk.Text(frame, wrap=tk.NONE)
        text.insert(tk.END, df_string)
        text.config(state=tk.DISABLED)
        text.pack(fill=tk.BOTH, expand=True)

        # Confirmation buttons
        btn_frame = tk.Frame(top)
        btn_frame.pack(pady=10)

        btn_save = tk.Button(btn_frame, text="Save", command=lambda: self.ask_filename_and_save(df, top))
        btn_save.pack(side=tk.LEFT, padx=5)

        btn_cancel = tk.Button(btn_frame, text="Cancel", command=top.destroy)
        btn_cancel.pack(side=tk.LEFT, padx=5)

    def ask_filename_and_save(self, df, top):
        top.destroy()
        # Ask the user for the filename
        filename = simpledialog.askstring("Save File", "Enter the name for the output file (without extension):")
        if filename:
            filename += ".xlsx"

            # Save the filtered data to an Excel file
            with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='FilteredData', index=False)

                # Create BusInfo sheet with Project ID column
                bus_info_df = df[['Project ID']].copy()
                bus_info_df.to_excel(writer, sheet_name='BusInfo', index=False)

            # Log the details
            self.log_save_operation(filename, len(df), self.controller.selected_status_log, self.controller.fuel,
                                    self.controller.sop, self.controller.mwv, self.controller.cands,
                                    self.controller.sclog)

            messagebox.showinfo("Success", f"Data saved to {filename} successfully!")

            # Exit the current script and return to base.py
            self.exit_and_return_to_base()
        else:
            messagebox.showwarning("Warning", "File not saved. No filename provided.")

    def log_save_operation(self, filename, num_rows, selected_statuses, fuel, sop, mwv, cands, sclog):
        log_filename = "save_log.txt"
        log_data = {
            "RTO": "PJM",
            "Time of Save": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Filename": filename,
            "Number of Rows": num_rows,
            "States & Counties": cands,
            "Transmission Owners": sclog,
            "Selected Statuses": selected_statuses,
            "Fuel Types": fuel,
            "Capacity or Energy": sop,
            "Megawatt value": mwv
            # "Step 1 Options": self.controller.step1_options  # Assuming step 1 options are stored in the controller
        }

        # Append log data to the log file
        with open(log_filename, "a") as log_file:
            log_file.write(str(log_data) + "\n")

    def exit_and_return_to_base(self):
        self.controller.quit()  # Closes the current application

    def go_back(self):
        # Check if step 4 or step 5 was completed
        if self.controller.check_step_4 or self.controller.check_step_5:
            self.controller.check_step_4 = False
            self.controller.check_step_5 = False
            if self.controller.last_step == 4:
                self.controller.df_filtered_by_date = self.controller.old_s5
                self.controller.show_frame(Step3)
            elif self.controller.last_step == 5:
                self.controller.df_filtered_by_date = self.controller.old_s4
                self.controller.show_frame(Step3)


# Main Application Class
class DataFilterApp(tk.Tk):

    def __init__(self):
        tk.Tk.__init__(self)
        self.title("PJM")
        self.geometry("600x400")
        self.df = None
        self.df_filtered_by_date = None
        self.df_status_filtered = None
        self.df_final_filtered = None
        self.check_step_4 = False
        self.check_step_5 = False
        self.old_df = None
        self.old_s4 = None
        self.old_s5 = None
        self.last_step = None
        self.selected_status_log = None
        self.unique_statuses_log = None
        self.fuel = None
        self.sop = None
        self.mwv = None
        self.cands = None
        self.sclog = None
        self.df_filtered_by_mw = pd.DataFrame()
        container = tk.Frame(self)
        container.pack(fill="both", expand=True)

        self.frames = {}
        for F in (Step1, Step2, Step3, Step4, Step5, Step6, Step7):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame(Step1)  # Pass class object here

    def show_frame(self, page_class):
        page_name = page_class.__name__  # Get the class name as a string
        frame = self.frames[page_name]  # Use the class name string to access the frame
        if page_name == "Step4":
            frame.load_state_county_options()
        elif page_name == "Step5":
            frame.load_transmission_owner_options()
        frame.tkraise()


def filter_by_dates(df, start_date, end_date):
    df["Commercial Operation Milestone"] = pd.to_datetime(df["Commercial Operation Milestone"])
    return df[(df["Commercial Operation Milestone"] >= start_date) & (df["Commercial Operation Milestone"] <= end_date)]


def filter_by_states_counties(df, state_county_dict):
    mask = False
    for state, counties in state_county_dict.items():
        mask |= (df["State"] == state) & (df["County"].isin(counties))
    return df[mask]


def filter_by_transmission_owner(df, transmission_owners):
    return df[df["Transmission Owner"].isin(transmission_owners)]


if __name__ == "__main__":
    print(sys.executable)
    app = DataFilterApp()
    app.mainloop()
