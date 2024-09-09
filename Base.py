import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk, ImageDraw
import subprocess
import sys

class RTOSelectionApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RTO Selection")
        self.geometry("400x300")

        # Load button image with rounded corners
        self.round_button_image = self.create_round_button_image()

        # Initialize RTO selection
        self.create_rto_selection()

    def create_rto_selection(self):
        self.clear_window()

        self.label = tk.Label(self, text="Select RTO:", font=("Arial", 12))
        self.label.pack(pady=20)

        self.rto_var = tk.StringVar(value=getattr(self, 'selected_rto', ""))
        self.rto_dropdown = ttk.Combobox(self, textvariable=self.rto_var, values=["ISONE", "PJM", "NYISO"],
                                         state="readonly")
        self.rto_dropdown.pack(pady=10)

        self.next_button = tk.Button(self, text="Next", command=self.handle_rto_selection,
                                     bg="darkblue", fg="white", activebackground="navy", activeforeground="white",
                                     relief="flat", borderwidth=0, highlightthickness=0)
        self.next_button.config(image=self.round_button_image, compound="center")
        self.next_button.pack(pady=20)

    def handle_rto_selection(self):
        rto = self.rto_var.get()
        if rto == "PJM":
            self.run_pjm_script()
        elif rto == "ISONE":
            self.handle_isone_selection()
        elif rto == "NYISO":
            self.handle_nyiso_selection()
        else:
            messagebox.showwarning("Input Error", "Please select an RTO.")

    def run_pjm_script(self):
        # Run the PJM-specific script and capture the output
        try:
            process = subprocess.Popen(
                [r'S:\sasiproj2\venv\Scripts\python.exe', "pjm.py"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            stdout, stderr = process.communicate()  # Capture all output and error
            print(stdout)  # Print standard output
            if stderr:
                print(stderr)  # Print standard error if any
            process.wait()  # Wait for the subprocess to finish
            # self.quit()  # Close the current application window
        except Exception as e:
            messagebox.showerror("Error", f"Failed to run PJM script: {e}")

    def handle_isone_selection(self):
        # messagebox.showinfo("ISONE Selected", "PJM's done and dusted! If you want more, hire me, and let's make it happen. 😎💰")
        try:
            process = subprocess.Popen(
                [r'S:\sasiproj2\venv\Scripts\python.exe', "isone.py"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            stdout, stderr = process.communicate()  # Capture all output and error
            print(stdout)  # Print standard output
            if stderr:
                print(stderr)  # Print standard error if any
            process.wait()  # Wait for the subprocess to finish
            # self.quit()  # Close the current application window
        except Exception as e:
            messagebox.showerror("Error", f"Failed to run isone script: {e}")
        # Implement ISONE specific logic here

    def handle_nyiso_selection(self):
        # messagebox.showinfo("NYISO Selected", "PJM was quick! This one will be even faster. Ready to go? Hire me, and let's get paid! 💼🚀")
        try:
            process = subprocess.Popen(
                [r'S:\sasiproj2\venv\Scripts\python.exe', "nyiso.py"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            stdout, stderr = process.communicate()  # Capture all output and error
            print(stdout)  # Print standard output
            if stderr:
                print(stderr)  # Print standard error if any
            process.wait()  # Wait for the subprocess to finish
            # self.quit()  # Close the current application window
        except Exception as e:
            messagebox.showerror("Error", f"Failed to run NYISO script: {e}")
        # Implement NYISO specific logic here

    def clear_window(self):
        for widget in self.winfo_children():
            widget.pack_forget()

    def create_round_button_image(self):
        # Create an image with a rounded rectangle (PIL Image) for the button
        width, height = 150, 50
        radius = 25
        image = Image.new("RGBA", (width, height), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        draw.rounded_rectangle([0, 0, width, height], radius=radius, fill="darkblue")
        return ImageTk.PhotoImage(image)

if __name__ == "__main__":
    print(sys.executable)
    app = RTOSelectionApp()
    app.mainloop()
