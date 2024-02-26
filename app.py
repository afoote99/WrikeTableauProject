#BUGS:
#Chloe shows up twice
#if not selected in the list, shouldnt show up on the final excel sheet -> right now it is pulling many people not included?

import tkinter as tk
from tkinter import simpledialog, messagebox, filedialog, ttk, Checkbutton, Scrollbar, VERTICAL, HORIZONTAL
import requests
import webbrowser
from datetime import datetime
import pandas as pd

class TimeAudit(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Time Audit Generator")
        self.geometry("500x600")
        self.configure(bg='#f0f0f0')
        self.access_token = None
        self.user_ids = {}
        self.user_checks = {}
        self.initialize_ui_components()

    def initialize_ui_components(self):
        self.setup_buttons()
        self.setup_search_entry()
        self.setup_checkbox_frame()

    def setup_search_entry(self):
        # Initially, don't pack the search entry; it will be packed when relevant
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(self, textvariable=self.search_var, font=('Arial', 12), bd=1, relief='solid', width=50)
        self.search_entry.bind("<KeyRelease>", self.filter_list)

    def setup_buttons(self):
        # Change button color to blue (#007bff)
        button_style = {'bg': '#007bff', 'fg': 'white', 'bd': 0, 'padx': 20, 'pady': 10, 'font': ('Arial', 12, 'bold')}

        # Only initialize buttons here, don't pack them
        self.auth_button = tk.Button(self, text="Authorize with Wrike", command=self.generate_access_code, **button_style)
        self.fetch_people_button = tk.Button(self, text="Show Wrike People", command=self.fetch_wrike_people, **button_style)
        self.upload_button = tk.Button(self, text="Upload Tableau File", command=self.upload_tableau_file, **button_style)
        self.generate_button = tk.Button(self, text="Generate Report", command=self.generate_report, **button_style)

        # Initially display only the 'Authorize with Wrike' button
        self.auth_button.pack(pady=20)

    def setup_checkbox_frame(self):
        self.checkbox_frame = tk.Frame(self)
        self.checkbox_frame.pack(fill="both", expand=True)
        
        self.canvas = tk.Canvas(self.checkbox_frame, bg='#f0f0f0')
        self.canvas.pack(side="left", fill="both", expand=True)
        
        self.scrollbar = Scrollbar(self.checkbox_frame, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side="right", fill="y")
        
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        # Now, bind the mousewheel scroll event to the canvas, and adjust to bind to the frame if needed
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)
        # Optionally, bind directly to the scrollable_frame if the canvas binding doesn't work as expected
        # self.scrollable_frame.bind("<MouseWheel>", self.on_mousewheel)

    def on_mousewheel(self, event):
        """Scroll handler for mouse wheel over the canvas."""
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    
    
    def setup_buttons(self):
        # Common style for all buttons, changing bg to Royal Blue '#4169E1'
        button_style = {'bg': '#4169E1', 'fg': 'white', 'bd': 0, 'padx': 20, 'pady': 10, 'font': ('Arial', 12, 'bold')}

        # Authorization button
        self.auth_button = tk.Button(self, text="Authorize with Wrike", command=self.generate_access_code, **button_style)
        self.auth_button.pack(pady=20)

        # Fetch Wrike People button - hidden initially
        self.fetch_people_button = tk.Button(self, text="Show Wrike People", command=self.fetch_wrike_people, **button_style)

        # Upload Tableau File button - hidden initially
        self.upload_button = tk.Button(self, text="Upload Tableau File", command=self.upload_tableau_file, **button_style)

        # Generate Report button - hidden initially
        self.generate_button = tk.Button(self, text="Generate Report", command=self.generate_report, **button_style)

    def filter_list(self, event=None):
        search_term = self.search_var.get().lower()

        # Clear current checkboxes
        for widget in self.scrollable_frame.winfo_children():
            widget.pack_forget()

        # Filter and sort names based on the search term
        filtered_names = [name for name in self.user_checks.keys() if search_term in name.lower()]
        sorted_filtered_names = sorted(filtered_names)

        # If search term is empty, reset to original list
        if not search_term:
            sorted_filtered_names = sorted(self.user_checks.keys())

        # Repack checkboxes for filtered and sorted names
        for name in sorted_filtered_names:
            var = self.user_checks[name]
            chk = tk.Checkbutton(self.scrollable_frame, text=name, variable=var, bg='#f0f0f0')
            chk.pack(anchor="w")


    def generate_access_code(self):
        client_id = "Gj7SNTXf"
        client_secret = "RffPhFeCG4Xtmi8LhvXQy60yT35Dkw6WnanMBtHmqozkOQCpTQRfWZbeyMOsWdWW"
        redirect_uri = "http://localhost"
        token_url = "https://login.wrike.com/oauth2/token"
        code = self.generate_code()  # Adjusted to call within class
        payload = {
            "client_id": client_id,
            "client_secret": client_secret,
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": redirect_uri,
        }

        try:
            response = requests.post(token_url, data=payload)
            response.raise_for_status()
            self.access_token = response.json().get("access_token")
            if not self.access_token:
                raise Exception("No Access Token Found")
            self.auth_button.pack_forget()  # This will remove the duplicate issue
            self.fetch_people_button.pack(pady=20)  # Now correctly displayed after auth
        except Exception as e:
            messagebox.showerror("Error", f"Failed to obtain access token: {e}")

    def generate_code(self):
        auth_url = "https://login.wrike.com/oauth2/authorize/v4?client_id=Gj7SNTXf&response_type=code&redirect_uri=http://localhost&scope=Default"
        webbrowser.open_new(auth_url)
        code = simpledialog.askstring("Authorization Code", "Enter the code generated by Wrike:", parent=self)
        return code

    def fetch_wrike_people(self):
        if not self.access_token:
            messagebox.showerror("Error", "No access token. Please authenticate first.")
            return
        headers = {'Authorization': f'Bearer {self.access_token}'}
        contacts_url = 'https://www.wrike.com/api/v4/contacts'
        response = requests.get(contacts_url, headers=headers)
        if response.status_code == 200:
            self.fetch_people_button.pack_forget()
            # Clear existing checkboxes
            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()

            contacts = response.json()['data']
            for contact in contacts:
                name = f"{contact.get('firstName', '')} {contact.get('lastName', '')}".strip()
                var = tk.BooleanVar()
                chk = tk.Checkbutton(self.scrollable_frame, text=name, variable=var, bg='#f0f0f0')
                chk.pack(anchor="w")
                self.user_ids[name] = contact['id']
                self.user_checks[name] = var

            # Now, manage the UI components based on the fetched data
            if not self.search_entry.winfo_ismapped():
                self.search_entry.pack(pady=10)  # Display the search entry
                self.checkbox_frame.pack(fill="both", expand=True)
                self.canvas.pack(side="left", fill="both", expand=True)
                self.scrollbar.pack(side="right", fill="y")
                self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
                self.canvas.configure(scrollregion=self.canvas.bbox("all"), yscrollcommand=self.scrollbar.set)
                self.upload_button.pack(pady=20)  # Display the upload button as the next step
        else:
            messagebox.showerror("Error", f"Failed to fetch contacts: {response.status_code}")

    def upload_tableau_file(self):
        self.tableau_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.tableau_file_path:
            messagebox.showinfo("File Selected", f"Selected file: {self.tableau_file_path}")
            self.upload_button.pack_forget()  # Hide upload button
            self.generate_button.pack(pady=10)  # Show generate report button

    
    def process_tableau_file(self):
        if not self.tableau_file_path:
            messagebox.showerror("Error", "Tableau file not uploaded.")
            return None

        try:
            # Load the Excel file
            df = pd.read_excel(self.tableau_file_path)

            # Assuming 'Unnamed: 3' and 'Unnamed: 1' are the columns for first and last names in your Excel file
            # Adjust these as necessary to match your actual file
            df['First Name'] = df['Unnamed: 3'].ffill()
            df['Last Name'] = df['Unnamed: 1'].ffill()

            # Recreate the 'Full Name' column now that we've filled down the names
            df['Full Name'] = df['First Name'].str.strip() + " " + df['Last Name'].str.strip()

            # Recalculate the total hours per person
            total_hours_per_person_filled = df.groupby('Full Name')['Hours'].sum().reset_index()

            # Filter out empty names if any exist after grouping
            total_hours_per_person_filled = total_hours_per_person_filled[total_hours_per_person_filled['Full Name'].str.strip() != ""]

            return total_hours_per_person_filled
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process Tableau file: {e}")
            return None

    def generate_report(self):
        selected_employee_names = [name for name, var in self.user_checks.items() if var.get()]
        selected_employee_ids = {self.user_ids[name]: name for name in selected_employee_names if name in self.user_ids}

        start_date = simpledialog.askstring("Input", "Enter start date (YYYY-MM-DD):", parent=self)
        end_date = simpledialog.askstring("Input", "Enter end date (YYYY-MM-DD):", parent=self)

        # Validate the input dates
        try:
            datetime.strptime(start_date, '%Y-%m-%d')
            datetime.strptime(end_date, '%Y-%m-%d')
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter valid dates in YYYY-MM-DD format.")
            return

        total_hours_by_user = self.fetch_total_hours(selected_employee_ids, start_date, end_date, self.access_token)
        hours_worked_df = self.process_tableau_file()

        if hours_worked_df is not None:
            # Assuming you have methods or logic here to integrate and display or save the combined report data
            self.save_report(total_hours_by_user, hours_worked_df)

    def fetch_total_hours(self, user_ids, start_date, end_date, access_token):
        headers = {'Authorization': f'Bearer {access_token}'}
        total_hours_by_user = {}

        start_date_dt = datetime.strptime(start_date, '%Y-%m-%d').date()
        end_date_dt = datetime.strptime(end_date, '%Y-%m-%d').date()

        for user_id, user_name in user_ids.items():
            timelogs_url = f'https://www.wrike.com/api/v4/contacts/{user_id}/timelogs'
            response = requests.get(timelogs_url, headers=headers)

            if response.status_code == 200:
                timelogs = response.json()['data']
                filtered_timelogs = [log for log in timelogs if start_date_dt <= datetime.strptime(log['trackedDate'], '%Y-%m-%d').date() <= end_date_dt]
                total_hours = sum(log['hours'] for log in filtered_timelogs)
                total_hours_by_user[user_name] = total_hours
            else:
                print(f"Failed to fetch timelogs for {user_name}: {response.status_code}, {response.text}")
                total_hours_by_user[user_name] = 0

        return total_hours_by_user

    def save_report(self, wrike_hours, tableau_hours):
        # First, ensure tableau_hours only contains selected employees.
        # Filter the tableau_hours DataFrame to include only selected names.
        selected_employee_names = [name for name, var in self.user_checks.items() if var.get()]
        tableau_hours_filtered = tableau_hours[tableau_hours['Full Name'].isin(selected_employee_names)]
        
        # Convert 'wrike_hours' dict to DataFrame for merging
        wrike_df = pd.DataFrame(list(wrike_hours.items()), columns=['Full Name', 'Hours Billed'])
        
        # Merge the filtered tableau_hours with wrike_hours on 'Full Name'
        final_report = pd.merge(tableau_hours_filtered, wrike_df, on='Full Name', how='outer').fillna(0)

        # Save the final_report to an Excel file
        try:
            output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if output_path:
                final_report.to_excel(output_path, index=False)
                messagebox.showinfo("Success", f"Report saved to {output_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save report: {e}")

            messagebox.showinfo("Success", "Report Generated!")

        if messagebox.askyesno("Generate Another Report", "Do you want to generate another report?"):
            self.reset_app()
        else:
            self.quit()

    def reset_app(self):
        # Clear the text variable for the search entry
        self.search_var.set("")

        # Destroy all widgets in the scrollable frame to clear the checkboxes
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        # Reset dictionaries
        self.user_ids.clear()
        self.user_checks.clear()

        # Hide elements that should not be visible at the start
        self.fetch_people_button.pack_forget()
        self.upload_button.pack_forget()
        self.generate_button.pack_forget()
        self.search_entry.pack_forget()
        self.checkbox_frame.pack_forget()

        # Reset the state of the application as needed
        # For example, clear any stored data or selections that are specific to the previous report

        # Finally, show the initial button to start the process over
        self.auth_button.pack(pady=20)

if __name__ == "__main__":
    app = TimeAudit()
    app.mainloop()


