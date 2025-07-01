import tkinter as tk
from tkinter import filedialog, messagebox
import csv
import http.client
import json
import os
import ssl
import certifi

# ---------------------------
# Configuration for serper.dev
# ---------------------------
SERPER_API_KEY = '2110612bb6edb33596de66cace6357506f26f266'  # <-- Replace with your actual API key
SERPER_HOST = "google.serper.dev"
SERPER_ENDPOINT = "/search"

class LinkedInScraperApp:
    def __init__(self, master):
        self.master = master
        self.master.title("LinkedIn Scraper with Certifi")

        # GUI Elements
        self.label_file = tk.Label(self.master, text="No CSV file selected.")
        self.button_browse = tk.Button(self.master, text="Select CSV", command=self.browse_csv)
        self.button_start = tk.Button(self.master, text="Start Search", command=self.start_search)
        self.button_exit = tk.Button(self.master, text="Exit", command=self.master.quit)

        # Layout
        self.label_file.pack(pady=10)
        self.button_browse.pack(pady=5)
        self.button_start.pack(pady=5)
        self.button_exit.pack(pady=5)

        self.csv_file_path = None

    def browse_csv(self):
        """Opens a file dialog for the user to select the CSV file."""
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            self.csv_file_path = file_path
            self.label_file.config(text=f"Selected: {os.path.basename(file_path)}")

    def start_search(self):
        """
        Reads the original CSV, performs a LinkedIn search for each row (based on first name, last name, and company),
        and appends the LinkedIn header (title) and URL as new columns (columns 4 and 5) in the original CSV.
        """
        if not self.csv_file_path:
            messagebox.showerror("Error", "Please select a CSV file first.")
            return

        try:
            # Read the original CSV into memory
            with open(self.csv_file_path, mode='r', encoding='utf-8-sig') as infile:
                reader = csv.reader(infile)
                rows = list(reader)

            if not rows:
                messagebox.showerror("Error", "CSV file is empty.")
                return

            # Assume the first row is a header row; append new headers
            header = rows[0] + ["Title", "URL"]
            new_rows = [header]

            # Process each data row (starting from the second row)
            for row in rows[1:]:
                # Ensure the row has at least 3 columns; if not, just add two empty columns
                if len(row) < 3:
                    row.extend(["", ""])
                    new_rows.append(row)
                    continue

                first_name, last_name, current_company = row[0], row[1], row[2]
                query = f"site:linkedin.com {first_name} {last_name} {current_company}"

                # Use the serper.dev API to get search results
                top_title, top_url = self.search_linkedin(query)

                if top_title and top_url:
                    row.extend([top_title, top_url])
                else:
                    row.extend(["No result found", ""])

                new_rows.append(row)

            # Overwrite the original CSV file with the updated data
            with open(self.csv_file_path, mode='w', newline='', encoding='utf-8') as outfile:
                writer = csv.writer(outfile)
                writer.writerows(new_rows)

            messagebox.showinfo("Success", f"Results added to {os.path.basename(self.csv_file_path)}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

    def search_linkedin(self, query):
        """
        Uses the serper.dev API to search for the query string.
        Returns the title and URL of the top result, or (None, None) if not found.
        """
        # Create a secure SSL context that uses certifi's CA bundle
        ssl_context = ssl.create_default_context(cafile=certifi.where())

        try:
            conn = http.client.HTTPSConnection(SERPER_HOST, context=ssl_context)
            payload = json.dumps({
                "q": query,
                "autocorrect": False
            })
            headers = {
                'X-API-KEY': SERPER_API_KEY,
                'Content-Type': 'application/json'
            }
            conn.request("POST", SERPER_ENDPOINT, payload, headers)
            res = conn.getresponse()
            data = res.read()
            conn.close()

            response_json = json.loads(data.decode("utf-8"))

            # Check for results
            if 'organic' in response_json and len(response_json['organic']) > 0:
                top_result = response_json['organic'][0]
                return top_result.get('title'), top_result.get('link')
            else:
                return None, None

        except Exception as e:
            print("Error searching with serper.dev:", e)
            return None, None

if __name__ == "__main__":
    root = tk.Tk()
    app = LinkedInScraperApp(root)
    root.mainloop()