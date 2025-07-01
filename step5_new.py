import tkinter as tk
from tkinter import filedialog, messagebox
import csv
import os
import requests
import time

def select_csv_file():
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if file_path:
        csv_file_path.set(file_path)
        csv_label.config(text=file_path)

def select_main_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        main_folder_path.set(folder_path)
        folder_label.config(text=folder_path)

def download_videos():
    csv_path = csv_file_path.get()
    main_folder = main_folder_path.get()
    api_key = api_key_var.get().strip()

    if not csv_path or not main_folder:
        messagebox.showerror("Error", "Please select both the CSV file and the main folder.")
        return

    # Use the x-api-key header instead of an Authorization Bearer token.
    headers = {"x-api-key": api_key} if api_key else {}

    try:
        with open(csv_path, newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile)
            all_rows = list(reader)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read CSV file: {e}")
        return

    if len(all_rows) < 2:
        messagebox.showerror("Error", "CSV file does not contain data rows.")
        return

    total_rows = len(all_rows) - 1  # Excluding header.
    progress_text.delete("1.0", tk.END)
    progress_text.insert(tk.END, f"Starting download for {total_rows} videos...\n\n")

    # Process each video (starting from row 2, because row 1 is header).
    for row_number, row in enumerate(all_rows[1:], start=2):
        if len(row) < 8:
            progress_text.insert(tk.END, f"Row {row_number}: Not enough columns. Skipping.\n")
            continue

        video_id = row[7].strip()  # 8th column (index 7)
        if not video_id:
            progress_text.insert(tk.END, f"Row {row_number}: No video ID found. Skipping.\n")
            continue

        # --- Step 1: Query the Video Status Endpoint ---
        video_status_url = f"https://api.heygen.com/v1/video_status.get?video_id={video_id}"
        progress_text.insert(tk.END, f"Row {row_number}: Querying video status at:\n{video_status_url}\n")
        progress_text.update_idletasks()

        max_attempts = 12  # wait up to 60 seconds (12*5 seconds)
        attempt = 0
        video_url = None

        while attempt < max_attempts:
            try:
                status_response = requests.get(video_status_url, headers=headers)
                status_response.raise_for_status()
                status_json = status_response.json()
            except Exception as e:
                progress_text.insert(tk.END, f"Row {row_number}: Error fetching video status: {e}\n")
                break

            # Expected JSON structure:
            # { "data": { "status": "...", "video_url": "...", "thumbnail_url": "..." } }
            data = status_json.get("data", {})
            video_status = data.get("status")
            if video_status == "completed":
                video_url = data.get("video_url")
                progress_text.insert(tk.END, f"Row {row_number}: Video completed!\nVideo URL: {video_url}\n")
                break
            elif video_status in ["processing", "pending"]:
                progress_text.insert(tk.END, f"Row {row_number}: Video is still {video_status}. Checking again in 5 seconds...\n")
                progress_text.update_idletasks()
                time.sleep(5)
                attempt += 1
                continue
            elif video_status == "failed":
                error_message = data.get("error", "Unknown error")
                progress_text.insert(tk.END, f"Row {row_number}: Video generation failed: {error_message}\n")
                break
            else:
                progress_text.insert(tk.END, f"Row {row_number}: Unexpected video status: {video_status}\n")
                break

        if not video_url:
            progress_text.insert(tk.END, f"Row {row_number}: Video URL not retrieved, skipping.\n\n")
            continue

        # --- Step 2: Download the Video ---
        progress_text.insert(tk.END, f"Row {row_number}: Downloading video from:\n{video_url}\n")
        progress_text.update_idletasks()
        try:
            video_response = requests.get(video_url)
            video_response.raise_for_status()
        except Exception as e:
            progress_text.insert(tk.END, f"Row {row_number}: Error downloading video: {e}\n")
            continue

        # --- Step 3: Save the Video in the Correct Subfolder ---
        # Find a subfolder in the main folder that starts with the row number.
        folder_candidates = [
            f for f in os.listdir(main_folder)
            if os.path.isdir(os.path.join(main_folder, f)) and f.startswith(str(row_number))
        ]
        if not folder_candidates:
            progress_text.insert(tk.END, f"Row {row_number}: No subfolder starting with '{row_number}' found. Skipping.\n")
            continue
        subfolder_path = os.path.join(main_folder, folder_candidates[0])

        # Create the "HeyGen Video" folder inside the subfolder.
        heyg_folder = os.path.join(subfolder_path, "HeyGen Video")
        os.makedirs(heyg_folder, exist_ok=True)

        video_file_path = os.path.join(heyg_folder, "video.mp4")
        try:
            with open(video_file_path, "wb") as video_file:
                video_file.write(video_response.content)
            progress_text.insert(tk.END, f"Row {row_number}: Video saved to:\n{video_file_path}\n\n")
        except Exception as e:
            progress_text.insert(tk.END, f"Row {row_number}: Error saving video: {e}\n\n")
            continue

    progress_text.insert(tk.END, "Download process completed.\n")
    messagebox.showinfo("Done", "Download process completed.")

# Set up the Tkinter GUI.
root = tk.Tk()
root.title("HeyGen Video Downloader")

csv_file_path = tk.StringVar()
main_folder_path = tk.StringVar()
# Pre-populate the API key field with your API key.
api_key_var = tk.StringVar(value="OTc5ZmVmZGNjYzUxNDljNDlmYzUxNGJhYjRjOTNmNDItMTcyNzI4NzE0MA==")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack()

# CSV File Selection
csv_button = tk.Button(frame, text="Select CSV File", command=select_csv_file)
csv_button.grid(row=0, column=0, sticky="w")
csv_label = tk.Label(frame, text="No CSV file selected", wraplength=400)
csv_label.grid(row=0, column=1, padx=10)

# Main Folder Selection
folder_button = tk.Button(frame, text="Select Main Recording Folder", command=select_main_folder)
folder_button.grid(row=1, column=0, sticky="w")
folder_label = tk.Label(frame, text="No folder selected", wraplength=400)
folder_label.grid(row=1, column=1, padx=10)

# API Key Entry
api_key_label = tk.Label(frame, text="API Key:")
api_key_label.grid(row=2, column=0, sticky="w")
api_key_entry = tk.Entry(frame, textvariable=api_key_var, width=50)
api_key_entry.grid(row=2, column=1, padx=10)
api_key_info = tk.Label(frame, text="(Keep your API key secure!)", fg="gray")
api_key_info.grid(row=3, column=1, sticky="w", padx=10)

# Download Button
download_button = tk.Button(frame, text="Download Videos", command=download_videos)
download_button.grid(row=4, column=0, columnspan=2, pady=10)

# Progress Text Widget
progress_text = tk.Text(frame, width=70, height=18)
progress_text.grid(row=5, column=0, columnspan=2, pady=10)

root.mainloop()