import os
import json
import threading
import requests
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import io
from PIL import Image  # Pillow is required
import csv
from openai import OpenAI

# ==========================================
# SET YOUR API KEYS HERE
# ==========================================
# Replace with your own OpenAI API key.
client = OpenAI(api_key="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")

# The OCR.space API key is hard-coded as in your sample code.
OCR_API_KEY = 'K85898372388957'  # Replace with your own if needed

# ==========================================
# Helper Function to Downscale Images
# ==========================================
def downscale_image_to_threshold(image, threshold):
    """
    Given a PIL Image object, returns an in-memory BytesIO
    containing the image encoded as JPEG that is under the given threshold (in bytes).
    The function first tries reducing the JPEG quality, and if necessary, then reduces image dimensions.
    """
    if image.mode not in ("RGB", "L"):
        image = image.convert("RGB")

    quality = 95
    img = image.copy()
    
    while True:
        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=quality)
        size = buffer.tell()
        if size <= threshold:
            return buffer
        # Reduce quality if possible
        if quality > 20:
            quality -= 5
        else:
            # If quality is low, then reduce dimensions by 10%
            new_width = int(img.width * 0.9)
            new_height = int(img.height * 0.9)
            if new_width < 1 or new_height < 1:
                return buffer
            img = img.resize((new_width, new_height), Image.ANTIALIAS)
            quality = 95  # reset quality for the resized image

# ==========================================
# OCR.space API Function
# ==========================================
def ocr_space_file(filename, overlay=False, api_key=OCR_API_KEY, language='eng'):
    """
    OCR.space API request with a local image file.
    If the file is larger than 1024 KB, it is downscaled first.
    :param filename: The full path to the image file.
    :param overlay: Whether OCR.space overlay is required.
    :param api_key: Your OCR.space API key.
    :param language: Language code for OCR.
    :return: The API response as a JSON-formatted string.
    """
    payload = {
        'isOverlayRequired': overlay,
        'apikey': api_key,
        'language': language,
    }
    
    threshold = 1024 * 1024  # 1024 KB in bytes
    file_size = os.path.getsize(filename)
    
    if file_size > threshold:
        try:
            image = Image.open(filename)
        except Exception as e:
            return json.dumps({"error": f"Unable to open image for downscaling: {e}"})
        buffer = downscale_image_to_threshold(image, threshold)
        buffer.seek(0)
        files = {os.path.basename(filename): buffer}
        response = requests.post(
            'https://api.ocr.space/parse/image',
            files=files,
            data=payload,
        )
    else:
        with open(filename, 'rb') as f:
            response = requests.post(
                'https://api.ocr.space/parse/image',
                files={os.path.basename(filename): f},
                data=payload,
            )
    return response.content.decode()

# ==========================================
# OpenAI ChatGPT API Call Function
# ==========================================
def call_chatgpt_api(prompt_text):
    """
    Call the OpenAI ChatGPT API with the given prompt_text.
    The prompt is prefixed with a detailed instruction message followed by the OCR output.
    :param prompt_text: The OCR text to send to the ChatGPT API.
    :return: The response text from the API.
    """
    full_prompt = (
        "The below text is the OCR output of multiple screenshots from someone's LinkedIn. "
        "Please analyze this text carefully, and once you understand the OCR output, create a 30-second script that will be spoken in a video format from me to the person from the LinkedIn profile. "
        "The goal of the script is to network and connect with this person leveraging the information on their LinkedIn profile, with the end goal being to get them to hop on a call with me. "
        "The script should:\n"
        "- Briefly introduce myself, my name is Ryan.\n"
        "- Mention that I came across their LinkedIn profile and was impressed by their background.\n"
        "- Be concise, natural, and sound like they are spoken, not written.\n"
        "- Use a casual yet professional tone that builds trust and connection.\n"
        "- Highlight specific achievements, experiences, or relatable details (e.g., shared locations or schools) ONLY if it connects to them. "
        "Information about me that you could relate to if it happens to connect is that I am a sophomore at Babson College in Boston, I am an Entrepreneur and have been running my own software business since high school, "
        "I grew up internationally, living in Venezuela, Poland, and Japan where I finished high school.\n"
        "- Avoid overly formal or robotic language and feel authentic, conversational, and engaging.\n"
        "- End with a genuine, friendly invitation like, “I honestly think you'd be a great connection to have, so let me know if you want to hop on a call at some point!\"\n"
        "- Search the internet to find more information about this person, and if any of it is new information (not on the LinkedIn profile) and relevant, include it in your script.\n"
        "- Output only the script itself, nothing else.\n"
        " Now, without saying anything else in your response, output your script. DO NOT OUTPUT ANYTHING OTHER THAN THE TEXT OF THE SCRIPT ITSELF. The OCR output is below:\n\n"
        + prompt_text
    )
    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",  # or use "gpt-4" if available and desired
            messages=[{"role": "user", "content": full_prompt}],
            temperature=0.7
        )
        answer = response.choices[0].message.content.strip()
        return answer
    except Exception as e:
        return f"Error calling ChatGPT API: {str(e)}"

# ==========================================
# Main Application Class
# ==========================================
class LinkedInOCRApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("LinkedIn OCR & ChatGPT Application")
        self.geometry("700x350")
        self.resizable(False, False)

        self.main_folder = None   # Folder with subfolders for each profile's screenshots
        self.main_csv = None      # Main CSV file used in previous steps

        # Title label
        self.title_label = tk.Label(self, text="LinkedIn OCR & ChatGPT Application", font=("Helvetica", 16))
        self.title_label.pack(pady=10)

        # Folder selection frame
        folder_frame = tk.Frame(self)
        folder_frame.pack(pady=5)
        self.folder_label = tk.Label(folder_frame, text="No main folder selected", fg="blue")
        self.folder_label.grid(row=0, column=0, padx=5)
        self.select_folder_btn = tk.Button(folder_frame, text="Select Main Folder", width=20, command=self.select_folder)
        self.select_folder_btn.grid(row=0, column=1, padx=5)

        # Main CSV selection frame
        csv_frame = tk.Frame(self)
        csv_frame.pack(pady=5)
        self.csv_label = tk.Label(csv_frame, text="No main CSV selected", fg="blue")
        self.csv_label.grid(row=0, column=0, padx=5)
        self.select_csv_btn = tk.Button(csv_frame, text="Select Main CSV", width=20, command=self.select_csv)
        self.select_csv_btn.grid(row=0, column=1, padx=5)

        # Start Processing button
        self.start_btn = tk.Button(self, text="Start Processing", width=25, command=self.start_processing, state=tk.DISABLED)
        self.start_btn.pack(pady=10)

        # Progress bar
        self.progress = ttk.Progressbar(self, orient="horizontal", mode="determinate", length=600)
        self.progress.pack(pady=10)

        # Status label
        self.status_label = tk.Label(self, text="", fg="green")
        self.status_label.pack(pady=5)

    def select_folder(self):
        """Select the main folder containing profile subfolders."""
        folder = filedialog.askdirectory()
        if folder:
            self.main_folder = folder
            self.folder_label.config(text=f"Main Folder: {folder}")
            self.update_start_button_state()
        else:
            self.folder_label.config(text="No main folder selected")
            self.main_folder = None
            self.update_start_button_state()

    def select_csv(self):
        """Select the main CSV file used in earlier steps."""
        file_path = filedialog.askopenfilename(title="Select Main CSV", filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if file_path:
            self.main_csv = file_path
            self.csv_label.config(text=f"Main CSV: {file_path}")
            self.update_start_button_state()
        else:
            self.csv_label.config(text="No main CSV selected")
            self.main_csv = None
            self.update_start_button_state()

    def update_start_button_state(self):
        """Enable the Start Processing button only if both main folder and main CSV are selected."""
        if self.main_folder and self.main_csv:
            self.start_btn.config(state=tk.NORMAL)
        else:
            self.start_btn.config(state=tk.DISABLED)

    def start_processing(self):
        """Start processing the subfolders in a separate thread."""
        if not (self.main_folder and self.main_csv):
            messagebox.showerror("Error", "Please select both a main folder and a main CSV file.")
            return

        # Disable buttons while processing
        self.start_btn.config(state=tk.DISABLED)
        self.select_folder_btn.config(state=tk.DISABLED)
        self.select_csv_btn.config(state=tk.DISABLED)
        self.status_label.config(text="Processing profiles, please wait...", fg="black")

        # Start processing in a new thread so the GUI remains responsive.
        thread = threading.Thread(target=self.process_subfolders)
        thread.start()

    def process_subfolders(self):
        """
        For each subfolder (profile), run OCR on all images, call ChatGPT with combined OCR text,
        and update the main CSV file by adding/updating columns 6 and 7:
          - Column 6: Profile Folder name (e.g. "2 - Joe Smith")
          - Column 7: ChatGPT (4o-mini) script output.
        The main CSV is re-written after processing each profile.
        """
        # Read the main CSV file.
        try:
            with open(self.main_csv, mode='r', encoding='utf-8-sig') as csvfile:
                csv_reader = csv.reader(csvfile)
                csv_data = list(csv_reader)
        except Exception as e:
            self.safe_update(lambda: self.status_label.config(text=f"Error reading main CSV: {str(e)}", fg="red"))
            self.safe_update(lambda: self.enable_buttons())
            return

        if not csv_data:
            self.safe_update(lambda: self.status_label.config(text="Main CSV is empty.", fg="red"))
            self.safe_update(lambda: self.enable_buttons())
            return

        # Ensure the header row has at least 7 columns.
        header = csv_data[0]
        if len(header) < 7:
            header += ["Profile Folder", "OpenAI Response"]
            csv_data[0] = header

        # Get list of subfolders in the main folder.
        subfolder_names = [name for name in os.listdir(self.main_folder) 
                           if os.path.isdir(os.path.join(self.main_folder, name))]
        if not subfolder_names:
            self.safe_update(lambda: self.status_label.config(text="No subfolders found in the main folder.", fg="red"))
            self.safe_update(lambda: self.enable_buttons())
            return

        total_subfolders = len(subfolder_names)

        # Process each subfolder (sorted by name).
        for index, subfolder in enumerate(sorted(subfolder_names), start=1):
            subfolder_path = os.path.join(self.main_folder, subfolder)
            valid_extensions = ('.png', '.jpg', '.jpeg', '.bmp', '.tiff')
            image_files = [os.path.join(subfolder_path, f) for f in os.listdir(subfolder_path)
                           if f.lower().endswith(valid_extensions)]
            
            if not image_files:
                openai_response = "No images found in subfolder"
                self.safe_update(lambda: self.status_label.config(text=f"{subfolder}: No images found.", fg="red"))
            else:
                ocr_combined_text = ""
                # Process each image file in the subfolder.
                for i, image_path in enumerate(image_files, start=1):
                    try:
                        ocr_response = ocr_space_file(filename=image_path, api_key=OCR_API_KEY, language='eng')
                        try:
                            parsed = json.loads(ocr_response)
                            pretty_response = json.dumps(parsed, indent=4)
                        except json.JSONDecodeError:
                            pretty_response = ocr_response
                    except Exception as e:
                        pretty_response = f"Error processing image: {str(e)}"
                    
                    ocr_combined_text += f"OCR Response from Screenshot {i}:\n\n{pretty_response}\n\n"

                # Optionally, write the combined OCR text to a temporary file (not required)
                temp_txt_path = os.path.join(subfolder_path, "ocr_combined.txt")
                try:
                    with open(temp_txt_path, "w", encoding="utf-8") as txt_file:
                        txt_file.write(ocr_combined_text)
                except Exception as e:
                    self.safe_update(lambda: self.status_label.config(text=f"Error writing OCR text for {subfolder}: {str(e)}", fg="red"))
                    openai_response = f"Error writing OCR text: {str(e)}"
                else:
                    # Call ChatGPT with the combined OCR text.
                    openai_response = call_chatgpt_api(ocr_combined_text)

            # Parse subfolder name to extract row number.
            # Expected format: "2 - Joe Smith" where "2" is the row number.
            try:
                row_num_str = subfolder.split(" - ")[0].strip()
                row_number = int(row_num_str)
                # CSV data row for row number N is at index N-1 (assuming header is at index 0).
                csv_row_index = row_number - 1
            except Exception as e:
                self.safe_update(lambda: self.status_label.config(text=f"Could not parse row number from folder '{subfolder}': {e}", fg="red"))
                continue

            if csv_row_index < 1 or csv_row_index >= len(csv_data):
                self.safe_update(lambda: self.status_label.config(text=f"Row number {row_number} from folder '{subfolder}' is out of range in main CSV.", fg="red"))
            else:
                row = csv_data[csv_row_index]
                if len(row) < 7:
                    row.extend([""] * (7 - len(row)))
                row[5] = subfolder          # Column 6: Profile Folder name
                row[6] = openai_response      # Column 7: ChatGPT output
                csv_data[csv_row_index] = row

            # Update progress bar and status.
            progress_percent = int((index / total_subfolders) * 100)
            self.safe_update(lambda val=progress_percent: self.progress.config(value=val))
            self.safe_update(lambda: self.status_label.config(text=f"Processed profile {index} of {total_subfolders}: {subfolder}", fg="black"))

            # Write updated CSV data back to the main CSV file (live update).
            try:
                with open(self.main_csv, mode='w', newline='', encoding="utf-8") as csvfile:
                    csv_writer = csv.writer(csvfile)
                    csv_writer.writerows(csv_data)
            except Exception as e:
                self.safe_update(lambda: self.status_label.config(text=f"Error writing updated main CSV: {str(e)}", fg="red"))

        self.safe_update(lambda: self.status_label.config(text=f"Processing complete. Main CSV updated at:\n{self.main_csv}", fg="green"))
        self.safe_update(lambda: self.enable_buttons())
        self.safe_update(lambda: self.progress.config(value=0))

    def enable_buttons(self):
        """Re-enable selection and start buttons."""
        self.start_btn.config(state=tk.NORMAL)
        self.select_folder_btn.config(state=tk.NORMAL)
        self.select_csv_btn.config(state=tk.NORMAL)

    def safe_update(self, func):
        """Safely schedule GUI updates from the worker thread."""
        self.after(0, func)

# ==========================================
# Main Program Entry Point
# ==========================================
if __name__ == "__main__":
    app = LinkedInOCRApp()
    app.mainloop()