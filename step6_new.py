import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import time
import concurrent.futures

# Global dictionary to hold progress widgets per folder
folder_widgets = {}

def select_main_folder():
    folder = filedialog.askdirectory(title="Select Main Videos Folder")
    if folder:
        main_folder_var.set(folder)
        folder_label.config(text=f"Selected folder:\n{folder}")
        # Clear previous progress widgets if any
        for widget in progress_container.winfo_children():
            widget.destroy()
        folder_widgets.clear()

def update_folder_progress(folder_name, progress, remaining):
    # Called from worker threads via root.after to update the widgets
    if folder_name in folder_widgets:
        progress_bar, progress_label = folder_widgets[folder_name]
        progress_bar['value'] = progress
        progress_label.config(text=f"{progress:.1f}%  ETA: {remaining:.1f}s")

def process_single_folder(subfolder_name, subfolder_path, global_output_dir):
    """Processes a single folder and saves the final video in the global output directory."""
    heygen_dir = os.path.join(subfolder_path, "HeyGen Video")
    screen_dir = os.path.join(subfolder_path, "Screen Recording")

    if not os.path.exists(heygen_dir) or not os.path.exists(screen_dir):
        print(f"Skipping folder {subfolder_name}: Missing required subfolders.")
        return

    # Get one video file from each folder (assuming one per folder)
    heygen_files = [f for f in os.listdir(heygen_dir) if os.path.isfile(os.path.join(heygen_dir, f))]
    screen_files = [f for f in os.listdir(screen_dir) if os.path.isfile(os.path.join(screen_dir, f))]
    if not heygen_files or not screen_files:
        print(f"Skipping folder {subfolder_name}: Could not find a video file in one or both subfolders.")
        return

    heygen_video = os.path.join(heygen_dir, heygen_files[0])
    screen_video = os.path.join(screen_dir, screen_files[0])

    # Use the global output directory instead of creating one in each subfolder.
    output_path = os.path.join(global_output_dir, f"{subfolder_name}.mp4")

    # Get the total duration of the screen video via ffprobe.
    ffprobe_cmd = [
        "ffprobe",
        "-v", "error",
        "-show_entries", "format=duration",
        "-of", "default=noprint_wrappers=1:nokey=1",
        screen_video
    ]
    try:
        duration_output = subprocess.check_output(ffprobe_cmd, stderr=subprocess.STDOUT, text=True)
        total_duration = float(duration_output.strip())
    except Exception as e:
        print(f"Error obtaining duration for folder {subfolder_name}: {e}")
        return

    # Build the filter chain:
    # 1. Use scale2ref to scale the overlay video (HeyGen) relative to the background (Screen Recording)
    #    so that its width and height become 14% of the background video's width.
    # 2. Convert the scaled overlay to RGBA and apply a circular mask.
    #    The geq filter uses the overlay's width (W) to compute the center (W/2, H/2) and radius (W/2).
    # 3. Overlay the circular video onto the background with a 20px margin from the top and right.
    filter_complex = (
        "[1:v][0:v]scale2ref=w=iw*0.14:h=iw*0.14[ovrl][base];"
        "[ovrl]format=rgba,"
        "geq=a='if(gt(pow(X-(W/2),2)+pow(Y-(H/2),2),(W/2)*(W/2)),0,255)':"
        "r='r(X,Y)':g='g(X,Y)':b='b(X,Y)'[circ];"
        "[base][circ]overlay=main_w-overlay_w-543:270:shortest=1"
    )

    # Construct FFmpeg command with hardware acceleration and progress reporting.
    ffmpeg_cmd = [
        "ffmpeg",
        "-y",                          # Overwrite output
        "-hwaccel", "videotoolbox",    # Use VideoToolbox for hardware decoding
        "-i", screen_video,            # Background video (Screen Recording)
        "-i", heygen_video,            # Overlay video (HeyGen Video)
        "-filter_complex", filter_complex,
        "-c:v", "h264_videotoolbox",   # Use VideoToolbox encoder for H.264
        "-c:a", "copy",                # Copy audio from background
        "-progress", "pipe:1",         # Send progress info to stdout
        output_path
    ]

    print(f"Processing folder {subfolder_name}...")
    start_time = time.time()
    process = subprocess.Popen(ffmpeg_cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1)

    # Read FFmpeg progress output line-by-line.
    while True:
        line = process.stdout.readline()
        if not line:
            if process.poll() is not None:
                break
            continue
        line = line.strip()
        if line.startswith("out_time_ms="):
            try:
                out_time_ms = int(line.split("=")[1])
                current_time_sec = out_time_ms / 1_000_000.0  # Convert microseconds to seconds.
                progress = (current_time_sec / total_duration) * 100
                elapsed = time.time() - start_time
                estimated_total = elapsed / (progress / 100) if progress > 0 else 0
                remaining = estimated_total - elapsed if estimated_total > elapsed else 0
                root.after(0, update_folder_progress, subfolder_name, progress, remaining)
            except Exception as e:
                print(f"Error parsing progress in folder {subfolder_name}: {e}")
    process.wait()
    root.after(0, update_folder_progress, subfolder_name, 100, 0)
    print(f"Folder {subfolder_name} processed. Output saved to: {output_path}")

def process_all_folders():
    main_folder = main_folder_var.get()
    if not main_folder:
        messagebox.showerror("Error", "Please select a main videos folder first.")
        return

    # Get a sorted list of subfolders based on the number preceding the '-' in the folder name.
    subfolders = []
    for name in os.listdir(main_folder):
        sub_path = os.path.join(main_folder, name)
        if os.path.isdir(sub_path):
            parts = name.split("-")
            if parts and parts[0].strip().isdigit():
                num = int(parts[0].strip())
                subfolders.append((num, name, sub_path))
    subfolders.sort(key=lambda x: x[0])
    if not subfolders:
        messagebox.showerror("Error", "No valid subfolders with a leading number were found.")
        return

    # Determine global output folder: the parent of the main folder plus "Final Video".
    global_output_dir = os.path.join(os.path.dirname(main_folder), "Final Video")
    os.makedirs(global_output_dir, exist_ok=True)

    # Create progress widgets for each folder.
    for widget in progress_container.winfo_children():
        widget.destroy()
    folder_widgets.clear()

    for _, folder_name, folder_path in subfolders:
        frame = tk.Frame(progress_container)
        frame.pack(fill="x", padx=5, pady=2)
        label = tk.Label(frame, text=f"Folder {folder_name}")
        label.pack(side="left")
        pb = ttk.Progressbar(frame, orient='horizontal', mode='determinate', maximum=100, length=200)
        pb.pack(side="left", padx=5)
        status = tk.Label(frame, text="0%  ETA: --s", width=20)
        status.pack(side="left")
        folder_widgets[folder_name] = (pb, status)

    # Process folders concurrently.
    max_workers = 3  # Adjust as needed.
    futures = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        for _, folder_name, folder_path in subfolders:
            future = executor.submit(process_single_folder, folder_name, folder_path, global_output_dir)
            futures.append(future)
        concurrent.futures.wait(futures)
    messagebox.showinfo("Done", "All folders have been processed.")
    process_button.config(state=tk.NORMAL)

def start_processing():
    process_button.config(state=tk.DISABLED)
    threading.Thread(target=process_all_folders, daemon=True).start()

# Set up the Tkinter GUI.
root = tk.Tk()
root.title("Parallel Video Overlay Processor")

main_folder_var = tk.StringVar()

select_button = tk.Button(root, text="Select Main Videos Folder", command=select_main_folder, width=40)
select_button.pack(padx=10, pady=10)

folder_label = tk.Label(root, text="No folder selected", wraplength=400, justify="center")
folder_label.pack(padx=10, pady=5)

process_button = tk.Button(root, text="Process Videos", command=start_processing, width=40)
process_button.pack(padx=10, pady=10)

# A container frame to hold progress bars for each folder.
progress_container = tk.Frame(root)
progress_container.pack(padx=10, pady=10, fill="x")

root.mainloop()