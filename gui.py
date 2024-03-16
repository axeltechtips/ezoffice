import os
import subprocess
from tqdm import tqdm
import requests
import tkinter as tk
from tkinter import ttk, filedialog

def download_file(url, save_path):
    try:
        response = requests.get(url, stream=True)
        if response.status_code == 200:
            if os.path.isdir(save_path):
                file_name = "OfficeSetup.exe"
                save_path = os.path.join(save_path, file_name)
            total_size = int(response.headers.get('content-length', 0))
            progress = tqdm(total=total_size, unit='B', unit_scale=True)
            with open(save_path, 'wb') as f:
                for data in response.iter_content(chunk_size=1024):
                    f.write(data)
                    progress.update(len(data))
            progress.close()
            result_label.config(text="\nFile downloaded successfully!\nFile saved as: {}\nFile size: {} KB".format(save_path, round(total_size / 1024, 2)))
            return save_path
        else:
            result_label.config(text="Failed to download file. Status code: {}".format(response.status_code))
            return None
    except Exception as e:
        result_label.config(text="An error occurred: {}".format(e))
        return None

def run_exe(exe_path):
    try:
        result_label.config(text="Running the downloaded executable...")
        subprocess.run([exe_path])
    except Exception as e:
        result_label.config(text="An error occurred while running the executable: {}".format(e))

def select_directory():
    directory = filedialog.askdirectory()
    if directory:
        directory_entry.delete(0, tk.END)
        directory_entry.insert(0, directory)

def download_and_run():
    choice = version_combobox.get()
    urls = {
        "Office 365": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365ProPlusRetail&platform=x64&language=en-us&version=O16GA",
        "Office 2021": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2021Retail&platform=x64&language=en-us&version=O16GA",
        "Office 2019": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2019Retail&platform=x64&language=en-us&version=O16GA",
        "Office 2016": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x64&language=en-us&version=O16GA"
    }

    if choice not in urls:
        result_label.config(text="Invalid choice.")
        return

    url = urls[choice]
    save_path = directory_entry.get()

    downloaded_file_path = download_file(url, save_path)
    if downloaded_file_path:
        run_exe(downloaded_file_path)

# GUI setup
root = tk.Tk()
root.title("EzOffice")
root.configure(bg="#1e1e1e")
root.geometry("500x300")

title_label = tk.Label(root, text="Welcome to the Office downloader and installer!", fg="white", bg="#1e1e1e", font=("Arial", 16))
title_label.pack(pady=10)

version_label = tk.Label(root, text="Choose the Office version you want to download:", fg="white", bg="#1e1e1e", font=("Arial", 12))
version_label.pack()

version_combobox = ttk.Combobox(root, values=["Office 365", "Office 2021", "Office 2019", "Office 2016"], font=("Arial", 12), state="readonly")
version_combobox.pack(pady=5)

directory_frame = tk.Frame(root, bg="#1e1e1e")
directory_frame.pack(pady=5)

directory_label = tk.Label(directory_frame, text="Save Directory:", fg="white", bg="#1e1e1e", font=("Arial", 12))
directory_label.pack(side=tk.LEFT)

directory_entry = tk.Entry(directory_frame, width=40, font=("Arial", 12))
directory_entry.pack(side=tk.LEFT)

browse_button = tk.Button(directory_frame, text="Browse", command=select_directory, bg="#363636", fg="white", font=("Arial", 10))
browse_button.pack(side=tk.LEFT, padx=(5,0))

download_button = tk.Button(root, text="Download and Install", command=download_and_run, bg="#363636", fg="white", font=("Arial", 12))
download_button.pack(pady=10)

result_label = tk.Label(root, text="", fg="white", bg="#1e1e1e", font=("Arial", 12), justify="left")
result_label.pack()

root.mainloop()
