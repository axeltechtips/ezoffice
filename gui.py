import os
import subprocess
import requests
import tkinter as tk
from tkinter import ttk, filedialog

class OfficeDownloaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("EzOffice - Office Downloader")
        self.root.geometry("540x400")
        self.root.configure(padx=20, pady=20, bg="#2E2E2E")  # Dark background color

        # Set up modern dark theme styles
        self.setup_styles()

        # URLs for different Office versions
        self.urls = {
            "Office 365": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365ProPlusRetail&platform=x64&language=en-us&version=O16GA",
            "Office 2021": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2021Retail&platform=x64&language=en-us&version=O16GA",
            "Office 2019": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2019Retail&platform=x64&language=en-us&version=O16GA",
            "Office 2016": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x64&language=en-us&version=O16GA"
        }

        # Create widgets
        self.create_widgets()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        # Configure styles for dark mode
        style.configure("TFrame", background="#2E2E2E")
        style.configure("TLabel", background="#2E2E2E", foreground="#FFFFFF", font=("Helvetica Neue", 12))
        style.configure("TCombobox", fieldbackground="#3E3E3E", background="#2E2E2E", foreground="#FFFFFF", font=("Helvetica Neue", 12))
        style.configure("TEntry", fieldbackground="#3E3E3E", background="#2E2E2E", foreground="#FFFFFF", font=("Helvetica Neue", 12))
        style.configure("TButton", background="#007ACC", foreground="#FFFFFF", font=("Helvetica Neue", 12), padding=5)
        style.map("TButton", background=[("active", "#005B99")])

    def create_widgets(self):
        # Title label
        title_label = ttk.Label(self.root, text="EzOffice", font=("Helvetica Neue", 20, "bold"))
        title_label.pack(pady=(10, 20))

        # Office version dropdown
        version_label = ttk.Label(self.root, text="Select Office Version")
        version_label.pack(anchor="w", padx=5, pady=(10, 5))
        
        self.version_combobox = ttk.Combobox(self.root, values=list(self.urls.keys()), state="readonly")
        self.version_combobox.pack(fill='x', pady=(0, 20), ipadx=5)

        # Directory selection frame
        directory_frame = ttk.Frame(self.root)
        directory_frame.pack(fill='x', pady=10)

        directory_label = ttk.Label(directory_frame, text="Save Directory")
        directory_label.pack(side=tk.LEFT, padx=(0, 10))

        self.directory_entry = ttk.Entry(directory_frame)
        self.directory_entry.pack(side=tk.LEFT, fill='x', expand=True)

        browse_button = ttk.Button(directory_frame, text="Browse", command=self.select_directory)
        browse_button.pack(side=tk.LEFT, padx=(10, 0))

        # Download button
        download_button = ttk.Button(self.root, text="Download and Install", command=self.download_and_run)
        download_button.pack(pady=(20, 10))

        # Result label
        self.result_label = ttk.Label(self.root, text="", wraplength=480, font=("Helvetica Neue", 10))
        self.result_label.pack(pady=(10, 0))

    def select_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.directory_entry.delete(0, tk.END)
            self.directory_entry.insert(0, directory)

    def download_file(self, url, save_path):
        try:
            response = requests.get(url, stream=True)
            response.raise_for_status()
            file_name = "OfficeSetup.exe"
            full_save_path = os.path.join(save_path, file_name)

            total_size = int(response.headers.get('content-length', 0))
            with open(full_save_path, 'wb') as f:
                for data in response.iter_content(chunk_size=1024):
                    f.write(data)

            self.result_label.config(text=f"\nFile downloaded successfully at: {full_save_path}")
            return full_save_path
        except requests.HTTPError as http_err:
            self.result_label.config(text=f"HTTP error occurred: {http_err}")
        except Exception as e:
            self.result_label.config(text=f"An error occurred: {e}")
        return None

    def run_exe(self, exe_path):
        try:
            self.result_label.config(text="Running the downloaded executable...")
            subprocess.run([exe_path], check=True)
        except subprocess.CalledProcessError as err:
            self.result_label.config(text=f"An error occurred: {err}")
        except Exception as e:
            self.result_label.config(text=f"An error occurred: {e}")

    def download_and_run(self):
        choice = self.version_combobox.get()
        if choice not in self.urls:
            self.result_label.config(text="Please select a valid Office version.")
            return

        url = self.urls[choice]
        save_path = self.directory_entry.get()

        downloaded_file_path = self.download_file(url, save_path)
        if downloaded_file_path:
            self.run_exe(downloaded_file_path)

if __name__ == "__main__":
    root = tk.Tk()
    app = OfficeDownloaderApp(root)
    root.mainloop()
