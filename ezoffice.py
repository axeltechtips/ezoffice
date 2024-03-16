import os
import subprocess
from tqdm import tqdm
import requests

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
            print("\nFile downloaded successfully!")
            print("File saved as:", save_path)
            print("File size:", round(total_size / 1024, 2), "KB")
            return save_path
        else:
            print("Failed to download file. Status code:", response.status_code)
            return None
    except Exception as e:
        print("An error occurred:", e)
        return None

def run_exe(exe_path):
    try:
        print("Running the downloaded executable...")
        subprocess.run([exe_path])
    except Exception as e:
        print("An error occurred while running the executable:", e)

def main():
    print("Welcome to the Office downloader and installer!")
    print("Choose the Office version you want to download:")
    print("1. Office 365")
    print("2. Office 2021")
    print("3. Office 2019")
    print("4. Office 2016")
    
    choice = input("Enter your choice (1, 2, 3, or 4): ")

    urls = {
        "1": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365ProPlusRetail&platform=x64&language=en-us&version=O16GA",
        "2": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2021Retail&platform=x64&language=en-us&version=O16GA",
        "3": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2019Retail&platform=x64&language=en-us&version=O16GA",
        "4": "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x64&language=en-us&version=O16GA"
    }

    if choice not in urls:
        print("Invalid choice.")
        return

    url = urls[choice]
    save_path = input("Enter the directory where you want to save the file: ")

    downloaded_file_path = download_file(url, save_path)
    if downloaded_file_path:
        run_exe(downloaded_file_path)

if __name__ == "__main__":
    main()
