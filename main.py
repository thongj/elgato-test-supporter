from tkinter import * 
from tkinter.ttk import *
from tkinter import messagebox
import shutil
import os
import psutil
import subprocess
import webbrowser

apps = {
    'WaveLink':'Elgato Wave Link',
    'CameraHub':'Elgato Camera Hub',
    'ControlCenter':'Elgato Control Center', 
    '4KCaptureUtility':'Elgato 4K Capture Utility',
    'StreamDeck':'Elgato Stream Deck',
    'Elgato.Studio':'Elgato.Studio_g54w8ztgkx496',
    'Elgato.WaveCast':'Elgato.WaveCast_g54w8ztgkx496'
    } 

master = Tk()
master.title("Elgato test supporter")
#master.geometry("500x200")

def open_app(app_name):
    if app_name == "CameraHub":
        path=f"C:\\Program Files\\Elgato\\{app_name}\\Camera Hub.exe"
    else:
        path=f"C:\\Program Files\\Elgato\\{app_name}\\{app_name}.exe"
    
    if '.' in app_name:
        subprocess.Popen(["explorer.exe", f"shell:AppsFolder\\{apps[app_name]}!App"])
    else:
        try:
            os.startfile(path)  # Open the application
            print(f"Opened application: {path}")
        except FileNotFoundError:
            print(f"Application not found: {path}")
        except Exception as e:
            print(f"Error opening application {path}: {e}")

def open_elgato_appdata_folder():
    appdata = os.getenv('APPDATA')
    path=f"{appdata}\\Elgato"
    try:
        os.startfile(path)  # Windows
        print(f"Opened folder: {path}")
    except FileNotFoundError:
        print(f"Folder not found: {path}")
    except Exception as e:
        print(f"Error opening folder {path}: {e}")

def open_folder(path):
    try:
        os.startfile(path)  # Windows
        print(f"Opened folder: {path}")
    except FileNotFoundError:
        print(f"Folder not found: {path}")
    except Exception as e:
        print(f"Error opening folder {path}: {e}")

def get_appdata_path(app):
    appdata = os.getenv('APPDATA')
    folder_path = f"{appdata}\\Elgato\\{app}"
    return folder_path

def kill_process_by_name(input_name):
    if input_name == "CameraHub":
        app_name= "Camera Hub.exe"
    else:
        app_name=input_name + ".exe"
    
    for process in psutil.process_iter(['pid', 'name']):
        if process.info['name'] == app_name:
            try:
                psutil.Process(process.info['pid']).terminate()
                print(f"Terminated process: {process.info['name']} (PID: {process.info['pid']})")
            except psutil.NoSuchProcess:
                print(f"No such process: {app_name}")
            except Exception as e:
                print(f"Error terminating process {app_name}: {e}")
        
def delete_folder(folder_path):
    if os.path.exists(folder_path):
        try:
            shutil.rmtree(folder_path)
            messagebox.showinfo("Success", f"Folder '{folder_path}' has been deleted.")
        except Exception as e:
            messagebox.showerror("Error", f"Error deleting folder: {e}")
    else:
        messagebox.showwarning("Warning", f"Folder '{folder_path}' does not exist.")

def clear_folder(folder_path: str):
    if not os.path.exists(folder_path):
        print(f"❌ no exist: {folder_path}")
        return
    for item in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item)
        try:
            if os.path.isfile(item_path) or os.path.islink(item_path):
                os.unlink(item_path) 
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path)  
        except Exception as e:
            print(f"⚠️ can't del '{item_path}': {e}")
    print(f"✅ Clean: {folder_path}")
 

def uninstall_app(app):
    if '.' in app:
        powershell_command = f"Get-AppxPackage *{app}* | Remove-AppxPackage"
        try:
            subprocess.run(["powershell", "-Command", powershell_command], check=True)
            print(f"✅ Uninstalled '{app}'")
        except subprocess.CalledProcessError as e:
            print(f"❌ Error: {e}")
    else:
        app_name = apps[app]
        try:
            # Escaping quotes for the correct PowerShell command
            command = f'Start-Process cmd -ArgumentList "/c wmic product where ""name = \'{app_name}\'"" call uninstall" -Verb runAs'
            # Running the PowerShell command using subprocess
            process = subprocess.run(["powershell", "-Command", command], capture_output=True, text=True)

            # Check if the command executed successfully
            if process.returncode == 0:
                print(f"Uninstall initiated for {app_name}.")
            else:
                print(f"Error: {process.stderr}")
        except Exception as e:
            print(f"An error occurred: {e}")

def web_mp():
    webbrowser.open("https://marketplace.elgato.com/")

def web_dowload():
    webbrowser.open("https://www.elgato.com/us/en/s/downloads")

def create_app_buttons(root, app_name, i):
    if '.' in app_name:
        appdata = os.getenv('APPDATA')
        appdata_base = os.path.dirname(appdata)
        app_data_path = f"{appdata_base}\\Local\\Packages\\{apps[app_name]}"
    else:
        app_data_path = get_appdata_path(app_name)

    l1 = Label(master, text = app_name)
    l1.grid(row = i, column = 1, sticky = E, pady = 5)

    open_button = Button(master, text=f"Open", command=lambda: open_app(app_name))
    quit_button = Button(master, text=f"Quit", command=lambda: kill_process_by_name(app_name))
    del_appdata_button = Button(master, text=f"Delete appdata", command=lambda p=app_data_path: clear_folder(p))
    open_appdata_folder_button = Button(master, text=f"appdata folder", command=lambda p=app_data_path: open_folder(p))
    uninstall_app_button = Button(master, text=f"Uninstall", command=lambda p=app_name: uninstall_app(p))

    open_button.grid(row = i, column = 2)
    quit_button.grid(row = i, column = 3)
    del_appdata_button.grid(row = i, column = 4, padx=5)
    open_appdata_folder_button.grid(row = i, column = 5)
    uninstall_app_button.grid(row = i, column = 6, padx=5)


appdata_button = Button(master, text=f"Open Appdata", command=lambda: open_elgato_appdata_folder())
appdata_button.grid(row = 0, column = 2, pady = 0, columnspan = 2, sticky = W)

apps_installed = Button(master, text=f"Programs and Feature", command=lambda: os.system("appwiz.cpl"))
apps_installed.grid(row = 1, column = 2, pady = 5, columnspan = 2, sticky = W)

control_panel = Button(master, text=f"Control Panel", command=lambda: os.system("control"))
control_panel.grid(row = 1, column = 3, columnspan = 2)

open_web_mp = Button(master, text=f"Market Place", command=lambda: web_mp())
open_web_mp.grid(row = 2, column = 2, pady = 1, columnspan = 2, sticky = W)

open_web_download = Button(master, text=f"Elgato Download", command=lambda: web_dowload())
open_web_download.grid(row = 2, column = 3, pady = 1, columnspan = 2, sticky = W)


i=3
for app in apps:
    create_app_buttons(master, app, i)
    i+=1

master.mainloop()
