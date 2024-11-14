from math import ceil
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter import dialog
import webbrowser
import psutil
import subprocess
import wmi
import win32com.client
import win32serviceutil

program_name = "CZFEdit"

# Define the main application class
class LaptopChecklistApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Laptop Controle Checklist")
        self.root.attributes()
        
        self.filename = None

        # Define the fields for the checklist
        self.order_number = tk.StringVar()
        self.brand_name = tk.StringVar()
        self.windows_version = tk.StringVar()
        self.battery_health = tk.StringVar()
        self.ram = tk.StringVar()
        self.storage_space = tk.StringVar()
        self.cpu_name = tk.StringVar()
        self.updates_installed = tk.StringVar()
        self.hp_hotkeys_installed = tk.StringVar()
        self.conexant_installed = tk.StringVar()
        self.drivers_up_to_date = tk.BooleanVar()
        self.audio_works = tk.BooleanVar()
        self.keyboard_works = tk.BooleanVar()
        self.touchscreen_works = tk.BooleanVar()
        self.camera_works = tk.BooleanVar()
        self.office_activated = tk.BooleanVar()
        self.controller_name = tk.StringVar()
        
        # Initialize menu bar
        self.create_menu()
        
        # Create checklist form
        self.create_form()
    
    def create_menu(self):
        menu_bar = tk.Menu(self.root)

        # File Menu
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Nieuw...", command=self.new_file)
        file_menu.add_command(label="Openen...", command=self.open_file)
        file_menu.add_command(label="Opslaan", command=self.save_file)
        file_menu.add_command(label="Opslaan als...", command=self.save_as_file)
        file_menu.add_separator()
        file_menu.add_command(label="Afsluiten", command=self.root.quit)
        menu_bar.add_cascade(label="Bestand", menu=file_menu)

        # Edit Menu
        edit_menu = tk.Menu(menu_bar, tearoff=0)
        edit_menu.add_command(label="Bewerk ruwe gegevens...", command=self.edit_raw_data)
        #menu_bar.add_cascade(label="Bewerken", menu=edit_menu)

        # Controller Menu
        controller_menu = tk.Menu(menu_bar, tearoff=0)
        controller_menu.add_command(label="Beheer eerste controllers...", command=self.manage_first_controllers)
        controller_menu.add_command(label="Beheer tweede controllers...", command=self.manage_second_controllers)
        # More controllers as required
        #menu_bar.add_cascade(label="Controlleur", menu=controller_menu)

        # Checks Menu
        checks_menu = tk.Menu(menu_bar, tearoff=0)
        general_checks = tk.Menu(checks_menu, tearoff=0)
        general_checks.add_command(label="Scan voor Windows Updates...", command=self.start_update_scan)
        general_checks.add_separator()
        general_checks.add_command(label="Apparaatbeheer...", command=self.open_device_manager)
        general_checks.add_command(label="Geluid...", command=self.open_sound_manager)
        general_checks.add_command(label="Toetsenbord...", state="disabled", command=self.open_keyboard_tester)
        general_checks.add_command(label="Touchscreen...", command=self.open_touchscreen_tester)
        general_checks.add_command(label="Camera...", command=self.open_camera)
        general_checks.add_separator()
        general_checks.add_command(label="Installeer MS Office 2021...", command=self.install_ms_office)
        checks_menu.add_cascade(label="Algemeen", menu=general_checks)

        hp_checks = tk.Menu(checks_menu, tearoff=0)
        hp_checks.add_command(label="Installeer Hotkeys...", command=self.install_hp_hotkeys)
        hp_checks.add_command(label="Deactiveer Conexant Audio Service...", command=self.disable_conexant_audio)
        checks_menu.add_cascade(label="HP", menu=hp_checks)

        menu_bar.add_cascade(label="Controles", menu=checks_menu)
        
        # Advanced Menu
        advanced_menu = tk.Menu(menu_bar, tearoff=0)
        advanced_menu.add_command(label="Open Opdrachtprompt", command=self.open_cmd)
        advanced_menu.add_command(label="Open Opdrachtprompt als Administrator", command=self.open_admin_cmd)
        menu_bar.add_cascade(label="Geavanceerd", menu=advanced_menu)

        # Help Menu
        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="Documentatie", command=self.open_documentation)
        help_menu.add_command(label="Over...", command=self.show_about)
        help_menu.add_command(label="Bluesky", command=self.open_bluesky)
        menu_bar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menu_bar)

    def create_form(self):
        # Order Number
        order_number = None
        ttk.Label(self.root, text="Ordernummer:").grid(row=0, column=0, padx=5, pady=5)
        self.order = ttk.Entry(self.root, textvariable=order_number)
        self.order.grid(row=0, column=1, padx=5, pady=5)

        # Laptop Brand
        brand_text = None
        ttk.Label(self.root, text="Merk Laptop:").grid(row=1, column=0, padx=5, pady=5)
        self.brand = ttk.Combobox(self.root, values=["Acer", "Asus", "Dell", "Dynabook", "Dell", "Fujitsu", "HP", "Lenovo", "Surface", "Toshiba"], textvariable=brand_text)
        self.brand.grid(row=1, column=1, padx=5, pady=5)

        # Windows Version
        winver_text = None
        ttk.Label(self.root, text="Windows-versie:").grid(row=2, column=0, padx=5, pady=5)
        self.windows_version = ttk.Combobox(self.root, values=["Windows 10 Home", "Windows 10 Pro", "Windows 11 Home", "Windows 11 Pro"], textvariable=winver_text)
        self.windows_version.grid(row=2, column=1, padx=5, pady=5)
        
        # Battery Health
        ttk.Label(self.root, text="Kwaliteit Accu:").grid(row=3, column=0, padx=5, pady=5)
        self.battery_health = ttk.Label(self.root, text=self.get_battery_health())
        battery_health = self.get_battery_health()
        
        # Set text color based on the battery health value
        if battery_health == "Geen batterij!" or battery_health == "Onbekend":
            self.battery_health.config(foreground="red")
        else:
            self.battery_health.config(foreground="black")  # Optional: set to black for valid percentages

        self.battery_health.grid(row=3, column=1, padx=5, pady=5)

        # RAM in GB
        ttk.Label(self.root, text="RAM (GB):").grid(row=4, column=0, padx=5, pady=5)
        self.ram_label = ttk.Label(self.root, text=self.get_ram())
        self.ram_label.grid(row=4, column=1, padx=5, pady=5)

        # Storage space of C: drive
        ttk.Label(self.root, text="Opslagruimte C: (GB):").grid(row=5, column=0, padx=5, pady=5)
        self.storage_label = ttk.Label(self.root, text=self.get_storage())
        self.storage_label.grid(row=5, column=1, padx=5, pady=5)

        # CPU Name
        ttk.Label(self.root, text="CPU Naam:").grid(row=6, column=0, padx=5, pady=5)
        self.cpu_label = ttk.Label(self.root, text=self.get_cpu())
        self.cpu_label.grid(row=6, column=1, padx=5, pady=5)

        # Are updates installed?
        ttk.Label(self.root, text="Updates geinstalleerd?").grid(row=7, column=0, padx=5, pady=5)
        self.updates_label = ttk.Label(self.root, text=self.check_updates())
        self.updates_label.grid(row=7, column=1, padx=5, pady=5)

        # Is HP Hotkeys installed?
        ttk.Label(self.root, text="HP Hotkeys geinstalleerd?").grid(row=8, column=0, padx=5, pady=5)
        self.hp_hotkeys_label = ttk.Label(self.root, text=self.check_hp_hotkeys())
        self.hp_hotkeys_label.grid(row=8, column=1, padx=5, pady=5)

        # Is Conexant Audio disabled or nonexistent?
        ttk.Label(self.root, text="Conexant Audio gedeactiveerd of niet aanwezig?").grid(row=9, column=0, padx=5, pady=5)
        self.conexant_audio_label = ttk.Label(self.root, text=self.check_conexant_audio())
        self.conexant_audio_label.grid(row=9, column=1, padx=5, pady=5)

        # Drivers up-to-date (Checkbox)
        ttk.Label(self.root, text="Stuurprogramma's up-to-date?").grid(row=10, column=0, padx=5, pady=5)
        self.drivers_check = tk.BooleanVar()
        ttk.Checkbutton(self.root, name="drivers_box", variable=self.drivers_check).grid(row=10, column=1, padx=5, pady=5)

        # Audio works (Checkbox)
        audio_text = None
        ttk.Label(self.root, text="Audio werkt?").grid(row=11, column=0, padx=5, pady=5)
        self.audio_check = tk.BooleanVar()
        ttk.Checkbutton(self.root, name="audio_box", variable=self.audio_check, textvariable=audio_text).grid(row=11, column=1, padx=5, pady=5)

        # Microphone works (Checkbox)
        mic_text = None
        ttk.Label(self.root, text="Microfoon werkt?").grid(row=12, column=0, padx=5, pady=5)
        self.microphone_check = tk.BooleanVar()
        ttk.Checkbutton(self.root, name="mic_box", variable=self.microphone_check, textvariable=mic_text).grid(row=12, column=1, padx=5, pady=5)

        # Keyboard works (Checkbox)
        key_text = None
        ttk.Label(self.root, text="Toetsenbord werkt?").grid(row=13, column=0, padx=5, pady=5)
        self.keyboard_check = tk.BooleanVar()
        ttk.Checkbutton(self.root, name="key_box", variable=self.keyboard_check, textvariable=key_text).grid(row=13, column=1, padx=5, pady=5)
        
        # Touchscreen works (Checkbox)
        touch_text = None
        ttk.Label(self.root, text="Touchscreen werkt?").grid(row=14, column=0, padx=5, pady=5)
        self.touch_check = tk.BooleanVar()
        ttk.Checkbutton(self.root, name="touch_box", variable=self.touch_check, textvariable=touch_text).grid(row=14, column=1, padx=5, pady=5)

        # Camera works (Checkbox)
        cam_text = None
        ttk.Label(self.root, text="Camera werkt?").grid(row=15, column=0, padx=5, pady=5)
        self.camera_check = tk.BooleanVar()
        ttk.Checkbutton(self.root, name="cam_box", variable=self.camera_check, textvariable=cam_text).grid(row=15, column=1, padx=5, pady=5)
        
        # Office installed & activated (Checkbox)
        office_text = None
        ttk.Label(self.root, text="Office geactiveerd?").grid(row=16, column=0, padx=5, pady=5)
        self.office_check = tk.BooleanVar()
        ttk.Checkbutton(self.root, name="office_box", variable=self.office_check, textvariable=office_text).grid(row=16, column=1, padx=5, pady=5)

        # Name of person doing the checklist
        controller_name = None
        ttk.Label(self.root, text="Naam uitvoerder:").grid(row=17, column=0, padx=5, pady=5)
        self.executor_name = ttk.Entry(self.root, textvariable=controller_name)
        self.executor_name.grid(row=17, column=1, padx=5, pady=5)

    # System information retrieval functions
    def get_ram(self):
        return str(f"{ceil(int(psutil.virtual_memory().total / (1024 ** 3)))} GB")

    def get_storage(self):
        return str(f"{ceil(int(psutil.disk_usage('C:').total / (1024 ** 3)))} GB")

    def get_cpu(self):
        return str(subprocess.check_output("wmic cpu get name").decode().split('\n')[1].strip())
    
    def get_battery_health(self):
        # Initialize WMI interface
        c = wmi.WMI()
    
        # Retrieve battery information
        try:
            battery = c.Win32_Battery()[0]
        except:
            return "Geen batterij!"
    
        # Get design capacity and full charge capacity
        design_capacity = battery.DesignCapacity
        full_charge_capacity = battery.FullChargeCapacity
    
        # Calculate battery health percentage
        if design_capacity and full_charge_capacity:
            battery_health = (full_charge_capacity / design_capacity) * 100
            return str(f"{battery_health:.2f}%")
        else:
            return "Onbekend"

    def check_updates(self):
        try:
            # Initialize the update session
            session = win32com.client.Dispatch("Microsoft.Update.Session")
            updateSearcher = session.CreateUpdateSearcher()

            # Search for pending updates (excluding preview updates)
            searchResult = updateSearcher.Search("IsInstalled=0 and IsHidden=0")

            if searchResult.Updates.Count == 0:
                return "Windows is bijgewerkt!"
            else:
                return f"{searchResult.Updates.Count} update(s) beschikbaar"
        except Exception as e:
            return str(e)

    def check_hp_hotkeys(self):
        service_name = "HPSysInfo"  # This is the service name for HP Hotkey UWP Service.
        pass

        try:
            # Check if the service is running
            status = win32serviceutil.QueryServiceStatus(service_name)
        
            # The status is a tuple, and the first value is the state of the service (4 = running)
            if status[1] == 4:
                return "Ja"
            else:
                return "Wordt niet uitgevoerd"
    
        except:
            return "Nee"

    def check_conexant_audio(self):
        service_name = "CxUtilSvc"  # Conexant Audio Service name
        pass

        try:
            # Check if the service is running
            status = win32serviceutil.QueryServiceStatus(service_name)
        
            # The status is a tuple, and the first value is the state of the service (4 = running)
            if status[1] == 4:
                return "Nee"
            else:
                return "Wordt niet uitgevoerd"
    
        except:
            return "Ja"

    # Function placeholders for menu actions
    def new_file(self):
        # Clear current data and reset the checklist
        if self.ask_save_changes():
            python = sys.executable  # Path to the Python interpreter
            os.execv(python, ['python'] + sys.argv)  # Executes the script with the same arguments

    def open_file(self):
        self.filename = filedialog.askopenfilename(defaultextension=".czf", filetypes=[("CZF-bestanden", "*.czf")])
        # Read the file and populate the fields
        with open(self.filename, 'r') as file:
            lines = file.readlines()

        for line in lines:
            if line.startswith("ORDER"):
                self.order.insert(0, line.split("ORDER")[1].strip())
            elif line.startswith("BRAND"):
                self.brand.set(line.split("BRAND")[1].strip())
            elif line.startswith("WINVER"):
                self.windows_version.set(line.split("WINVER")[1].strip())
            elif line.startswith("DRVUTD"):
                self.drivers_check.set(line.split("DRVUTD")[1].strip())
            elif line.startswith("POSAUD"):
                self.audio_check.set(line.split("POSAUD")[1].strip())
            elif line.startswith("POSKEY"):
                self.keyboard_check.set(line.split("POSKEY")[1].strip())
            elif line.startswith("POSTCH"):
                self.touch_check.set(line.split("POSTCH")[1].strip())
            elif line.startswith("POSCAM"):
                self.camera_check.set(line.split("POSCAM")[1].strip())
            elif line.startswith("OFFACT"):
                self.office_check.set(line.split("OFFACT")[1].strip())
            elif line.startswith("CTRLNAME"):
                self.executor_name.insert(0, line.split("CTRLNAME")[1].strip())

    def save_file(self):
        if self.filename:
            self.save_to_file(self.filename)
        else:
            self.save_as_file()

    def save_as_file(self):
        file = filedialog.asksaveasfilename(defaultextension=".czf", filetypes=[("CZF-bestanden", "*.czf")])
        if file:
            self.filename = file
            self.save_to_file(file)

    def save_to_file(self, filename):
        # Format the data and save it into the .CZF file
        with open(filename, 'w') as file:
            file.write("+=========================+\n")
            file.write(f"ORDER {self.order.get()}\n")
            file.write(f"BRAND {self.brand.get()}\n")
            file.write(f"WINVER {self.windows_version.get()}\n")
            file.write(f"BATQLTY {self.get_battery_health()}\n")
            file.write(f"RAM {self.get_ram()}\n")
            file.write(f"STRSPC {self.get_storage()}\n")
            file.write(f"CPU {self.get_cpu()}\n")
            file.write(f"WINUTD {self.check_updates()}\n")
            file.write(f"HPHKINST {self.check_hp_hotkeys()}\n")
            file.write(f"CONEXANT {self.check_conexant_audio()}\n")
            file.write(f"DRVUTD {self.drivers_check.get()}\n")
            file.write(f"POSAUD {self.audio_check.get()}\n")
            file.write(f"POSKEY {self.keyboard_check.get()}\n")
            file.write(f"POSTCH {self.touch_check.get()}\n")
            file.write(f"POSCAM {self.camera_check.get()}\n")
            file.write(f"OFFACT {self.office_check.get()}\n")
            file.write(f"CTRLNAME {self.executor_name.get()}\n")
            file.write("+=========================+\n")
        messagebox.showinfo("Opgeslagen!", f"Checklist opgeslagen als {filename}")

    def export_as_txt(self):
        pass
    
    def exit_program(self):
        if self.ask_save_changes():
            self.root.quit()
            
    def edit_raw_data(self):
        pass
    
    def manage_first_controllers(self):
        pass
    
    def manage_second_controllers(self):
        pass

    def open_cmd(self):
        subprocess.Popen("cmd")

    def open_admin_cmd(self):
        subprocess.Popen("cmd", shell=True)  # Will require elevation handling
        
    def open_regedit(self):
        subprocess.Popen("regedit", shell=True)

    def start_update_scan(self):
        subprocess.Popen("control update")
        subprocess.Popen("usoclient StartInteractiveScan")
        
    def open_device_manager(self):
        subprocess.Popen("devmgmt.msc", shell=True)
        self.drivers_check = tk.BooleanVar(value=True)
        
    def open_sound_manager(self):
        subprocess.Popen("mmsys.cpl", shell=True)
        self.audio_check = tk.BooleanVar(value=True)
        self.microphone_check = tk.BooleanVar(value=True)
        
    def open_camera(self):
        subprocess.Popen("start microsoft.windows.camera:", shell=True)
        self.camera_check = tk.BooleanVar(value=True)
        
    def open_keyboard_tester(self):
        pass # The Keyboard Tester is not done yet!
    
    def open_touchscreen_tester(self):
        subprocess.Popen("plugins\tt.exe")
    
    def install_ms_office(self):
        pass
    
    def install_hp_hotkeys(self):
        pass
    
    def disable_conexant_audio(self):
        pass

    def open_documentation(self):
        webbrowser.open("", 2)
    
    def show_about(self):
        pass
    
    def open_bluesky(self):
        webbrowser.open("https://bsky.app/profile/porky.live", 2)
        
    def ask_save_changes(self):
        # Ask the user if they want to save changes before proceeding with a new file or closing
        if self.filename:  # Only ask if there's a file to save changes
            return messagebox.askyesno("Wijzigingen opslaan", "Weet u zeker dat u een nieuwe checklist wilt beginnen? Alle niet-opgeslagen wijzigingen gaan verloren!")
        return True  # If no file is open, allow the operation without saving
    
    # More function placeholders...

# Run the application
root = tk.Tk()

root.resizable(False, False)

app = LaptopChecklistApp(root)
root.mainloop()
