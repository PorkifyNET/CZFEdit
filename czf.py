from cProfile import label
import datetime
from datetime import datetime
from math import ceil
import os
import shutil
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import webbrowser
import psutil
import subprocess
import wmi
import win32com.client
import win32serviceutil

program_name = "CZFEdit"
program_version = "v1.0"
global program_location
program_location = "D:/PorkifyNET/CZFEdit"

def snap_to_nearest_power_of_2(n):
    # Edge case for n = 0 or negative values
    if n <= 0:
        return 1

    # Find the next power of 2 greater than or equal to n
    power_of_2_above = 1 << (n - 1).bit_length()
    # Find the previous power of 2 (divide power_of_2_above by 2 if n is not exactly a power of 2)
    power_of_2_below = power_of_2_above // 2

    # Return the closer power of 2
    return power_of_2_below if (n - power_of_2_below) < (power_of_2_above - n) else power_of_2_above

# Define the main application class
class LaptopChecklistApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"CZFEdit: Controle Checklist")
        self.root.attributes()
        
        self.filename = None

        # Define the fields for the checklist
        self.order_number = tk.StringVar()
        self.cz_number = tk.StringVar()
        self.order_date = tk.StringVar()
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
        self.execution_date = tk.StringVar()
        
        # Create checklist form
        self.create_form()

        # Initialize menu bar
        self.create_menu()
        
    def create_form(self):
        # Order Number
        order_number = None
        ttk.Label(self.root, text="Ordernummer:").grid(row=0, column=0, padx=5, pady=5)
        self.order = ttk.Entry(self.root, textvariable=order_number)
        self.order.grid(row=0, column=1, padx=5, pady=5)

        # CZ Number
        cz_number = None
        ttk.Label(self.root, text="CZ-nummer:").grid(row=1, column=0, padx=5, pady=5)
        self.cz = ttk.Entry(self.root, textvariable=cz_number)
        self.cz.grid(row=1, column=1, padx=5, pady=5)
        
        # Date as specified on the order
        order_date = None
        ttk.Label(self.root, text="Orderdatum:").grid(row=2, column=0, padx=5, pady=5)
        self.date = ttk.Entry(self.root, textvariable=order_date)
        self.date.grid(row=2, column=1, padx=5, pady=5)

        # Laptop Brand
        brand_text = None
        ttk.Label(self.root, text="Merk Laptop:").grid(row=3, column=0, padx=5, pady=5)
        self.brand = ttk.Combobox(self.root, values=["Acer", "Asus", "Dell", "Dynabook / Toshiba", "Fujitsu", "HP", "Lenovo", "Microsoft Surface"], textvariable=brand_text)
        self.brand.grid(row=3, column=1, padx=5, pady=5)

        # Windows Version
        winver_text = None
        ttk.Label(self.root, text="Windows-versie:").grid(row=4, column=0, padx=5, pady=5)
        self.windows_version = ttk.Combobox(self.root, values=["Windows 10 Home", "Windows 10 Pro", "Windows 11 Home", "Windows 11 Pro"], textvariable=winver_text)
        self.windows_version.grid(row=4, column=1, padx=5, pady=5)
        
        # Battery Health
        ttk.Label(self.root, text="Kwaliteit Accu:").grid(row=5, column=0, padx=5, pady=5)
        self.battery_health_label = ttk.Label(self.root, text=self.get_battery_health())
        battery_health = self.get_battery_health()
        
        # Set text color based on the battery health value
        if battery_health == "Geen batterij!" or battery_health == "Onbekend":
            self.battery_health_label.config(foreground="red")
        else:
            self.battery_health_label.config(foreground="black")  # Optional: set to black for valid percentages

        self.battery_health_label.grid(row=5, column=1, padx=5, pady=5)

        # RAM in GB
        ttk.Label(self.root, text="RAM (GB):").grid(row=6, column=0, padx=5, pady=5)
        self.ram_label = ttk.Label(self.root, text=self.get_ram())
        self.ram_label.grid(row=6, column=1, padx=5, pady=5)

        # Storage space of C: drive
        ttk.Label(self.root, text="Opslagruimte C: (GB):").grid(row=7, column=0, padx=5, pady=5)
        self.storage_label = ttk.Label(self.root, text=self.get_storage())
        self.storage_label.grid(row=7, column=1, padx=5, pady=5)

        # CPU Name
        ttk.Label(self.root, text="CPU Naam:").grid(row=8, column=0, padx=5, pady=5)
        self.cpu_label = ttk.Label(self.root, text=self.get_cpu())
        self.cpu_label.grid(row=8, column=1, padx=5, pady=5)

        # Are updates installed?
        ttk.Label(self.root, text="Updates geinstalleerd?").grid(row=9, column=0, padx=5, pady=5)
        self.updates_label = ttk.Label(self.root, text=self.check_updates())
        self.updates_label.grid(row=9, column=1, padx=5, pady=5)

        # Is HP Hotkeys installed?
        ttk.Label(self.root, text="HP Hotkeys geinstalleerd?").grid(row=10, column=0, padx=5, pady=5)
        self.hp_hotkeys_label = ttk.Label(self.root, text=self.check_hp_hotkeys())
        self.hp_hotkeys_label.grid(row=10, column=1, padx=5, pady=5)

        # Is Conexant Audio disabled or nonexistent?
        ttk.Label(self.root, text="Conexant Audio gedeactiveerd of niet aanwezig?").grid(row=11, column=0, padx=5, pady=5)
        self.conexant_audio_label = ttk.Label(self.root, text=self.check_conexant_audio())
        self.conexant_audio_label.grid(row=11, column=1, padx=5, pady=5)

        # Drivers up-to-date (Checkbox)
        ttk.Label(self.root, text="Stuurprogramma's up-to-date?").grid(row=12, column=0, padx=5, pady=5)
        self.drivers_check = tk.BooleanVar()
        ttk.Checkbutton(self.root, name="drivers_box", variable=self.drivers_check).grid(row=12, column=1, padx=5, pady=5)

        # Audio works (Checkbox)
        audio_text = None
        ttk.Label(self.root, text="Audio werkt?").grid(row=13, column=0, padx=5, pady=5)
        self.audio_check = tk.BooleanVar()
        ttk.Checkbutton(self.root, name="audio_box", variable=self.audio_check, textvariable=audio_text).grid(row=13, column=1, padx=5, pady=5)

        # Microphone works (Checkbox)
        mic_text = None
        ttk.Label(self.root, text="Microfoon werkt?").grid(row=14, column=0, padx=5, pady=5)
        self.microphone_check = tk.BooleanVar()
        ttk.Checkbutton(self.root, name="mic_box", variable=self.microphone_check, textvariable=mic_text).grid(row=14, column=1, padx=5, pady=5)

        # Keyboard works (Checkbox)
        key_text = None
        ttk.Label(self.root, text="Toetsenbord werkt?").grid(row=15, column=0, padx=5, pady=5)
        self.keyboard_check = tk.BooleanVar()
        ttk.Checkbutton(self.root, name="key_box", variable=self.keyboard_check, textvariable=key_text).grid(row=15, column=1, padx=5, pady=5)
        
        # Touchscreen works (Checkbox)
        touch_text = None
        ttk.Label(self.root, text="Touchscreen werkt?").grid(row=16, column=0, padx=5, pady=5)
        self.touch_check = tk.BooleanVar()
        ttk.Checkbutton(self.root, name="touch_box", variable=self.touch_check, textvariable=touch_text).grid(row=16, column=1, padx=5, pady=5)

        # Camera works (Checkbox)
        cam_text = None
        ttk.Label(self.root, text="Camera werkt?").grid(row=17, column=0, padx=5, pady=5)
        self.camera_check = tk.BooleanVar()
        ttk.Checkbutton(self.root, name="cam_box", variable=self.camera_check, textvariable=cam_text).grid(row=17, column=1, padx=5, pady=5)
        
        # Office installed & activated (Checkbox)
        office_text = None
        ttk.Label(self.root, text="Office geactiveerd?").grid(row=18, column=0, padx=5, pady=5)
        self.office_check = tk.BooleanVar()
        ttk.Checkbutton(self.root, name="office_box", variable=self.office_check, textvariable=office_text).grid(row=18, column=1, padx=5, pady=5)

        # Name of person doing the checklist
        controller_name = None
        ttk.Label(self.root, text="Naam uitvoerder:").grid(row=19, column=0, padx=5, pady=5)
        self.executor_name = ttk.Entry(self.root, textvariable=controller_name)
        self.executor_name.grid(row=19, column=1, padx=5, pady=5)
        
        # Date of CZF creation
        czf_create_date = None
        ttk.Label(self.root, text="Datum aangemaakt:").grid(row=20, column=0, padx=5, pady=5)
        self.execution_date_label = ttk.Label(self.root, text=self.get_current_date())
        self.execution_date_label.grid(row=20, column=1, padx=5, pady=5)
    
    def create_menu(self):
        menu_bar = tk.Menu(self.root)
        
        def is_wifi_enabled():
            # Get all network interfaces
            interfaces = psutil.net_if_stats()

            # Check common Wi-Fi interface names (can vary by system)
            wifi_names = ["Wi-Fi", "wlan0", "wlan1", "Wireless Network Connection"]

            # Check each Wi-Fi interface to see if it's up
            for interface in wifi_names:
                if interface in interfaces and interfaces[interface].isup:
                    return True  # Wi-Fi is enabled

            return False  # Wi-Fi is disabled

        # File Menu
        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="Nieuw...", command=self.new_file)
        file_menu.add_command(label="Openen...", command=self.open_file)
        file_menu.add_command(label="Opslaan", command=self.save_file)
        file_menu.add_command(label="Opslaan als...", command=self.save_as_file)
        file_menu.add_separator()
        file_menu.add_command(label="Stel plugin-locatie in...", command=self.set_location)
        file_menu.add_separator()
        file_menu.add_command(label="Sluiten", command=self.root.quit)
        menu_bar.add_cascade(label="Bestand", menu=file_menu)

        # Controller Menu
        controller_menu = tk.Menu(menu_bar, tearoff=0)
        controller_menu.add_command(label="Beheer controllers...", command=self.manage_first_controllers)
        controller_menu.add_command(label="Beheer eindcontrollers...", command=self.manage_second_controllers)
        controller_menu.add_command(label="Beheer schoonmakers...", command=self.manage_cleaners)
        controller_menu.add_command(label="Beheer inpakkers...", command=self.manage_packagers)
        # More controllers as required
        #menu_bar.add_cascade(label="Controllers", menu=controller_menu)

        # Checks Menu
        checks_menu = tk.Menu(menu_bar, tearoff=0)
        checks_menu.add_command(label="Scan voor Windows Updates...", command=self.start_update_scan)
        checks_menu.add_separator()
        checks_menu.add_command(label="Verwijder wachtwoord leeftijdslimiet", command=self.reset_pw_age)
        checks_menu.add_command(label="Snel opstarten uitschakelen", command=self.disable_fast_startup)
        checks_menu.add_separator()
        checks_menu.add_command(label="Accu-informatie...", command=self.open_batterylifeinfo)
        checks_menu.add_command(label="Apparaatbeheer...", command=self.open_device_manager)
        checks_menu.add_command(label="Geluid...", command=self.open_sound_manager)
        checks_menu.add_command(label="Toetsenbord...", state="disabled", command=self.open_keyboard_tester)
        checks_menu.add_command(label="Touchscreen...", command=self.open_touchscreen_tester)
        checks_menu.add_command(label="Camera...", command=self.open_camera)
        checks_menu.add_separator()
        checks_menu.add_command(label="Windows Activeren...", command=self.open_windows_activation)
        checks_menu.add_separator()
        checks_menu.add_command(label="Installeer MS Office 2021...", command=self.install_ms_office)

        office_shortcuts = tk.Menu(checks_menu, tearoff=0)
        office_shortcuts.add_command(label="ProgramData", command=self.office_shortcuts_programdata)
        office_shortcuts.add_command(label="D: Schijf", command=self.office_shortcuts_d_drive)
        office_shortcuts.add_separator()
        office_shortcuts.add_command(label="Aangepast...", command=self.office_shortcuts_custom, state="disabled")
        checks_menu.add_cascade(label="Office Snelkoppelingen Maken", menu=office_shortcuts)

        office_activation = tk.Menu(checks_menu, tearoff=0)
        office_activation.add_command(label="DigitalProducts (Aanbevolen)", command=self.office_web_activation)
        office_activation.add_command(label="ADB (D:)", command=self.office_adb_activation)
        office_activation.add_command(label="MOAT / Automato", command=self.office_moat_activation)
        office_activation.add_command(label="PNET ID-Gen", state="disabled")
        checks_menu.add_cascade(label="Office Activeren", menu=office_activation)

        checks_menu.add_separator()
        
        network_enabled = is_wifi_enabled()
        acer_checks = tk.Menu(checks_menu, tearoff=0)
        acer_checks.add_checkbutton(label="Netwerk uitschakelen", variable=network_enabled, command=self.toggle_wifi)
        checks_menu.add_cascade(label="Acer", menu=acer_checks, state="disabled")
        
        asus_checks = tk.Menu(checks_menu, tearoff=0)
        checks_menu.add_cascade(label="Asus", menu=asus_checks, state="disabled")
        
        dell_checks = tk.Menu(checks_menu, tearoff=0)
        dell_driver_updates = tk.Menu(dell_checks, tearoff=0)
        dell_checks.add_cascade(label="Verwijder Aangepaste Bootscreen Logo", menu=dell_driver_updates)
        checks_menu.add_cascade(label="Dell", menu=dell_checks)

        dynabook_toshiba_checks = tk.Menu(checks_menu, tearoff=0)
        checks_menu.add_cascade(label="Dynabook / Toshiba", menu=dynabook_toshiba_checks, state="disabled")
        
        fujitsu_checks = tk.Menu(checks_menu, tearoff=0)
        fujitsu_checks.add_command(label="Charging Tool uitschakelen", command=self.disable_fujitsu_battery_charging_tool)
        checks_menu.add_cascade(label="Fujitsu", menu=fujitsu_checks)
        
        hp_checks = tk.Menu(checks_menu, tearoff=0)
        hp_checks.add_command(label="Installeer Hotkeys...", command=self.install_hp_hotkeys)
        hp_checks.add_command(label="Deactiveer Conexant Audio Service...", command=self.disable_conexant_audio)
        checks_menu.add_cascade(label="HP", menu=hp_checks)

        lenovo_checks = tk.Menu(checks_menu, tearoff=0)
        checks_menu.add_cascade(label="Lenovo", menu=lenovo_checks, state="disabled")
        
        surface_checks = tk.Menu(checks_menu, tearoff=0)
        checks_menu.add_cascade(label="Microsoft Surface", menu=surface_checks, state="disabled")

        menu_bar.add_cascade(label="Controles", menu=checks_menu)
        
        checks_menu.add_separator()
        checks_menu.add_command(label="Versturen naar CZ server...", command=self.send_to_server)
        
        # Advanced Menu
        advanced_menu = tk.Menu(menu_bar, tearoff=0)
        advanced_menu.add_command(label="Open Opdrachtprompt", command=self.open_cmd)
        advanced_menu.add_command(label="Open Register-editor", command=self.open_regedit)
        advanced_menu.add_separator()
        advanced_menu.add_command(label="Activatie-status", command=self.show_system_overview)
        advanced_menu.add_command(label="Open Snappy...", command=self.open_snappy)
        advanced_menu.add_separator()
        advanced_menu.add_command(label="Open CZFEdit map...", command=self.open_self)
        advanced_menu.add_separator()
        advanced_menu.add_command(label="Open .CZF-bestand als ruwe data...", command=self.edit_raw_data)
        menu_bar.add_cascade(label="Geavanceerd", menu=advanced_menu)
        
        # Refresh Menu
        refresh_menu = tk.Menu(menu_bar, tearoff=0)
        refresh_menu.add_command(label="Kwaliteit Accu", command=self.refresh_battery_health_label())
        refresh_menu.add_command(label="RAM", command=self.refresh_ram_label())
        refresh_menu.add_command(label="Opslagruimte C:", command=self.refresh_storage_label())
        refresh_menu.add_command(label="CPU Naam", command=self.refresh_cpu_label())
        refresh_menu.add_command(label="Updates", command=self.refresh_updates_label())
        refresh_menu.add_separator()
        refresh_menu.add_command(label="HP Hotkeys Service", command=self.refresh_hp_label())
        refresh_menu.add_command(label="Conexant Audio Service", command=self.refresh_conexant_label())
        menu_bar.add_cascade(label="Vernieuwen", menu=refresh_menu, state="disabled")

        #Shutdown Menu
        shutdown_menu = tk.Menu(menu_bar, tearoff=0)
        shutdown_menu.add_command(label="Afsluiten", command=self.manual_shutdown)
        shutdown_menu.add_command(label="Opnieuw opstarten", command=self.manual_restart)
        shutdown_menu.add_separator()
        shutdown_menu.add_command(label="Afsluiten naar Firmware Setup", command=self.firmware)
        shutdown_menu.add_command(label="Opnieuw opstarten naar herstelomgeving", command=self.winre)
        menu_bar.add_cascade(label="Afsluiten", menu=shutdown_menu)

        # Help Menu
        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(label="Documentatie", command=self.open_documentation)
        help_menu.add_command(label="Over...", command=self.show_about)
        help_menu.add_command(label="Meer freeware op Porkify.NET", command=self.open_porkifynet)
        menu_bar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menu_bar)

    # System information retrieval functions
    def get_ram(self):
        return str(f"{(ceil(int(psutil.virtual_memory().total / (1024 ** 3)))) + 1 } GB")

    def get_storage(self):
        return str(f"{ceil(int(psutil.disk_usage('C:').total / (1024 ** 3)))} GB ({snap_to_nearest_power_of_2(ceil(int(psutil.disk_usage('C:').total / (1024 ** 3))))})")

    def get_cpu(self):
        return str(subprocess.check_output("wmic cpu get name").decode().split('\n')[1].strip())
    
    def get_battery_health(self):
        battery = psutil.sensors_battery()
        if battery is None:
            return "Geen batterij!"
        else:
            # Run BatteryInfoView with /stext to save output as plain text
            subprocess.run([f"{program_location}/plugins/BatteryInfoView.exe", "/scomma", "battery_info.txt"], check=True)

            try:
                with open("battery_info.txt", "r") as file:
                    lines = file.readlines()

                    # Parse each line to find Design Capacity and Full Charge Capacity
                    battery_health_var = lines[10][15:].strip()
                    return battery_health_var
                
            except FileNotFoundError:
                return "Leesfout"
        
    def open_batterylifeinfo(self):
        subprocess.Popen(f"{program_location}/plugins/BatteryInfoView.exe", shell=True)
        
    def open_snappy(self):
        subprocess.Popen(f"{program_location}/plugins/Snappy64.exe", shell=True)

    def check_updates(self):
        try:
            # Initialize the update session
            session = win32com.client.Dispatch("Microsoft.Update.Session")
            updateSearcher = session.CreateUpdateSearcher()

            # Search for pending updates (excluding preview updates)
            searchResult = updateSearcher.Search("IsInstalled=0 and IsHidden=0")

            if searchResult.Updates.Count == 0:
                return "Windows is bijgewerkt!"
            elif type(searchResult.Updates.Count) == int:
                return f"{searchResult.Updates.Count} update(s) beschikbaar"
            else:
                return "Onbekend"
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
        
    def get_current_date(self):
        return str(f"{datetime.now().strftime("%d/%m/%Y")}")
        
    def toggle_wifi(self, enable=True):
        # Set the name of your Wi-Fi interface here
        interface_name = "Wi-Fi"  # Change this to match your Wi-Fi adapter name
    
        # Command to enable or disable Wi-Fi
        action = "enable" if enable else "disable"
        command = f'netsh interface set interface "{interface_name}" {action}'
    
        # Execute the command
        try:
            subprocess.run(command, shell=True)
            print(f"Wi-Fi {action}d successfully.")
        except subprocess.CalledProcessError:
            print(f"Failed to {action} Wi-Fi. Ensure you have admin privileges and the interface name is correct.")

    def reset_pw_age(self):
        try:
            subprocess.Popen("net accounts /maxpwage:unlimited", shell=True)
            subprocess.Popen("""Set-LocalUser -Name ^"Gebruiker^" -PasswordNeverExpires 1""", shell=True)
            messagebox.showinfo(program_name, "Leeftijdslimiet wachtwoord is succesvol verwijderd.")
        except Exception as e:
            messagebox.showerror(program_name, e)
            
    def disable_fast_startup(self):
        try:
            subprocess.Popen("powercfg /hibernate off", shell=True)
            messagebox.showinfo(program_name, "Snel opstarten is succesvol uitgeschakeld.")
        except Exception as e:
            messagebox.showerror(program_name, e)

    def disable_fujitsu_battery_charging_tool(self):
        try:
            subprocess.Popen("taskkill /f /im BatteryCtrlUpdate.exe", shell=True)
            messagebox.showinfo(program_name, "Fujitsu Battery Charging Tool is succesvol uitgeschakeld.")
        except Exception as e:
            messagebox.showerror(program_name, e)
            
    def refresh_battery_health_label(self):
        self.battery_health_label.config(text=self.get_battery_health())
        
    def refresh_ram_label(self):
        self.ram_label.config(text=self.get_ram())
        
    def refresh_storage_label(self):
        self.storage_label.config(text=self.get_storage())
        
    def refresh_cpu_label(self):
        self.cpu_label.config(text=self.get_cpu())
        
    def refresh_updates_label(self):
        self.updates_label.config(text=self.check_updates())
        
    def refresh_hp_label(self):
        self.hp_hotkeys_label.config(text=self.check_hp_hotkeys())
        
    def refresh_conexant_label(self):
        self.conexant_audio_label.config(text=self.check_conexant_audio())

    # Function placeholders for menu actions
    def new_file(self):
        # Clear current data and reset the checklist
        if self.ask_save_changes():
            python = sys.executable  # Path to the Python interpreter
            os.execv(python, ['python'] + sys.argv)  # Executes the script with the same arguments

    def open_file(self):
        self.filename = filedialog.askopenfilename(defaultextension=".czf", filetypes=[("CZF-bestanden", "*.czf"), ("Ruwe data-bestanden", "*.txt")])
        # Read the file and populate the fields
        with open(self.filename, 'r') as file:
            lines = file.readlines()
            
            root.title(f"CZFEdit: Controle Checklist - {self.filename}")

        try:
            for line in lines:
                if line.startswith("ORDER"):
                    self.order.delete(0, tk.END)
                    self.order.insert(0, line.split("ORDER")[1].strip())
                elif line.startswith("CZNUM"):
                    self.cz.delete(0, tk.END)
                    self.cz.insert(0, line.split("CZNUM")[1].strip())
                elif line.startswith("ORDDAT"):
                    self.date.delete(0, tk.END)
                    self.date.insert(0, line.split("ORDDAT")[1].strip())
                elif line.startswith("BRAND"):
                    self.brand.set(line.split("BRAND")[1].strip())
                elif line.startswith("WINVER"):
                    self.windows_version.set(line.split("WINVER")[1].strip())
                elif line.startswith("BATQLTY"):
                    self.battery_health_label.config(text=line.split("BATQLTY")[1].strip())
                elif line.startswith("RAM"):
                    self.ram_label.config(text=line.split("RAM")[1].strip())
                elif line.startswith("STRSPC"):
                    self.storage_label.config(text=line.split("STRSPC")[1].strip())
                elif line.startswith("CPU"):
                    self.cpu_label.config(text=line.split("CPU")[1].strip())
                elif line.startswith("WINUTD"):
                    self.updates_label.config(text=line.split("WINUTD")[1].strip())
                elif line.startswith("HPHKINST"):
                    self.hp_hotkeys_label.config(text=line.split("HPHKINST")[1].strip())
                elif line.startswith("CONEXANT"):
                    self.conexant_audio_label.config(text=line.split("CONEXANT")[1].strip())
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
                    self.executor_name.delete(0, tk.END)
                    self.executor_name.insert(0, line.split("CTRLNAME")[1].strip())
                elif line.startswith("CTRLDATE"):
                    self.execution_date_label.config(text=line.split("CTRLDATE")[1].strip(),foreground="black")
        except Exception as e:
            messagebox.showerror(program_name, e)

    def save_file(self):
        if self.filename:
            self.save_to_file(self.filename)
            root.title(f"CZFEdit: Controle Checklist - {self.filename}")
        else:
            self.save_as_file()

    def save_as_file(self):
        file = filedialog.asksaveasfilename(defaultextension=".czf", filetypes=[("CZF-bestanden", "*.czf")], initialfile=self.order.get())
        if file:
            self.filename = file
            self.save_to_file(file)

    def save_to_file(self, filename):
        # Format the data and save it into the .CZF file
        with open(filename, 'w') as file:
            file.write("+=========================+\n")
            file.write(f"ORDER {self.order.get()}\n")
            file.write(f"CZNUM {self.cz.get()}\n")
            file.write(f"ORDDAT {self.date.get()}\n")
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
            file.write(f"CTRLDATE {self.get_current_date()}\n")
            file.write("+=========================+\n")
        messagebox.showinfo("Opgeslagen!", f"Checklist opgeslagen als {filename}")
        root.title(f"CZFEdit: Controle Checklist - {self.filename}")

    def set_location(self):
        global program_location
        if program_location == None:
            program_location_suggestion = "D:/PorkifyNET/CZFEdit"
        else:
            program_location_suggestion = program_location
        program_location = filedialog.askdirectory(initialdir=program_location_suggestion, mustexist=True)
    
    def exit_program(self):
        if self.ask_save_changes():
            self.root.quit()
            
    def edit_raw_data(self):
        file = filedialog.askopenfilename(defaultextension=".czf", filetypes=[("CZF-bestanden", "*.czf")])
        if file:
            subprocess.Popen(f"notepad.exe {file}", shell=True)
    
    def manage_first_controllers(self):
        pass
    
    def manage_second_controllers(self):
        pass
    
    def manage_cleaners(self):
        pass
    
    def manage_packagers(self):
        pass
    
    def manual_shutdown(self):
        confirm = messagebox.askyesno(program_name, "Weet je zeker dat je wilt afsluiten?\n\nAlle niet-opgeslagen wijzigingen gaan verloren! Sla alles op voordat je verder gaat.")
        if confirm:
            subprocess.Popen("shutdown -s -t 0 -f", shell=True)
        
    def manual_restart(self):
        confirm = messagebox.askyesno(program_name, "Weet je zeker dat je opnieuw wilt opstarten?\n\nAlle niet-opgeslagen wijzigingen gaan verloren! Sla alles op voordat je verder gaat.")
        if confirm:
            subprocess.Popen("shutdown -r -t 0 -f", shell=True)
        
    def firmware(self):
        confirm = messagebox.askyesno(program_name, "Weet je zeker dat je wilt afsluiten?\n\nAlle niet-opgeslagen wijzigingen gaan verloren! Sla alles op voordat je verder gaat.\n\nDe volgende keer dat je de computer opnieuw opstart, zal de BIOS automatisch naar de BIOS setup gaan.")
        if confirm:
            subprocess.Popen("shutdown -s -fw -t 0 -f", shell=True)
        
    def winre(self):
        confirm = messagebox.askyesno(program_name, "Weet je zeker dat je opnieuw wilt opstarten?\n\nAlle niet-opgeslagen wijzigingen gaan verloren! Sla alles op voordat je verder gaat.\n\nWindows zal opnieuw opstarten naar de herstelomgeving.")
        if confirm:
            subprocess.Popen("shutdown -r -o -t 0 -f", shell=True)

    def open_cmd(self):
        subprocess.Popen("cmd")
        
    def open_regedit(self):
        subprocess.Popen("regedit", shell=True)
        
    def open_self(self):
        subprocess.Popen(f"explorer.exe {os.path.dirname(os.path.abspath(__file__))}", shell=True)

    def open_windows_activation(self):
        subprocess.Popen("start ms-settings:activation", shell=True)
        
    def show_system_overview(self):
        subprocess.Popen("slmgr /dli", shell=True)

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
        subprocess.Popen("plugins/tt.exe")
    
    def install_ms_office(self):
        subprocess.Popen("thirdparty/o64.exe")
        
    def office_shortcuts_programdata(self):
        try:
            subprocess.Popen("""xcopy "C:/ProgramData/Microsoft/Windows/Start Menu/Programs/Word.lnk" %USERPROFILE%/Desktop""", shell=True)
            subprocess.Popen("""xcopy "C:/ProgramData/Microsoft/Windows/Start Menu/Programs/Excel.lnk" %USERPROFILE%/Desktop""", shell=True)
            subprocess.Popen("""xcopy "C:/ProgramData/Microsoft/Windows/Start Menu/Programs/Powerpoint.lnk" %USERPROFILE%/Desktop""", shell=True)
            subprocess.Popen("""xcopy "C:/ProgramData/Microsoft/Windows/Start Menu/Programs/Outlook (Classic).lnk" %USERPROFILE%/Desktop""", shell=True)
            messagebox.showinfo(program_name, "Office snelkoppelingen aangemaakt!")
        except Exception as e:
            messagebox.showerror(program_name, f"Kon Office snelkoppelingen niet maken: {e}")
            
    def office_shortcuts_d_drive(self):
        try:
            subprocess.Popen(f"""xcopy "{program_location}/shrtcts/Word.lnk" %USERPROFILE%/Desktop""", shell=True)
            subprocess.Popen(f"""xcopy "{program_location}/shrtcts/Excel.lnk" %USERPROFILE%/Desktop""", shell=True)
            subprocess.Popen(f"""xcopy "{program_location}/shrtcts/Powerpoint.lnk" %USERPROFILE%/Desktop""", shell=True)
            subprocess.Popen(f"""xcopy "{program_location}/shrtcts/Outlook (Classic).lnk" %USERPROFILE%/Desktop""", shell=True)
            messagebox.showinfo(program_name, "Office snelkoppelingen aangemaakt!")
        except Exception as e:
            messagebox.showerror(program_name, f"Kon Office snelkoppelingen niet maken: {e}")
            
    def office_shortcuts_custom(self):
        try:
            custom_shortcut_copy_location = filedialog.askopenfilenames(defaultextension="*.lnk", initialdir="C:/ProgramData/Microsoft/Windows/Start Menu/Programs", filetypes=[("Snelkoppelingen", "*.lnk"), ("Toepassingen", "*.exe")])
            desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
            for file in custom_shortcut_copy_location:
                subprocess.Popen(f"""xcopy {file} %USERPROFILE%/Desktop""", shell=True)
            messagebox.showinfo(program_name, "Office snelkoppelingen aangemaakt!")
        except Exception as e:
            messagebox.showerror(program_name, f"Kon Office snelkoppelingen niet maken: {e}")

    def office_web_activation(self):
        subprocess.Popen("start msedge --inprivate digital-products.eu/cz", shell=True)
        
    def office_adb_activation(self):
        prepared_wireless_debugging = messagebox.askyesnocancel(program_name, "Deze methode van Office activeren maakt gebruik van draadloze USB-foutopsporing op Android telefoons. Is je telefoon opgezet voor USB-foutopsporing?\n\nLET OP: Deze methode is niet beschikbaar voor Apple-producten!")
        if prepared_wireless_debugging:
            subprocess.Popen(f"{program_location}/plugins/adb.exe devices", shell=True)
            subprocess.Popen(f"{program_location}/plugins/adb.exe shell am start -a android.intent.action.CALL -d tel:08000233487", shell=True)
        elif not prepared_wireless_debugging:
            subprocess.Popen("D:/PorkifyNET/CZEdit/plugins/adb-readme.txt", shell=True)
            
    def office_moat_activation(self):
        subprocess.Popen(f"{program_location}/plugins/moat.exe", shell=True)
    
    def install_hp_hotkeys(self):
        subprocess.Popen("thirdparty/hphk.exe")
    
    def disable_conexant_audio(self):
        try:
            # Stop the Conexant service
            subprocess.run(["sc", "stop", "CxUtilSvc"], check=True)
            # Disable the Conexant service
            subprocess.run(["sc", "config", "CxUtilSvc", "start=", "disabled"], check=True)
            print("Conexant service succesvol gestopt.")
        except subprocess.CalledProcessError:
            print("Kon Conexant niet stoppen. Heeft deze computer Conexant?")

    def open_documentation(self):
        webbrowser.open("github.com/PorkifyNET/CZFEdit/blob/main/README.md", 2)

    def send_to_server(self):
        # Opens Edge in InPrivate mode
        subprocess.run("start msedge --inprivate check.computerzaak.nl", shell=True)
    
    def show_about(self):
        messagebox.showinfo(program_name, f"""
        CZFEdit versie {program_version}
        Gemaakt door Porkify.NET
        Voor gebruik bij computerzaak.nl
    """)
    
    def open_porkifynet(self):
        webbrowser.open("https://porkify.net/", 2)
        
    def ask_save_changes(self):
        # Ask the user if they want to save changes before proceeding with a new file or closing
        if self.filename:  # Only ask if there's a file to save changes
            return messagebox.askyesno("Wijzigingen opslaan", "Weet u zeker dat u een nieuwe checklist wilt beginnen? Alle niet-opgeslagen wijzigingen gaan verloren!")
        return True  # If no file is open, allow the operation without saving
    
    # More function placeholders...

# Run the application
root = tk.Tk()

root.resizable(False, False)
root.iconbitmap("czf.ico")

app = LaptopChecklistApp(root)
root.mainloop()