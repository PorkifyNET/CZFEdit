import ctypes
import datetime
from datetime import datetime
from math import ceil
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import webbrowser
import psutil
import subprocess
import win32com.client

def run_as_admin():
    """
    Checks if the script is running as an administrator. If not, restarts the program with admin privileges.
    """
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
    except:
        is_admin = False

    if not is_admin:
        # Re-launch the program with elevated privileges
        try:
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, " ".join(sys.argv), None, 1
            )
        except Exception as e:
            print(f"Error: {e}")
        sys.exit()  # Exit the current program since it will be restarted as admin

# Call this function at the start of your program
if __name__ == "__main__":
    #run_as_admin()

    # Your program's main logic starts here
    print("Program is running with administrator privileges.")

program_name = "CZFEdit"
program_version = "v1.2"
global program_location
program_location = os.path.dirname(__file__) if os.path.exists(os.path.dirname(__file__)) else filedialog.askdirectory(initialdir=__file__, mustexist=True, title="Waar ben ik geinstalleerd?")

if not os.path.exists("%TEMP%/CZFEdit"):
    os.mkdir(f"{os.environ.get("TEMP")}/CZFEdit")
    print(f"No temporary folder was found, creating a new one in {os.environ.get("TEMP")}/CZFEdit...")

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

# Define tooltip class
class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None

        # Bind events to show and hide the tooltip
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        # Create a tooltip window
        if self.tooltip_window is not None:
            return

        x, y, _, _ = self.widget.bbox("insert")  # Get widget bounds
        x += self.widget.winfo_rootx() + 25      # Position tooltip slightly offset
        y += self.widget.winfo_rooty() + 25

        # Create the tooltip window
        self.tooltip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)  # Remove window decorations
        tw.wm_geometry(f"+{x}+{y}")

        # Create the label inside the tooltip window
        label = tk.Label(tw, text=self.text, background="white", relief="solid", borderwidth=1, padx=5, pady=2)
        label.pack()

    def hide_tooltip(self, event=None):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

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

        # System information retrieval functions
    def get_ram(self):
        try:
            return str(f"{(ceil(int(psutil.virtual_memory().total / (1024 ** 3)))) + 1 } GB")
        except:
            return "Onbekend"

    def get_storage(self):
        try:
            return str(f"{ceil(int(psutil.disk_usage('C:').total / (1024 ** 3)))} GB ({snap_to_nearest_power_of_2(ceil(int(psutil.disk_usage('C:').total / (1024 ** 3))))})")
        except:
            return "Onbekend"

    def get_cpu(self):
        try:
            return str(subprocess.check_output("wmic cpu get name").decode().split('\n')[1].strip())
        except:
            return "Onbekend"
    
    def get_battery_health(self):
        battery = psutil.sensors_battery()
        if battery is None:
            return "Geen batterij!"
        else:
            # Run BatteryInfoView with /stext to save output as plain text
            try:
                print(f"{program_location}/plugins/BatteryInfoView.exe")
                subprocess.run([f"{program_location}/plugins/BatteryInfoView.exe", "/scomma", "battery_info.txt"])
            except Exception as e:
                print(e)

            try:
                with open("battery_info.txt", "r") as file:
                    lines = file.readlines()

                    # Parse each line to find Design Capacity and Full Charge Capacity
                    battery_health_var = lines[10][15:].strip()
                    return battery_health_var
                
            except:
                return "Leesfout"

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
            messagebox.showerror(program_name, e)
            return "Onbekend"

    def check_hp_hotkeys(self):
        if os.path.exists("C:\Program Files (x86)\HP\HP Hotkey Support\HotkeyService.exe"):
            return "Ja"
        else: return "Nee"

    def check_conexant_audio(self):
        if os.path.exists("C:\Program Files\CONEXANT"):
            return "Nee"
        else: return "Ja"
        
    def get_current_date(self):
        return str(f"{datetime.now().strftime("%d/%m/%Y")}")

    def refresh_battery_health_label(self):
        self.battery_health.set(self.get_battery_health())
        
    def refresh_ram_label(self):
        try:
            self.ram.set(self.get_ram())
        except Exception as e:
            messagebox.showerror(program_name, e)
        
    def refresh_storage_label(self):
        try:
            self.storage_space.set(self.get_storage())
        except Exception as e:
            messagebox.showerror(program_name, e)
        
    def refresh_cpu_label(self):
        try:
            self.cpu_name.set(self.get_cpu())
        except Exception as e:
            messagebox.showerror(program_name, e)
        
    def refresh_updates_label(self):
        try:
            self.updates_installed.set(self.check_updates())
        except Exception as e:
            messagebox.showerror(program_name, e)
        
    def refresh_hp_label(self):
        try:
            self.hp_hotkeys_installed.set(self.check_hp_hotkeys())
        except Exception as e:
            messagebox.showerror(program_name, e)
        
    def refresh_conexant_label(self):
        try:
            self.conexant_installed.set(self.check_conexant_audio())
        except Exception as e:
            messagebox.showerror(program_name, e)
            
    def update_battery_health_style(self, *args):
        value = self.battery_health.get()
        print(f"Battery Health Updated: {value}")  # Debugging output

        if value == "Geen batterij!":
            self.battery_health_label.config(foreground="red")
            Tooltip(self.battery_health_label, "Deze laptop heeft geen accu, de accu is dood, of is niet goed ingeplugt. Check de accu.")
        elif value == "Onbekend":
            self.battery_health_label.config(foreground="red")
            Tooltip(self.battery_health_label, "De gezondheid van de accu kan niet worden bepaald. Dit kan aan de code van dit programma liggen, of BatteryLifeInfo is niet meegeleverd door ons.")
        else:
            self.battery_health_label.config(foreground="black")

    def update_ram_style(self, *args):
        value = self.ram.get()
        print(f"RAM Updated: {value}")  # Debugging output

        if value == "Onbekend":
            self.ram_label.config(foreground="red")
            Tooltip(self.ram_label, "De hoeveelheid RAM kon niet worden bepaald. Dit is hoogstwaarschijnlijk een fout aan onze kant. Probeer erachter te komen met cLauncher of met Taakbeheer.")
        else:
            self.ram_label.config(foreground="black")

    def update_cpu_style(self, *args):
        value = self.cpu_name.get()
        print(f"CPU Updated: {value}")  # Debugging output

        if value == "Onbekend":
            self.cpu_label.config(foreground="red")
            Tooltip(self.cpu_label, "De naam van de CPU kon niet worden gevonden. Dit is hoogstwaarschijnlijk een fout aan onze kant. Probeer erachter te komen met cLauncher of met Taakbeheer.")
        else:
            self.cpu_label.config(foreground="black")

    def update_storage_style(self, *args):
        value = self.storage_space.get()
        print(f"Storage Space Updated: {value}")  # Debugging output

        if value == "Onbekend":
            self.storage_label.config(foreground="red")
            Tooltip(self.storage_label, "De hoeveelheid opslag kon niet worden bepaald. Dit is kan een fout zijn aan onze kant, of je checkt de laptop voor het installeren van Windows. Probeer achter de schijfgrootte te komen met cLauncher of met Taakbeheer.")
        else:
            self.storage_label.config(foreground="black")
            Tooltip(self.storage_label, "LET OP: Dit geeft alleen de grootte van de C-schijf aan. Het kan voorkomen dat er een tweede schijf in de laptop zit, dus controleer dit voor de zekerheid.")

    def update_updates_style(self, *args):
        value = self.updates_installed.get()
        print(f"Windows Updates Status Updated: {value}")  # Debugging output

        if value == "Onbekend":
            self.updates_label.config(foreground="red")
            Tooltip(self.updates_label, "Er is iets misgegaan tijdens het checken voor updates. Dit kan zijn omdat je niet bent verbonden met het internet.")
        elif value == "Windows is bijgewerkt!":
            self.updates_label.config(foreground="black")
        else:
            self.updates_label.config(foreground="red")
            Tooltip(self.updates_label, "Er staan updates klaar. Deze moeten worden gedaan voordat je de controle kan afronden.")

    def update_hotkeys_style(self, *args):
        value = self.hp_hotkeys_installed.get()
        print(f"Hotkeys Updated: {value}")  # Debugging output

        if value == "Nee":
            self.hp_hotkeys_label.config(foreground="orange")
            Tooltip(self.hp_hotkeys_label, "Hotkeys zijn nodig op HP laptops. Is dit geen HP laptop? Dan kun je dit negeren!")
        else:
            self.hp_hotkeys_label.config(foreground="black")
            Tooltip(self.hp_hotkeys_label, "Als dit een HP laptop is is het nodig om HP Hotkeys te installeren. De installatie kan gevonden worden onder Controle > HP > Instaleer Hotkeys...")

    def update_conexant_style(self, *args):
        value = self.conexant_installed.get()
        print(f"Conexant Updated: {value}")  # Debugging output

        if value == "Nee":
            self.conexant_audio_label.config(foreground="orange")
            Tooltip(self.conexant_audio_label, "De Conexant audio driver kan wel eens geinstalleerd worden terwijl er geen Conexant-speakers in de laptop zitten. Dit veroorzaakt veel irritante pop-ups. Als dit niet het geval is kun je dit negeren.")
        else:
            self.conexant_audio_label.config(foreground="black")
        
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
        self.battery_health_label = ttk.Label(self.root, textvariable=self.battery_health)
        self.battery_health_label.grid(row=5, column=1, padx=5, pady=5)
        self.battery_health.trace_add("write", self.update_battery_health_style)

        # RAM in GB
        ttk.Label(self.root, text="RAM (GB):").grid(row=6, column=0, padx=5, pady=5)
        self.ram_label = ttk.Label(self.root, textvariable=self.ram)
        self.ram_label.grid(row=6, column=1, padx=5, pady=5)
        self.ram.trace_add("write", self.update_ram_style)

        # Storage space of C: drive
        ttk.Label(self.root, text="Opslagruimte C: (GB):").grid(row=7, column=0, padx=5, pady=5)
        self.storage_label = ttk.Label(self.root, textvariable=self.storage_space)
        self.storage_label.grid(row=7, column=1, padx=5, pady=5)
        self.storage_space.trace_add("write", self.update_storage_style)

        # CPU Name
        ttk.Label(self.root, text="CPU Naam:").grid(row=8, column=0, padx=5, pady=5)
        self.cpu_label = ttk.Label(self.root, textvariable=self.cpu_name)
        self.cpu_label.grid(row=8, column=1, padx=5, pady=5)
        self.cpu_name.trace_add("write", self.update_cpu_style)

        # Are updates installed?
        ttk.Label(self.root, text="Updates geinstalleerd?").grid(row=9, column=0, padx=5, pady=5)
        self.updates_label = ttk.Label(self.root, textvariable=self.updates_installed)
        self.updates_label.grid(row=9, column=1, padx=5, pady=5)
        self.updates_installed.trace_add("write", self.update_updates_style)

        # Is HP Hotkeys installed?
        ttk.Label(self.root, text="HP Hotkeys geinstalleerd?").grid(row=10, column=0, padx=5, pady=5)
        self.hp_hotkeys_label = ttk.Label(self.root, textvariable=self.hp_hotkeys_installed)
        self.hp_hotkeys_label.grid(row=10, column=1, padx=5, pady=5)
        self.hp_hotkeys_installed.trace_add("write", self.update_hotkeys_style)

        # Is Conexant Audio disabled or nonexistent?
        ttk.Label(self.root, text="Conexant Audio gedeactiveerd of niet aanwezig?").grid(row=11, column=0, padx=5, pady=5)
        self.conexant_audio_label = ttk.Label(self.root, textvariable=self.conexant_installed)
        self.conexant_audio_label.grid(row=11, column=1, padx=5, pady=5)
        self.conexant_installed.trace_add("write", self.update_conexant_style)

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
        
        export_menu = tk.Menu(file_menu, tearoff=0)
        export_menu.add_command(label="Inventaris-bestand", command=self.export_file)
        export_menu.add_command(label="Sales-pitch", command=self.export_sales_pitch)
        file_menu.add_cascade(label="Exporteren als...", menu=export_menu)

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
        fujitsu_checks.add_separator()
        fujitsu_checks.add_command(label="Toon backup BIOS wachtwoorden", command=self.show_fujitsu_backup_bios_pw)
        fujitsu_checks.add_command(label="Open BiosPW", command=self.open_biospw)
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
        #menu_bar.add_cascade(label="Vernieuwen", menu=refresh_menu)

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
            subprocess.Popen("""powershell Set-LocalUser -Name ^"Gebruiker^" -PasswordNeverExpires 1""", shell=True)
            subprocess.Popen("""powershell Set-LocalUser -Name ^"Gebruiker^" -PasswordNeverExpires 1""", shell=True)
            messagebox.showinfo(program_name, "Leeftijdslimiet wachtwoord is succesvol verwijderd.")
        except Exception as e:
            messagebox.showerror(program_name, e)
            
    def disable_fast_startup(self):
        try:
            subprocess.Popen("Powercfg -h off", shell=True)
            messagebox.showinfo(program_name, "Snel opstarten is succesvol uitgeschakeld.")
        except Exception as e:
            messagebox.showerror(program_name, e)

    def disable_fujitsu_battery_charging_tool(self):
        try:
            subprocess.Popen("taskkill /f /im BatteryCtrlUpdate.exe", shell=True)
            messagebox.showinfo(program_name, "Fujitsu Battery Charging Tool is succesvol uitgeschakeld.")
        except Exception as e:
            messagebox.showerror(program_name, e)

    def show_fujitsu_backup_bios_pw(self):
        messagebox.showinfo(program_name, """3hqgo3
                            jqw534
                            0qww294e""")
        
    def open_biospw(self):
        subprocess.run("start msedge --inprivate bios-pw.org", shell=True)

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
            print(f"NEW FILE LOADED: {self.filename}")
            for line in lines:
                if line.startswith("ORDER"):
                    self.order.delete(0, tk.END)
                    self.order.insert(0, line.split("ORDER")[1].strip())
                    print(f"Found ORDER {line.split("ORDER")[1].strip()}")
                elif line.startswith("CZNUM"):
                    self.cz.delete(0, tk.END)
                    self.cz.insert(0, line.split("CZNUM")[1].strip())
                    print(f"Found CZNUM {line.split("CZNUM")[1].strip()}")
                elif line.startswith("ORDDAT"):
                    self.date.delete(0, tk.END)
                    self.date.insert(0, line.split("ORDDAT")[1].strip())
                    print(f"Found ORDDAT {line.split("ORDDAT")[1].strip()}")
                elif line.startswith("BRAND"):
                    self.brand.set(line.split("BRAND")[1].strip())
                    print(f"Found BRAND {line.split("BRAND")[1].strip()}")
                elif line.startswith("WINVER"):
                    self.windows_version.set(line.split("WINVER")[1].strip())
                    print(f"Found WINVER {line.split("WINVER")[1].strip()}")
                elif line.startswith("BATQLTY"):
                    self.battery_health_label.config(text=line.split("BATQLTY")[1].strip())
                    self.battery_health.set(line.split("BATQLTY")[1].strip())
                    print(f"Found BATQLTY {line.split("BATQLTY")[1].strip()}")
                elif line.startswith("RAM"):
                    self.ram_label.config(text=line.split("RAM")[1].strip())
                    self.ram.set(line.split("RAM")[1].strip())
                    print(f"Found RAM {line.split("RAM")[1].strip()}")
                elif line.startswith("STRSPC"):
                    self.storage_label.config(text=line.split("STRSPC")[1].strip())
                    self.storage_space.set(line.split("STRSPC")[1].strip())
                    print(f"Found STRSPC {line.split("STRSPC")[1].strip()}")
                elif line.startswith("CPU"):
                    self.cpu_label.config(text=line.split("CPU")[1].strip())
                    self.cpu_name.set(line.split("CPU")[1].strip())
                    print(f"Found CPU {line.split("CPU")[1].strip()}")
                elif line.startswith("WINUTD"):
                    self.updates_label.config(text=line.split("WINUTD")[1].strip())
                    self.updates_installed.set(line.split("WINUTD")[1].strip())
                    print(f"Found WINUTD {line.split("WINUTD")[1].strip()}")
                elif line.startswith("HPHKINST"):
                    self.hp_hotkeys_label.config(text=line.split("HPHKINST")[1].strip())
                    self.hp_hotkeys_installed.set(line.split("HPHKINST")[1].strip())
                    print(f"Found HPHKINST {line.split("HPHKINST")[1].strip()}")
                elif line.startswith("CONEXANT"):
                    self.conexant_audio_label.config(text=line.split("CONEXANT")[1].strip())
                    self.conexant_installed.set(line.split("CONEXANT")[1].strip())
                    print(f"Found CONEXANT {line.split("CONEXANT")[1].strip()}")
                elif line.startswith("DRVUTD"):
                    self.drivers_check.set(line.split("DRVUTD")[1].strip())
                    print(f"Found DRVUTD {line.split("DRVUTD")[1].strip()}")
                elif line.startswith("POSAUD"):
                    self.audio_check.set(line.split("POSAUD")[1].strip())
                    print(f"Found POSAUD {line.split("POSAUD")[1].strip()}")
                elif line.startswith("POSKEY"):
                    self.keyboard_check.set(line.split("POSKEY")[1].strip())
                    print(f"Found POSKEY {line.split("POSKEY")[1].strip()}")
                elif line.startswith("POSTCH"):
                    self.touch_check.set(line.split("POSTCH")[1].strip())
                    print(f"Found POSTCH {line.split("POSTCH")[1].strip()}")
                elif line.startswith("POSCAM"):
                    self.camera_check.set(line.split("POSCAM")[1].strip())
                    print(f"Found POSCAM {line.split("POSCAM")[1].strip()}")
                elif line.startswith("OFFACT"):
                    self.office_check.set(line.split("OFFACT")[1].strip())
                    print(f"Found OFFACT {line.split("OFFACT")[1].strip()}")
                elif line.startswith("CTRLNAME"):
                    self.executor_name.delete(0, tk.END)
                    self.executor_name.insert(0, line.split("CTRLNAME")[1].strip())
                    print(f"Found CTRLNAME {line.split("CTRLNAME")[1].strip()}")
                elif line.startswith("CTRLDATE"):
                    self.execution_date_label.config(text=line.split("CTRLDATE")[1].strip(),foreground="black")
                    print(f"Found CTRLDATE {line.split("CTRLDATE")[1].strip()}")
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

    def export_file(self):
        file = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Tekstbestand", "*.txt")], initialfile=self.order.get())
        if file:
            self.filename = file
            self.export_file_as_inventory_file(self, file)

    def export_file_as_inventory_file(self, filename):
        # Format the data and save it into the .CZF file
        with open(filename, 'w') as file:
            file.write("+=====[ Specificaties Winkelvoorraad ]=====+\n")
            file.write(f"Ordernummer: {self.order.get()}\n")
            file.write(f"CZ-nummer: {self.cz.get()}\n")
            file.write(f"Orderdatum: {self.date.get()}\n")
            file.write(f"Merk laptop: {self.brand.get()}\n")
            file.write(f"Windows Versie: {self.windows_version.get()}\n")
            file.write(f"Kwaliteit accu: {self.get_battery_health()}\n")
            file.write(f"RAM: {self.get_ram()}\n")
            file.write(f"Opslag: {self.get_storage()}\n")
            file.write(f"CPU: {self.get_cpu()}\n")
            
            if self.check_updates() == "Windows is bijgewerkt!":
                file.write(f"Windows is geupdated!\n")
            else:
                file.write(f"Windows mist nog updates!\n")
                
            if self.check_hp_hotkeys() == True and self.brand.get() == "HP":
                file.write(f"Hotkeys zijn geinstalleerd!\n")
            elif self.check_hp_hotkeys() == False and self.brand.get() == "HP":
                file.write(f"Hotkeys zijn niet geinstalleerd!\n")
            else:
                file.write(f"Hotkeys zijn niet nodig, gezien dit geen HP is!\n")
                
            if self.conexant_installed.get() == "Ja":
                file.write(f"Conexant is geinstalleerd!\n")
            else:
                file.write(f"Conexant is niet geinstalleerd!\n")
                
            if self.drivers_check.get() == True:
                file.write(f"Drivers zijn up-to-date!\n")
            else:
                file.write(f"Drivers zijn niet up-to-date!\n")
                
            if self.audio_check.get() == True:
                file.write(f"Audio is gecheckt en werkt!\n")
            else:
                file.write(f"Audio is niet gecheckt!\n")
                
            if self.keyboard_check.get() == True:
                file.write(f"Toetsenbord is gecheckt en werkt volledig!\n")
            else:
                file.write(f"Toetsenbord is niet gecheckt!\n")
                
            if self.touch_check.get() == True:
                file.write(f"Touchscreen is gecheckt en werkt!\n")
            else:
                file.write(f"Touchscreen is niet gecheckt, of deze laptop heeft geen touchscreen!\n")
                
            if self.camera_check.get() == True:
                file.write(f"Camera is gecheckt en werkt!\n")
            else:
                file.write(f"Camera is niet gecheckt, of deze laptop heeft geen camera!\n")
                
            if self.office_check.get() == True:
                file.write(f"Office is geinstalleerd en geactiveerd!\n")
            else:
                file.write(f"Office is niet geinstalleerd, geactiveerd, of was niet nodig!\n")
            file.write(f"Gecontroleerd door {self.executor_name.get()} op {self.get_current_date()}\n")
            file.write("+================[ CZEdit ]================+\n")
        messagebox.showinfo("Geexporteerd!", f"Winkelvoorraad-checklist opgeslagen als {filename}")

        self.battery_health.trace_add("write", self.update_battery_health_style)

    def export_sales_pitch(self):
        file = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Tekstbestand", "*.txt")], initialfile="Welkom bij uw nieuwe laptop!")
        if file:
            self.filename = file
            self.export_file_as_sales_pitch(file)

    def export_file_as_sales_pitch(self, filename):
    # Format the data into a sales pitch and save it into the .txt file
        with open(filename, 'w') as file:
            file.write("+=====[ Laptop Te Koop! ]=====+\n")
            file.write(f"Ben je op zoek naar een betrouwbare laptop? Zoek niet verder! Deze {self.brand.get()} laptop heeft alles wat je nodig hebt:\n\n")
            file.write(f"\u2022 **Ordernummer**: {self.order.get()}\n")
            file.write(f"\u2022 **Windows Versie**: {self.windows_version.get()} - Geniet van de nieuwste functies en beveiliging!\n")
            file.write(f"\u2022 **Accu Kwaliteit**: {self.get_battery_health()} - Perfect voor urenlang gebruik zonder opladen.\n")
            file.write(f"\u2022 **RAM**: {self.get_ram()} - Multitasken gaat soepel en snel!\n")
            file.write(f"\u2022 **Opslagruimte**: {self.get_storage()} - Genoeg ruimte voor al je bestanden, games en meer.\n")
            file.write(f"\u2022 **Processor**: {self.get_cpu()} - Geniet van snelheid en prestaties op topniveau.\n\n")

            if self.check_updates() == "Windows is bijgewerkt!":
                file.write("Deze laptop is volledig bijgewerkt met de nieuwste Windows-updates.\n")
            else:
                file.write("Let op: deze laptop heeft mogelijk nog enkele Windows-updates nodig.\n")

            if self.check_hp_hotkeys() == True and self.brand.get() == "HP":
                file.write("Inclusief functionele HP Hotkeys voor snel en handig gebruik.\n")
            elif self.brand.get() != "HP":
                file.write("De perfecte keuze, zelfs zonder HP Hotkeys - dit model maakt het simpel!\n")

            if self.conexant_installed.get() == "Ja":
                file.write("Uitgerust met Conexant Audio voor een geweldige geluidservaring.\n")
            else:
                file.write("Audio is van hoge kwaliteit, zelfs zonder extra Conexant Audio-software.\n")

            if self.drivers_check.get() == True:
                file.write("Alle stuurprogramma's zijn up-to-date. Plug-and-play werkt naadloos!\n")
            else:
                file.write("Een kleine update kan de ervaring nog beter maken.\n")

            if self.audio_check.get() == True:
                file.write("Audio is getest en werkt perfect - ideaal voor werk en entertainment.\n")
            else:
                file.write("Audio is nog niet getest, maar biedt standaard hoge prestaties.\n")

            if self.keyboard_check.get() == True:
                file.write("Het toetsenbord is volledig getest en werkt foutloos - typen was nog nooit zo fijn!\n")
            else:
                file.write("Het toetsenbord biedt een comfortabele typ-ervaring, getest of niet!\n")

            if self.touch_check.get() == True:
                file.write("Het touchscreen is getest en werkt moeiteloos - perfect voor creatieve projecten.\n")
            else:
                file.write("Geen touchscreen? Geen probleem, het scherm is helder en responsief!\n")

            if self.camera_check.get() == True:
                file.write("De camera is getest en klaar voor videogesprekken en selfies!\n")
            else:
                file.write("Camera niet getest, maar altijd klaar om te presteren.\n")

            if self.office_check.get() == True:
                file.write("Inclusief volledig geinstalleerde en geactiveerde Microsoft Office - klaar voor werk of studie!\n")
            else:
                file.write("Microsoft Office niet nodig? Gebruik de laptop zoals jij wilt.\n")

            file.write(f"\nDit alles wordt aangeboden voor een geweldige prijs. Grijp deze kans en voeg deze krachtige {self.brand.get()} laptop vandaag nog toe aan je leven!\n\n")
            file.write(f"Gecontroleerd door {self.executor_name.get()} op {self.get_current_date()}.\n")
            file.write("+================[ CZFEdit ]================+\n")

        messagebox.showinfo("Verkooppitch Geexporteerd!", f"Verkooppitch opgeslagen als {filename}")



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
            
    def open_batterylifeinfo(self):
        subprocess.Popen(f"{program_location}/plugins/BatteryInfoView.exe", shell=True)
        
    def open_snappy(self):
        subprocess.Popen(f"{program_location}/plugins/Snappy64.exe", shell=True)
    
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
        subprocess.Popen(f"{program_location}/plugins/tt.exe")
    
    def install_ms_office(self):
        subprocess.Popen(f"{program_location}/plugins/o64.exe")
        
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
        subprocess.Popen(f"{program_location}/plugins/hphk.exe")
    
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
        subprocess.run("start msedge --inprivate github.com/PorkifyNET/CZFEdit/blob/main/README.md", shell=True)

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