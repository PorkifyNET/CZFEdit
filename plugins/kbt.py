import tkinter as tk
from tkinter import ttk

class KeyboardWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Keyboard Tester")
        self.geometry("800x400")
        
        # Stack panels (Frames in tkinter)
        self.functions_frame = tk.Frame(self)
        self.functions_frame.pack(side=tk.TOP, fill=tk.X)

        self.numbers_frame = tk.Frame(self)
        self.numbers_frame.pack(side=tk.TOP, fill=tk.X)

        self.qwerty_frame = tk.Frame(self)
        self.qwerty_frame.pack(side=tk.TOP, fill=tk.X)

        # Add keys to each frame
        self.keys = {}
        self.add_function_keys()
        self.add_number_keys()
        self.add_qwerty_keys()

        # Event bindings
        self.bind("<KeyPress>", self.on_key_down)
        self.bind("<KeyRelease>", self.on_key_up)

    def add_function_keys(self):
        for key in ["Esc", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12"]:
            btn = ttk.Button(self.functions_frame, text=key, width=5)
            btn.pack(side=tk.LEFT, padx=2, pady=5)
            self.keys[key] = btn

    def add_number_keys(self):
        for key in ["`", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "-", "=", "Backspace"]:
            btn = ttk.Button(self.numbers_frame, text=key, width=5)
            btn.pack(side=tk.LEFT, padx=2, pady=5)
            self.keys[key] = btn

    def add_qwerty_keys(self):
        for key in ["Tab", "Q", "W", "E", "R", "T", "Y", "U", "I", "O", "P", "[", "]", "\\"]:
            btn = ttk.Button(self.qwerty_frame, text=key, width=5)
            btn.pack(side=tk.LEFT, padx=2, pady=5)
            self.keys[key] = btn

    def on_key_down(self, event):
        key = event.keysym
        if key in self.keys:
            self.keys[key].configure(style="Pressed.TButton")
            print(f"Key pressed: {key}")

    def on_key_up(self, event):
        key = event.keysym
        if key in self.keys:
            self.keys[key].configure(style="TButton")
            print(f"Key released: {key}")

if __name__ == "__main__":
    app = KeyboardWindow()
    app.mainloop()
