import tkinter as tk

class KeyboardTester:
    def __init__(self, root):
        self.root = root
        self.root.title("Keyboard Tester")
        self.root.geometry("800x600")  # Bigger window to accommodate all keys
        
        # Create a canvas to visualize the keys being pressed
        self.canvas = tk.Canvas(self.root, bg="white", width=800, height=600)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        # Define the keys to be displayed on the canvas
        self.key_labels = [
            "1", "2", "3", "4", "5", "6", "7", "8", "9", "0",
            "Q", "W", "E", "R", "T", "Y", "U", "I", "O", "P",
            "A", "S", "D", "F", "G", "H", "J", "K", "L",
            "Z", "X", "C", "V", "B", "N", "M",
            "Space", "Enter", "Shift", "Ctrl", "Alt", "Esc",
            "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12",
            "Up", "Down", "Left", "Right", "Backspace", "Tab", "Caps Lock", "Num Lock"
        ]
        
        # For key positions
        self.key_positions = {}
        self.create_keyboard()
        
        # Bind key events globally to root window
        self.root.bind("<KeyPress>", self.key_press)
    
    def create_keyboard(self):
        # Position each key on the canvas
        x, y = 20, 40
        key_width, key_height = 60, 60
        for label in self.key_labels:
            self.key_positions[label] = (x, y)
            self.canvas.create_rectangle(x, y, x + key_width, y + key_height, fill="lightgray", outline="black")
            self.canvas.create_text(x + key_width / 2, y + key_height / 2, text=label, font=("Arial", 12))
            x += key_width + 10
            if x > 800 - key_width:
                x = 20
                y += key_height + 10

    def key_press(self, event):
        key_name = event.keysym if event.keysym else event.char
        
        if not key_name:  # Fallback if no keysym or char detected
            return
        
        key_name = key_name.upper()  # Make sure to check the key in uppercase for consistency

        if key_name in self.key_positions:
            # Get the position of the key pressed and change color
            x, y = self.key_positions[key_name]
            self.canvas.create_rectangle(x, y, x + 60, y + 60, fill="yellow", outline="black")
            self.canvas.create_text(x + 30, y + 30, text=key_name, font=("Arial", 12))
    
    def key_release(self, event):
        key_name = event.keysym if event.keysym else event.char
        
        if not key_name:  # Fallback if no keysym or char detected
            return
        
        key_name = key_name.upper()  # Make sure to check the key in uppercase for consistency

        if key_name in self.key_positions:
            # Reset the color to default when the key is released
            x, y = self.key_positions[key_name]
            self.canvas.create_rectangle(x, y, x + 60, y + 60, fill="lightgray", outline="black")
            self.canvas.create_text(x + 30, y + 30, text=key_name, font=("Arial", 12))

# Initialize the main window
root = tk.Tk()
app = KeyboardTester(root)
root.mainloop()
