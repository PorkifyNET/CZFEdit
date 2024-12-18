import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
import barcode
from barcode.writer import ImageWriter

class AssetWindow(tk.Tk):
    def __init__(self):
        super().__init__()

        # Window configuration
        self.title("Asset Window")
        self.resizable(False, False)

        # Set position to top-right corner
        screen_width = self.winfo_screenwidth()
        self.geometry(f"+{screen_width - 400}+0")

        # Generate and display barcode
        self.img_barcode = ttk.Label(self)
        self.img_barcode.pack()

        self.create_barcode("CC0A")  # Example barcode value

        # Overriding close event
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_barcode(self, data):
        """Generate a barcode image and display it in the window."""
        ean = barcode.get('ean13', data, writer=ImageWriter())
        barcode_path = "barcode"
        ean.save(barcode_path)

        # Load and display barcode image
        image = Image.open(f"{barcode_path}.png")
        barcode_img = ImageTk.PhotoImage(image)
        self.img_barcode.config(image=barcode_img)
        self.img_barcode.image = barcode_img

    def on_closing(self):
        """Override the close button to hide the window instead."""
        self.withdraw()


if __name__ == "__main__":
    app = AssetWindow()
    app.mainloop()
