import tkinter as tk

class TouchpadTester:
    def __init__(self, root):
        self.root = root
        self.root.title("Touchpad/Touchscreen Tester")
        self.root.attributes("-fullscreen", True)  # Set to full-screen
        
        # Set up canvas with a white background
        self.canvas = tk.Canvas(self.root, bg="white")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        # Bind mouse and key events
        self.canvas.bind("<B1-Motion>", self.draw)
        self.canvas.bind("<Button-1>", self.start_draw)
        self.root.bind("<Key>", self.change_color)
        self.root.bind("<Up>", self.increase_thickness)
        self.root.bind("<Down>", self.decrease_thickness)
        self.canvas.bind("<MouseWheel>", self.adjust_thickness_scroll)

        # Initial drawing settings
        self.last_x, self.last_y = None, None
        self.current_color = "red"  # Default color
        self.brush_thickness = 30  # Default radius of the brush (diameter is twice this value)

    def start_draw(self, event):
        # Initialize the last position for the drawing
        self.last_x, self.last_y = event.x, event.y

    def draw(self, event):
        # Draw a circle (brush) at the new position
        if self.last_x is not None and self.last_y is not None:
            radius = self.brush_thickness / 2
            self.canvas.create_oval(
                event.x - radius, event.y - radius, event.x + radius, event.y + radius,
                fill=self.current_color, outline=self.current_color
            )
        
        # Update the last position
        self.last_x, self.last_y = event.x, event.y

    def change_color(self, event):
        # Change color or close based on key press
        color_map = {
            '1': "red",
            '2': "blue",
            '3': "green",
            '4': "white",  # Eraser
            '5': "black"
        }
        
        if event.char in color_map:
            self.current_color = color_map[event.char]
        else:
            self.root.destroy()  # Close the application on other key presses

    def increase_thickness(self, event):
        # Increase brush thickness (diameter)
        self.brush_thickness = min(self.brush_thickness + 5, 100)

    def decrease_thickness(self, event):
        # Decrease brush thickness (diameter)
        self.brush_thickness = max(self.brush_thickness - 5, 10)

    def adjust_thickness_scroll(self, event):
        # Adjust thickness with mouse scroll
        if event.delta > 0:
            self.increase_thickness(event)
        else:
            self.decrease_thickness(event)

# Initialize Tkinter window
root = tk.Tk()
app = TouchpadTester(root)
root.mainloop()
