import tkinter as tk
import time

# Create the main window
root = tk.Tk()
root.title("Clock")

# Create a label to display the time
time_label = tk.Label(root, font=("Helvetica", 50))
time_label.pack()

# Create a function to update the time
def update_time():
    current_time = time.strftime("%H:%M:%S")
    time_label.config(text=current_time)
    time_label.after(1000, update_time)

# Call the update function to start the clock
update_time()

# Run the Tkinter event loop
root.mainloop()