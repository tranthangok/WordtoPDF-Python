import tkinter as tk
from convertfileprogress import FileConverter

if __name__ == "__main__":
    root = tk.Tk()
    app = FileConverter(root)
    root.mainloop()
