import customtkinter as ctk
from app.ui import AppUI
from app.file_manager import FileManager

def main():
    root = ctk.CTk()
    file_manager = FileManager()
    app = AppUI(root, file_manager)
    root.mainloop()

if __name__ == "__main__":
    main()
