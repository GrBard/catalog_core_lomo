import customtkinter as ctk
from app.ui import AppUI
from app.file_manager import FileManager
from app.utils import resource_path  # Импортируем resource_path

def main():
    root = ctk.CTk()

    # Устанавливаем иконку для окна
    try:
        icon_path = resource_path("resources/my_icon.ico")  # Укажи путь к иконке
        print(f"Пытаемся загрузить иконку: {icon_path}")
        root.iconbitmap(icon_path)  # Устанавливаем иконку
    except Exception as e:
        print(f"Ошибка загрузки иконки: {e}")
        # Продолжаем выполнение без иконки

    file_manager = FileManager()
    app = AppUI(root, file_manager)
    root.mainloop()

if __name__ == "__main__":
    main()