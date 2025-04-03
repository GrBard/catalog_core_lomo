import os
from setuptools import setup
import PyInstaller.__main__

# Определяем пути к ресурсам
RESOURCES_DIR = os.path.join(os.path.dirname(__file__), 'resources')
ICON_PATH = os.path.join(RESOURCES_DIR, 'my_icon.ico')
SCALE_PATH = os.path.join(RESOURCES_DIR, 'scale.jpg')
SHKALA_PATH = os.path.join(RESOURCES_DIR, 'shkala.jpg')
ARIAL_PATH = os.path.join(RESOURCES_DIR, 'arial.ttf')

# Проверяем, существуют ли файлы
for resource in [ICON_PATH, SCALE_PATH, SHKALA_PATH, ARIAL_PATH]:
    if not os.path.exists(resource):
        raise FileNotFoundError(f"Ресурс не найден: {resource}")

# Настройки PyInstaller
pyinstaller_args = [
    'app/main.py',  # Главный скрипт
    '--name=CoreCatalog',  # Имя приложения
    '--onefile',  # Собрать в один файл
    '--windowed',  # Без консоли (для GUI)
    f'--icon={ICON_PATH}',  # Указываем иконку для .exe
    # Добавляем ресурсы
    f'--add-data={ICON_PATH};resources',
    f'--add-data={SCALE_PATH};resources',
    f'--add-data={SHKALA_PATH};resources',
    f'--add-data={ARIAL_PATH};resources',
    '--clean',  # Очищаем временные файлы перед сборкой
    '--noconfirm',  # Не запрашивать подтверждение
]

# Если не на Windows, заменяем разделитель для --add-data
if os.name != 'nt':
    pyinstaller_args = [arg.replace(';', ':') for arg in pyinstaller_args]

# Запускаем PyInstaller
PyInstaller.__main__.run(pyinstaller_args)

# Настройки для setuptools (опционально, если нужно создать установочный пакет)
setup(
    name="CoreCatalog",
    version="1.0.0",
    description="Приложение для создания каталога фотографий керна",
    author="Grigorii Bardyshev",
    author_email="grbard@yandex.ru",
    packages=['app'],
    install_requires=[
        'customtkinter==5.2.2',
        'pandas==2.2.2',
        'python-docx==1.1.2',
        'Pillow==10.4.0',
        'openpyxl==3.1.5',
        'docx2pdf==0.1.8',
        'pywin32==306',
    ],
    entry_points={
        'gui_scripts': [
            'corecatalog=app.main:main',
        ]
    },
    include_package_data=True,
    package_data={
        'app': ['resources/*'],
    },
)
