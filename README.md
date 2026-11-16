# pcb-auto-excel

Программа с графическим интерфейсом (Windows) для автоматического заполнения пустых строк в накладных. 
Сводит данные из "Посылки", "Сводки ПП (Голова)" и "Сводки Филиала", находя совпадения по номеру заказа (PS). Автоматически проставляет количество (По счету) и тип доставки (ДМ, К, СДЭК, Самовывоз).

## Инструкция по сборке .exe
1. Установите Python и зависимости: `pip install -r requirements.txt`
2. Установите PyInstaller: `pip install pyinstaller`
3. Соберите программу: `pyinstaller --noconfirm --onedir --windowed --add-data "venv/Lib/site-packages/customtkinter;customtkinter/" main.py`
