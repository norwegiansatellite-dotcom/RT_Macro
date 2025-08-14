import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from typing import Any

from config import EXCEL_FORMAT, RESULT_FILE_NAME
from utils import (get_headers_and_row, get_result_data, get_result_excel,
                   get_sheet_excel)


WINDOW = tk.Tk()


def get_path_to_excel() -> str:
    """Получение пути до Excel файла.

    Returns:
        str: Полный путь до Excel файла.ч
    """
    return filedialog.askopenfilename(
        title="Выберите Excel-файл",
        filetypes=[("Excel файлы", EXCEL_FORMAT)]
    )



def get_user_header(headers_table: list[Any]) -> str | None:
    """Открывает модальное окно с Listbox и возвращает наименование столбца или None.

    Args:
        headers_table (list[Any]): Наименования столбцов в Excel файле.

    Returns:
        str | None: Наименование столбца, иначе None.
    """
    win = tk.Toplevel(WINDOW)
    win.title("Выбор столбца")
    win.transient()
    win.grab_set()

    lbl = tk.Label(win, text="Выберите столбец из списка, по которому нужно отфильтровать данные:")
    lbl.pack(padx=12, pady=(12, 6), anchor="w")

    frame = tk.Frame(win)
    frame.pack(padx=12, pady=6, fill="both", expand=True)

    scrollbar = tk.Scrollbar(frame, orient="vertical")
    listbox = tk.Listbox(frame, selectmode=tk.SINGLE, height=10, yscrollcommand=scrollbar.set)
    scrollbar.config(command=listbox.yview)
    scrollbar.pack(side="right", fill="y")
    listbox.pack(side="left", fill="both", expand=True)

    for header in headers_table:
        listbox.insert(tk.END, header)

    btns = tk.Frame(win)
    btns.pack(padx=12, pady=(6, 12), anchor="e")

    result = {"value": None}

    def on_ok():
        if not (sel := listbox.curselection()):
            messagebox.showwarning("Нет выбора", "Сначала выберите столбец.")
            return
        result["value"] = listbox.get(sel[0])
        win.destroy()

    def on_cancel():
        result["value"] = None
        win.destroy()

    ok_btn = tk.Button(btns, text="OK", width=10, command=on_ok)
    cancel_btn = tk.Button(btns, text="Отмена", width=10, command=on_cancel)
    ok_btn.pack(side="left", padx=(0, 6))
    cancel_btn.pack(side="left")

    win.update_idletasks()
    listbox.focus_set()
    try:
        px = WINDOW.winfo_rootx()
        py = WINDOW.winfo_rooty()
        pw = WINDOW.winfo_width()
        ph = WINDOW.winfo_height()
        ww = win.winfo_width()
        wh = win.winfo_height()
        x = px + (pw - ww) // 2
        y = py + (ph - wh) // 2
        win.geometry(f"+{x}+{y}")
    except Exception:
        pass

    WINDOW.wait_window(win)
    return result["value"]


def process_filtration_excel_file() -> None:
    """Запуск основного процесса филтрации файла"""
    if not (path_to_excel := get_path_to_excel()):
        return

    sheet = get_sheet_excel(path_to_excel)
    if isinstance(sheet, str):
        messagebox.showerror("Ошибка", f"Ошибка при открытии файла: {sheet}")
        return

    datas_headers_and_row = get_headers_and_row(sheet)

    try:
        headers_table, headers_row = datas_headers_and_row[0], datas_headers_and_row[1]
    except TypeError:
        messagebox.showerror("Ошибка", "Данный Excel файл не подходит для этой утилиты!")
        return

    if not headers_table:
        messagebox.showerror("Ошибка", "Не удалось получить заголовки столбцов из файла.")
        return

    if not (user_header := get_user_header(headers_table)):
        return

    user_filter = simpledialog.askstring(
        "Фильтрация",
        f"Введите значение для «{user_header}»:",
        parent=WINDOW
    )

    if not user_header:
        messagebox.showwarning("Нет выбора", "Вы не выбрали столбец для фильтрации.")
        return
    if not user_filter:
        messagebox.showwarning("Нет значения", "Вы не ввели значение для фильтрации.")
        return

    result_data = get_result_data(
        word_filter=user_filter,
        sheet=sheet,
        headers_row=headers_row,
        user_header=user_header,
        headers_table=headers_table
    )

    if not result_data:
        messagebox.showerror("Нет данных", f"Не смог найти данные \"{user_filter}\" в столбце \"{user_header}\"")
        return

    save_path = filedialog.asksaveasfilename(
        title="Сохранить как...",
        defaultextension=".xlsx",
        filetypes=[("Excel файлы", "*.xlsx")],
        initialfile=RESULT_FILE_NAME.format(user_header)
    )

    if not save_path:
        return

    if not save_path.lower().endswith(".xlsx"):
        save_path += ".xlsx"

    try:
        result_excel = get_result_excel(result_data=result_data)
        result_excel.save(save_path)
        messagebox.showinfo("Готово", f"Файл сохранён:\n{save_path}")
    except Exception as ex:
        messagebox.showerror("Ошибка", f"Ошибка при сохранении файла:\n{ex}")


def start_user_window() -> None:
    """Основной процесс запуска GUI."""
    WINDOW.title("Фильтрация столбца в Excel")

    window_width = 400
    window_height = 150

    screen_width = WINDOW.winfo_screenwidth()
    screen_height = WINDOW.winfo_screenheight()

    x = int((screen_width / 2) - (window_width / 2))
    y = int((screen_height / 2) - (window_height / 2))

    WINDOW.geometry(f"{window_width}x{window_height}+{x}+{y}")

    label = tk.Label(WINDOW, text="Выберите Excel файл", font=("Arial", 12))
    label.pack(pady=20)

    button = tk.Button(WINDOW, text="Выбрать файл", command=process_filtration_excel_file, bg="lightblue")
    button.pack(pady=10)

    WINDOW.mainloop()
