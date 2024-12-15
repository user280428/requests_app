import tkinter as tk
from tkinter import filedialog, messagebox
from functions import *
import os


def open_file(entry):
    file_path = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Книга Excel", "*.xlsx")],
                                                        initialdir=os.getcwd())
    entry.delete(0, tk.END)
    entry.insert(0, file_path)

def process_files():
    file1_path = entry_file1.get()
    file2_path = entry_file2.get()

    if not file1_path or not file2_path:
        messagebox.showerror("Ошибка", "Пожалуйста, выберите оба файла.")
        return

    try:
        df_base = initial_df_to_base(file1_path)
        request_numbers_list = request_df_to_request_list(file2_path)
        nums_list = num_finder(df_base, request_numbers_list)
        result_df = request_handler(nums_list, file1_path)

        # Сохраняем результат в файл
        path_2 = file2_path.split('/')[-1]
        result_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Книга Excel", "*.xlsx")],
                                                        initialdir=os.getcwd(), initialfile=f'Ответ на {path_2}')
        if result_file_path:
            result_df.save(result_file_path)
            messagebox.showinfo("Успех", f"Результат сохранен в {result_file_path}")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка при обработке файлов: {e}")

# Создаем главное окно
root = tk.Tk()
root.title("Запросы в базу")

# Создаем и размещаем элементы интерфейса
label_file1 = tk.Label(root, text="Выберите файл Базы:")
label_file1.grid(row=0, column=0, padx=15, pady=15)

entry_file1 = tk.Entry(root, width=40)
entry_file1.grid(row=0, column=1, padx=15, pady=15)
entry_file1.insert(0,f'{os.getcwd()}\\База.xlsx')

button_file1 = tk.Button(root, text="Обзор", command=lambda: open_file(entry_file1))
button_file1.grid(row=0, column=2, padx=15, pady=15)

label_file2 = tk.Label(root, text="Выберите файл Запроса:")
label_file2.grid(row=1, column=0, padx=15, pady=15)

entry_file2 = tk.Entry(root, width=40)
entry_file2.grid(row=1, column=1, padx=15, pady=15)

button_file2 = tk.Button(root, text="Обзор", command=lambda: open_file(entry_file2))
button_file2.grid(row=1, column=2, padx=15, pady=15)

button_process = tk.Button(root, text="Обработать файлы", command=process_files)
button_process.grid(row=2, column=1, padx=15, pady=15)

# Запускаем главный цикл обработки событий
root.mainloop()


# pyinstaller --onefile --windowed --icon=icon.ico --name=Запросы app.py