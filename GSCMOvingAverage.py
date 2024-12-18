import pandas as pd
import matplotlib

matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import tkinter as tk
from tkinter import ttk


file_path = 'E:/WORK/Python/GSCMovingAverage/newestData.xlsx'

try:
    df = pd.read_excel(file_path, sheet_name='Даты')
except Exception as e:
    print(f"Ошибка при чтении файла: {e}")
    exit()


df['Дата'] = pd.to_datetime(df['Дата'], errors='coerce')


columns = [col for col in df.columns if col != 'Дата']

window_sizes = [7, 14, 28, 54]


root = tk.Tk()
root.title('График скользящего среднего')

# Создание словаря для хранения состояния флажков
checkbox_states = {col: tk.BooleanVar(value=True) for col in columns}
checkbox_avg_states = {col: tk.BooleanVar(value=True) for col in columns}


fig, ax = plt.subplots(figsize=(12, 6), dpi=100)
canvas = FigureCanvasTkAgg(fig, master=root)
canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

# Функция для построения графика скользящего среднего
def plot_moving_average(window_size):
    # Очистка текущего графика
    ax.clear()

    
    for col in columns:
        if checkbox_states[col].get():
            ax.plot(df['Дата'], df[col], label=f'{col} (значения)')
        if checkbox_avg_states[col].get():
            df[f'Скользящее Среднее ({col})'] = df[col].rolling(window=window_size).mean()
            ax.plot(df['Дата'], df[f'Скользящее Среднее ({col})'],
                    label=f'Скользящее Среднее {col} (окно={window_size})', linewidth=2)

    ax.set_xlabel('Дата')
    ax.set_ylabel('Значение')
    ax.set_title(f'График скользящего среднего (окно={window_size})')
    ax.legend()
    ax.grid(True)

   
    canvas.draw()


def on_slider_change(value):
    window_size = int(float(value))
    plot_moving_average(window_size)


def on_combobox_change(event):
    plot_moving_average(window_size=int(window_size_var.get()))


def on_checkbox_change():
    plot_moving_average(window_size=int(window_size_var.get()))



window_size_var = tk.StringVar(value=str(window_sizes[0]))  
window_size_combobox = ttk.Combobox(root, textvariable=window_size_var, values=[str(size) for size in window_sizes],
                                    state='readonly')
window_size_combobox.bind('<<ComboboxSelected>>', on_combobox_change)
window_size_combobox.pack(padx=10, pady=10)


for col in columns:
    frame = tk.Frame(root)
    frame.pack(anchor='w', padx=10, pady=2)

   
    chk_values = ttk.Checkbutton(frame, text=f'{col} (значения)', variable=checkbox_states[col],
                                 command=on_checkbox_change)
    chk_values.pack(side='left')

   
    chk_avg = ttk.Checkbutton(frame, text=f'Скользящее Среднее {col}', variable=checkbox_avg_states[col],
                              command=on_checkbox_change)
    chk_avg.pack(side='left')


slider = ttk.Scale(root, from_=7, to=54, orient='horizontal', command=on_slider_change)
slider.set(window_sizes[0])
slider.pack(padx=10, pady=10)


plot_moving_average(window_size=window_sizes[0])

root.mainloop()
