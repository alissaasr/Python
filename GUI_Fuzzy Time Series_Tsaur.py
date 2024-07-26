#!/usr/bin/env python
# coding: utf-8

# In[3]: 


import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import math
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        df = pd.read_excel(file_path)
        column_name = 'Harga'
        try:
            hasil = process_data(df, column_name)
            display_results(hasil)
            plot_forecasts(df['Harga'].values, hasil['forecasts'], hasil['mtx_final_forecast'])
        except Exception as e:
            messagebox.showerror("Error", str(e))

def process_data(df, column_name):
    nilai_maximum = df[column_name].max()
    nilai_minimum = df[column_name].min()

    def get_data_max_min(df):
        D1 = 0
        D2 = 2
        data_maximum = max(df['Harga']) + D2
        data_minimum = min(df['Harga']) - D1
        return data_maximum, data_minimum

    data_maximum, data_minimum = get_data_max_min(df)
    data_length = len(df)
    K = (1 + (3.322*math.log(data_length, 10)))
    K_rounded_up = round(K/1)*1
    if (K_rounded_up < K): K_rounded_up += 1

    def panjang_interval(data_maximum, data_minimum, K_rounded_up):
        return (data_maximum - data_minimum) / K_rounded_up

    panjang_interval = panjang_interval(data_maximum, data_minimum, K_rounded_up)

    def generate_interval(panjang_interval, K_rounded_up, data_minimum):
        mtx_interval = []
        counter = data_minimum
        for i in range(int(K_rounded_up)):
            lower_bound = counter
            upper_bound = lower_bound + panjang_interval
            mtx_interval.append((int(lower_bound), int(upper_bound)))
            counter = upper_bound
        return mtx_interval

    mtx_interval = generate_interval(panjang_interval, K_rounded_up, data_minimum)

    categories = {
        1: "sangat murah",
        2: "murah",
        3: "agak murah",
        4: "normal",
        5: "agak mahal",
        6: "mahal",
        7: "sangat mahal"
    }

    def mid_point(mtx_interval):
        mid_points = []
        for interval in mtx_interval:
            mid_point = sum(interval) / 2
            mid_points.append(mid_point)
        return mid_points

    mid_points = mid_point(mtx_interval)

    def fuzzify(df, mtx_interval, data_length):
        mtx_fuzzify = []
        for i in range(data_length):
            curr_data = df['Harga'][i]
            for j in range(len(mtx_interval)):
                if (j == 0):
                    if (mtx_interval[j][0] <= curr_data) and (mtx_interval[j][1] >= curr_data):
                        mtx_fuzzify.append([i+1, j+1])
                else:
                    if (mtx_interval[j][0] < curr_data) and (mtx_interval[j][1] >= curr_data):
                        mtx_fuzzify.append([i+1, j+1])
        return mtx_fuzzify

    mtx_fuzzify = fuzzify(df, mtx_interval, data_length)

    def flr(mtx_fuzzify):
        mtx_flr = []
        for i in range(len(mtx_fuzzify)-1):
            mtx_flr.append((i+2, mtx_fuzzify[i][1], mtx_fuzzify[i+1][1]))
        return mtx_flr

    mtx_flr = flr(mtx_fuzzify)

    def flrg(mtx_flr, K_rounded_up):
        int_K_rounded_up = int(K_rounded_up)
        mtx_flrg = [[] for _ in range(int_K_rounded_up)]
        for i in range(len(mtx_flr)):
            temp = mtx_flr[i][1] - 1  # Kurangi 1 dari temp agar sesuai dengan rentang indeks
            if 0 <= temp < int_K_rounded_up:
                mtx_flrg[temp].append(mtx_flr[i][2])
            else:
                raise IndexError(f"Index {temp} out of range for mtx_flrg with size {int_K_rounded_up}")
        return mtx_flrg

    mtx_flrg = flrg(mtx_flr, K_rounded_up)

    def generate_big_mtx(flrg_matrix):
        flrg_matrix = np.array(flrg_matrix)
        row_sums = np.sum(flrg_matrix, axis=1)
        prob_array = np.zeros_like(flrg_matrix, dtype=float)

        for i, row_sum in enumerate(row_sums):
            if row_sum != 0:
                prob_array[i] = flrg_matrix[i] / row_sum
            else:
                prob_array[i] = np.zeros(flrg_matrix.shape[1])  # Atur baris dengan nilai nol

        np.set_printoptions(precision=3, suppress=True)
        return prob_array

    # Membuat matriks dari jumlah FLRG
    flrg_matrix = [[0 for _ in range(K_rounded_up)] for _ in range(K_rounded_up)]
    for index, first, second in mtx_flr:
        flrg_matrix[first - 1][second - 1] += 1

    big_mtx = generate_big_mtx(flrg_matrix)

    def forecast_next_period(mtx_fuzzify, mid_points, big_mtx, last_forecast):
        last_period_fuzzify = mtx_fuzzify[-1][1]  # Mendapatkan hasil mtx_fuzzify dari periode terakhir

        # Mencari indeks baris yang memiliki probabilitas 1
        idx_prob_1 = -1
        for idx, prob_row in enumerate(big_mtx):
            if all(prob == 0 for prob in prob_row[:-1]) and prob_row[-1] == 1:
                idx_prob_1 = idx
                break

        # Periksa apakah indeks valid
        if idx_prob_1 >= 0 and idx_prob_1 < len(mid_points):
            # Menghitung nilai peramalan berdasarkan aturan yang diberikan
            if last_period_fuzzify == idx_prob_1 + 1:
                forecast = mid_points[idx_prob_1]
            else:
                forecast_sum = 0
                for i, prob in enumerate(big_mtx[last_period_fuzzify - 1]):
                    if i == last_period_fuzzify - 1:
                        forecast_sum += last_forecast * prob
                    else:
                        forecast_sum += mid_points[i] * prob
                forecast = forecast_sum
        else:
            forecast = None

        return forecast

    def forecast_all_periods(mtx_fuzzify, mid_points, big_mtx, actual_data):
        forecasts = []
        for i in range(1, len(mtx_fuzzify)):
            last_forecast = actual_data[i - 1]  # Menggunakan data aktual
            forecast = forecast_next_period(mtx_fuzzify[:i], mid_points, big_mtx, last_forecast)
            forecasts.append(forecast)
        return forecasts

    # Data yang diketahui
    mtx_fuzzify = fuzzify(df, mtx_interval, data_length)
    actual_data = df['Harga'].values

    # Peramalan untuk semua periode
    forecasts = forecast_all_periods(mtx_fuzzify, mid_points, big_mtx, actual_data)

    def generate_mtx_dt(mtx_flr, panjang_interval):
        mtx_dt = []
        for i in range(len(mtx_flr)):
            elem = mtx_flr[i]
            selisih = elem[2] - elem[1]
            if selisih == 0:
                mtx_dt.append(0)
            else:
                mtx_dt.append(panjang_interval / 2 * selisih)
        return mtx_dt

    nilai_penyesuaian = generate_mtx_dt(mtx_flr, panjang_interval)

    def generate_mtx_final_forecast(df, forecasts, mtx_dt):
        mtx = [None]  # Periode pertama tidak memiliki peramalan
        for i in range(len(forecasts)):
            temp = abs(forecasts[i] + mtx_dt[i])
            mtx.append(temp)
        return mtx

    mtx_final_forecast = generate_mtx_final_forecast(df, forecasts, nilai_penyesuaian)
    
    # Peramalan untuk periode berikutnya
    def forecast_next_period(mtx_fuzzify, mid_points, big_mtx, forecasts):
        last_period_fuzzify = mtx_fuzzify[-1][1]  # Mendapatkan hasil mtx_fuzzify dari periode terakhir

        # Mencari indeks baris yang memiliki probabilitas 1
        idx_prob_1 = -1
        for idx, prob_row in enumerate(big_mtx):
            if all(prob == 0 for prob in prob_row[:-1]) and prob_row[-1] == 1:
                idx_prob_1 = idx
                break

        # Periksa apakah indeks valid
        if idx_prob_1 >= 0 and idx_prob_1 < len(mid_points):
            # Menghitung nilai peramalan berdasarkan aturan yang diberikan
            if last_period_fuzzify == idx_prob_1 + 1:
                forecast = mid_points[idx_prob_1]
            else:
                forecast_sum = 0
                for i, prob in enumerate(big_mtx[last_period_fuzzify - 1]):
                    if i == last_period_fuzzify - 1:
                        forecast_sum += forecasts * prob
                    else:
                        forecast_sum += mid_points[i] * prob
                forecast = forecast_sum
        else:
            forecast = None

        return forecast

    next_forecast = forecast_next_period(mtx_fuzzify, mid_points, big_mtx, forecasts[-1])

    # Hitung MAPE
    def mean_absolute_percentage_error(y_true, y_pred):
        y_true, y_pred = np.array(y_true), np.array(y_pred)
        return np.mean(np.abs((y_true - y_pred) / y_true)) * 100

    mape = mean_absolute_percentage_error(actual_data[1:], mtx_final_forecast[1:])

    hasil = {
        'nilai_penyesuaian': nilai_penyesuaian,
        'forecasts': forecasts,
        'mtx_final_forecast': mtx_final_forecast,
        'next_forecast': next_forecast,
        'mape': mape
    }

    return hasil

def display_results(hasil):
    results_window = tk.Toplevel(root)
    results_window.title("Results")

    text_area = tk.Text(results_window, wrap=tk.WORD)
    text_area.pack(expand=True, fill=tk.BOTH)

    text_area.insert(tk.END, "\nPeramalan Awal:\n")
    for i, forecast in enumerate(hasil["forecasts"]):
        if forecast is not None:
            text_area.insert(tk.END, f"Peramalan Awal {i + 2}: {forecast}\n")
        else:
            text_area.insert(tk.END, f"No forecast available for period {i + 2}\n")

    text_area.insert(tk.END, "\nPenyesuaian Nilai:\n")
    for i, nilai in enumerate(hasil["nilai_penyesuaian"], start=2):  # Mulai dari 2 karena periode pertama kosong
        text_area.insert(tk.END, f"Nilai penyesuaian t={i} adalah {nilai}\n")

    text_area.insert(tk.END, "\nPeramalan Akhir:\n")
    for i, forecast in enumerate(hasil["mtx_final_forecast"][1:], start=2):  # Mulai dari 1 karena periode pertama kosong
        text_area.insert(tk.END, f"Peramalan Akhir Periode {i}: {forecast}\n")

    text_area.insert(tk.END, f"Peramalan periode selanjutnya: {hasil['next_forecast']}\n")

    text_area.insert(tk.END, f"\nAverage MAPE: {hasil['mape']:.2f}%\n")

def plot_forecasts(actual_data, forecasts, final_forecast):
    fig, ax = plt.subplots()
    ax.plot(actual_data, label='Data Aktual')
    ax.plot(range(1, len(forecasts) + 1), forecasts, label='Peramalan')
    ax.plot(range(len(final_forecast)), final_forecast, label='Peramalan Akhir', linestyle='dashed')
    ax.legend()
    canvas = FigureCanvasTkAgg(fig, master=frame_plot)
    canvas.draw()
    canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

# Antarmuka GUI
root = tk.Tk()
root.title("Fuzzy Time Series Forecasting")

frame_buttons = tk.Frame(root)
frame_buttons.pack(side=tk.TOP, fill=tk.X)

frame_results_plot = tk.Frame(root)
frame_results_plot.pack(side=tk.TOP, fill=tk.BOTH, expand=1)

load_button = tk.Button(frame_buttons, text="Load Excel File", command=load_file)
load_button.pack(side=tk.LEFT)

results_text = tk.Text(frame_results_plot, wrap=tk.WORD)
results_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

frame_plot = tk.Frame(frame_results_plot)
frame_plot.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

root.mainloop()


# In[ ]:




