#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import math
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

def load_file_chen():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        df = pd.read_excel(file_path)
        column_name = 'Harga'
        try:
            hasil = process_data_chen(df, column_name)
            display_results_chen(hasil)
            plot_forecasts(df['Harga'].values, hasil['forecasts'])
        except Exception as e:
            messagebox.showerror("Error", str(e))

def process_data_chen(df, column_name):
    nilai_maximum = df[column_name].max()
    nilai_minimum = df[column_name].min()

    def get_data_max_min(df):
        D1 = 0
        D2 = 1212.71
        data_maximum = max(df[column_name]) + D2
        data_minimum = min(df[column_name]) - D1
        return data_maximum, data_minimum

    data_maximum, data_minimum = get_data_max_min(df)
    data_length = len(df)
    K = (1 + (3.322 * math.log(data_length, 10)))
    K_rounded_up = round(K / 1) * 1
    if K_rounded_up < K:
        K_rounded_up += 1

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
            curr_data = (df[column_name][i])
            for j in range(len(mtx_interval)):
                if j == 0:
                    if (mtx_interval[j][0] <= curr_data) and (mtx_interval[j][1] >= curr_data):
                        mtx_fuzzify.append([i + 1, j + 1])
                else:
                    if (mtx_interval[j][0] < curr_data) and (mtx_interval[j][1] >= curr_data):
                        mtx_fuzzify.append([i + 1, j + 1])
        return mtx_fuzzify

    fuzzyfikasi = fuzzify(df, mtx_interval, data_length)

    def flr(mtx_fuzzify):
        mtx_flr = []
        for i in range(len(mtx_fuzzify) - 1):
            mtx_flr.append((i + 2, mtx_fuzzify[i][1], mtx_fuzzify[i + 1][1]))
        return mtx_flr

    mtx_flr = flr(fuzzyfikasi)

    def flrg(mtx_flr, K_rounded_up):
        int_K_rounded_up = int(K_rounded_up)
        mtx_flrg = [[] for _ in range(int_K_rounded_up)]
        for i in range(len(mtx_flr)):
            temp = mtx_flr[i][1] - 1
            val = mtx_flr[i][2]
            if 0 <= temp < int_K_rounded_up:
                if val not in mtx_flrg[temp]:
                    mtx_flrg[temp].append(val)
        return mtx_flrg

    mtx_flrg = flrg(mtx_flr, K_rounded_up)

    def forecast_next_period(mtx_fuzzify, mtx_flrg, mid_points):
        if len(mtx_fuzzify) < 2:
            raise ValueError("Not enough data to make a forecast for the next period.")
        last_period_fuzzify = mtx_fuzzify[-1][1]
        previous_period_fuzzify = mtx_fuzzify[-2][1]
        forecast_values = mtx_flrg[previous_period_fuzzify - 1]
        forecast_sum = sum(mid_points[val - 1] for val in forecast_values)
        forecast = forecast_sum / len(forecast_values) if len(forecast_values) > 0 else 0
        return forecast

    def forecast_all_periods(mtx_fuzzify, mtx_flrg, mid_points):
        forecasts = []
        for i in range(len(mtx_fuzzify)):
            if i == 0:
                forecasts.append(None)
            else:
                previous_period_fuzzify = mtx_fuzzify[i - 1][1]
                forecast_values = mtx_flrg[previous_period_fuzzify - 1]
                forecast_sum = sum(mid_points[val - 1] for val in forecast_values)
                forecast = forecast_sum / len(forecast_values) if len(forecast_values) > 0 else 0
                forecasts.append(forecast)
        return forecasts

    forecasts = forecast_all_periods(fuzzyfikasi, mtx_flrg, mid_points)
    next_forecast = forecast_next_period(fuzzyfikasi, mtx_flrg, mid_points)

    def calculate_mape(actual, forecast):
        if len(actual) != len(forecast):
            raise ValueError("The length of actual and forecast lists must be the same")
        mape = []
        for a, f in zip(actual, forecast):
            if a != 0:
                mape.append(abs((a - f) / a) * 100)
        return mape

    actual = df[column_name].tolist()
    mape_values = calculate_mape(actual[1:], [f for f in forecasts[1:] if f is not None])
    overall_mape = sum(mape_values) / len(mape_values)

    hasil = {
        "intervals": mtx_interval,
        "categories": categories,
        "mid_points": mid_points,
        "fuzzyfikasi": fuzzyfikasi,
        "flr": mtx_flr,
        "flrg": mtx_flrg,
        "forecasts": forecasts,
        "next_forecast": next_forecast,
        "mape_values": mape_values,
        "overall_mape": overall_mape
    }
    return hasil

def display_results_chen(hasil):    
    results_window = tk.Toplevel(root)
    results_window.title("Results")

    text_area = tk.Text(results_window, wrap=tk.WORD)
    text_area.pack(expand=True, fill=tk.BOTH)

    text_area.insert(tk.END, "\nHasil Peramalan FTS Chen:\n")
    for i, forecast in enumerate(hasil["forecasts"][1:], start=2):  # Mulai dari 1 karena periode pertama kosong
        text_area.insert(tk.END, f"Peramalan periode {i}: {forecast}\n")

    text_area.insert(tk.END, f"Peramalan periode selanjutnya: {hasil['next_forecast']}\n")

    text_area.insert(tk.END, f"\nNilai MAPE: {hasil['overall_mape']:.2f}%\n")

def plot_forecasts(actual_data, forecasts):
    fig, ax = plt.subplots()
    ax.plot(actual_data, label='Data Aktual')
    ax.plot(range(1, len(forecasts)+1), forecasts, label='Forecast', linestyle='--')
    ax.set_title('FTS Chen')
    ax.set_xlabel('Period')
    ax.set_ylabel('Harga')
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

load_button = tk.Button(frame_buttons, text="Load Excel File", command=load_file_chen)
load_button.pack(side=tk.LEFT)

results_text = tk.Text(frame_results_plot, wrap=tk.WORD)
results_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

frame_plot = tk.Frame(frame_results_plot)
frame_plot.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)

root.mainloop()


# In[ ]:
