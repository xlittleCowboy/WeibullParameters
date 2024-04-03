import customtkinter
from os import path
from scipy.stats import weibull_min
from scipy.stats import gamma
import glob
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import sys

font_size = 18

frame_size_x = 450
frame_size_y = 200

frame_border_width = 7.5

xls = None
df = pd.DataFrame()
data = []

shape = loc = scale = 0

dist_for_params = ""

customtkinter.set_appearance_mode("Dark")

window = customtkinter.CTk()
window.title("Оценка параметров распределения Вейбулла")
window.geometry("1000x700")
window.resizable(False, False)

MainFrame = customtkinter.CTkFrame(window)
MainFrame.pack(expand=True, anchor="center")

buttonsFrame = customtkinter.CTkFrame(MainFrame, border_width=frame_border_width, width=frame_size_x, height=frame_size_y)
buttonsFrame.grid_rowconfigure(0, weight=1)
buttonsFrame.grid_columnconfigure(0, weight=1)
buttonsFrame.grid_rowconfigure(1, weight=1)
buttonsFrame.grid_columnconfigure(1, weight=1)
buttonsFrame.grid_propagate(0)
buttonsFrame.grid(sticky="nsew", row=1, column=1, padx=10, pady=10)

fileFrame = customtkinter.CTkFrame(MainFrame, border_color="red", border_width=frame_border_width, width=frame_size_x, height=frame_size_y)
fileFrame.grid_rowconfigure(1, weight=1)
fileFrame.grid_columnconfigure(2, weight=1)
fileFrame.grid_propagate(0)
fileFrame.grid(sticky="nsew", row=1, column=2, padx=10, pady=10)

calculateFrame = customtkinter.CTkFrame(MainFrame, border_width=frame_border_width, width=frame_size_x, height=frame_size_y)
calculateFrame.grid_rowconfigure(2, weight=1)
calculateFrame.grid_columnconfigure(1, weight=1)
calculateFrame.grid_propagate(0)
calculateFrame.grid(sticky="nsew", row=2, column=1, padx=10, pady=10)

paramsFrame = customtkinter.CTkFrame(MainFrame, border_color="red", border_width=frame_border_width, width=frame_size_x, height=frame_size_y)
paramsFrame.grid_rowconfigure(2, weight=1)
paramsFrame.grid_columnconfigure(2, weight=1)
paramsFrame.grid_propagate(0)
paramsFrame.grid(sticky="nsew", row=2, column=2, padx=10, pady=10)

settingsFrame = customtkinter.CTkFrame(MainFrame, border_width=frame_border_width, width=frame_size_x, height=frame_size_y)
settingsFrame.grid_rowconfigure(3, weight=1)
settingsFrame.grid_columnconfigure(2, weight=1)
settingsFrame.grid_propagate(0)
settingsFrame.grid(sticky="nsew", row=3, column=2, padx=10, pady=10)

probabilityFrame = customtkinter.CTkFrame(MainFrame, border_width=frame_border_width, width=frame_size_x, height=frame_size_y)
probabilityFrame.grid_rowconfigure(3, weight=1)
probabilityFrame.grid_columnconfigure(1, weight=1)
probabilityFrame.grid_propagate(0)
probabilityFrame.grid(sticky="nsew", row=3, column=1, padx=10, pady=10)

choice_btn = customtkinter.CTkButton(buttonsFrame, text="Выбрать файл", font=("CTkDefaultFont", font_size))
choice_btn.grid(sticky="e", row=0, column=0, columnspan=1, padx=10, pady=0)

delete_btn = customtkinter.CTkButton(buttonsFrame, text="Удалить файл", font=("CTkDefaultFont", font_size))
delete_btn.grid(sticky="w", row=0, column=1, columnspan=1, padx=10, pady=0)

sheet_lb = customtkinter.CTkLabel(buttonsFrame, text="Лист: ", font=("CTkDefaultFont", font_size))
sheet_lb.grid(sticky="e", row=1, column=0, padx=10, pady=0)

sheetName = customtkinter.StringVar(value="") 

sheet_cb = customtkinter.CTkComboBox(buttonsFrame, variable=sheetName, values=[], state="readonly", font=("CTkDefaultFont", font_size))
sheet_cb.grid(sticky="w", row=1, column=1, padx=10, pady=0)

column_lb = customtkinter.CTkLabel(buttonsFrame, text="Столбец: ", font=("CTkDefaultFont", font_size))
column_lb.grid(sticky="e", row=2, column=0, padx=10, pady=0)

columnName = customtkinter.StringVar(value="")

column_cb = customtkinter.CTkComboBox(buttonsFrame, variable=columnName, values=[], state="readonly", font=("CTkDefaultFont", font_size))
column_cb.grid(sticky="w", row=2, column=1, padx=10, pady=20)

filepath_lb = customtkinter.CTkLabel(fileFrame, text="Выберите Excel файл!", font=("CTkDefaultFont", font_size))
filepath_lb.grid(row=1, column=2, padx=10, pady=10)

distributions = {"Вейбулла трехпараметрическое" : "w3", "Вейбулла двухпараметрическое" : "w2", "Вейбулла экспоненциальное" : "we", "Гамма трехпараметрическое" : "g3", "Гамма двухпараметрическое" : "g2"}
distribution = customtkinter.StringVar(value="Вейбулла трехпараметрическое") 

dist_cb = customtkinter.CTkComboBox(calculateFrame, variable=distribution, values=distributions.keys(), state="readonly", font=("CTkDefaultFont", font_size))
dist_cb.grid(sticky="we", row=1, column=1, padx=40, pady=(20, 0))

methods = {"Метод максимального правдоподобия" : "MLE", "Метод моментов" : "MM"}

calcMethod = customtkinter.StringVar(value="Метод максимального правдоподобия") 

method_cb = customtkinter.CTkComboBox(calculateFrame, variable=calcMethod, values=methods.keys(), state="readonly", font=("CTkDefaultFont", font_size))
method_cb.grid(sticky="we", row=2, column=1, padx=40, pady=0)

calculate_btn = customtkinter.CTkButton(calculateFrame, text="Оценить параметры", font=("CTkDefaultFont", font_size))
calculate_btn.grid(row=3, column=1, padx=10, pady=(0, 10))

plot_btn = customtkinter.CTkButton(calculateFrame, text="Показать график", font=("CTkDefaultFont", font_size))
plot_btn.grid(row=4, column=1, padx=10, pady=(0, 20))

params_lb = customtkinter.CTkLabel(paramsFrame, text='Excel файл не выбран!', font=("CTkDefaultFont", font_size))
params_lb.grid(row=2, column=2, padx=10, pady=10)

settings_lb = customtkinter.CTkLabel(settingsFrame, text='Количество знаков после запятой', font=("CTkDefaultFont", font_size))
settings_lb.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

scale_commas_lb = customtkinter.CTkLabel(settingsFrame, text='Для параметра масштаба (a): ', font=("CTkDefaultFont", font_size))
scale_commas_lb.grid(sticky="w", row=1, column=0, padx=(20, 0), pady=10)

scale_commas_te = customtkinter.CTkEntry(settingsFrame, font=("CTkDefaultFont", font_size))
scale_commas_te.grid(sticky="w", row=1, column=1, padx=10, pady=10)
scale_commas_te.insert(0, "6")

shape_commas_lb = customtkinter.CTkLabel(settingsFrame, text='Для параметра формы (b): ', font=("CTkDefaultFont", font_size))
shape_commas_lb.grid(sticky="w", row=2, column=0, padx=(20, 0), pady=10)

shape_commas_te = customtkinter.CTkEntry(settingsFrame, font=("CTkDefaultFont", font_size))
shape_commas_te.grid(sticky="w", row=2, column=1, padx=10, pady=10)
shape_commas_te.insert(0, "6")

loc_commas_lb = customtkinter.CTkLabel(settingsFrame, text='Для параметра сдвига (c): ', font=("CTkDefaultFont", font_size))
loc_commas_lb.grid(sticky="w", row=3, column=0, padx=(20, 0), pady=10)

loc_commas_te = customtkinter.CTkEntry(settingsFrame, font=("CTkDefaultFont", font_size))
loc_commas_te.grid(sticky="w", row=3, column=1, padx=10, pady=10)
loc_commas_te.insert(0, "6")

probability_lb = customtkinter.CTkLabel(probabilityFrame, text='Вероятность попадания X в интервал', font=("CTkDefaultFont", font_size))
probability_lb.grid(row=0, column=0, columnspan=3, padx=(20, 0), pady=10)

lower_edge_lb = customtkinter.CTkLabel(probabilityFrame, text='Нижняя граница: ', font=("CTkDefaultFont", font_size))
lower_edge_lb.grid(sticky="w", row=1, column=0, padx=(20, 0), pady=10)

lower_edge_te = customtkinter.CTkEntry(probabilityFrame, font=("CTkDefaultFont", font_size))
lower_edge_te.grid(sticky="w", row=1, column=1, padx=(20, 0), pady=10)

upper_edge_lb = customtkinter.CTkLabel(probabilityFrame, text='Верхняя граница: ', font=("CTkDefaultFont", font_size))
upper_edge_lb.grid(sticky="w", row=2, column=0, padx=(20, 0), pady=10)

upper_edge_te = customtkinter.CTkEntry(probabilityFrame, font=("CTkDefaultFont", font_size))
upper_edge_te.grid(sticky="w", row=2, column=1, padx=(20, 0), pady=10)

probability_btn = customtkinter.CTkButton(probabilityFrame, text="Рассчитать вероятность", font=("CTkDefaultFont", font_size))
probability_btn.grid(row=3, column=0, padx=(20, 0), pady=10)

probability_result_lb = customtkinter.CTkLabel(probabilityFrame, text='F(x) = ', font=("CTkDefaultFont", font_size))
probability_result_lb.grid(sticky="w", row=3, column=1, padx=(20, 0), pady=10)

def update_column_cb():
	global data

	df.columns = df.columns.astype(str)
	column_cb.configure(values=df.columns)

	if len(df.columns) > 0:
		column_cb.set(df.columns[0])
		data = df[df.columns[0]]

		params_lb.configure(text='Выберите нужный лист и колонку, а затем\nнажмите кнопку "Оценить параметры"')
		paramsFrame.configure(border_color="green")
	else:
		params_lb.configure(text='В выбранном листе нет столбцов!')
		paramsFrame.configure(border_color="red")

def open_file():
	global xls, df, data

	file = None
	filepath = customtkinter.filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
	if (filepath != ""):
		files = glob.glob(filepath)
		if (len(files) > 0):
			file = files[0]

	if (file is None):
		delete_file()

		return
	else:
		filepath_lb.configure(text="Выбранный файл: \n\n" + path.basename(filepath))
		fileFrame.configure(border_color="green")

	xls = pd.ExcelFile(filepath)
	sheet_cb.configure(values=xls.sheet_names)
	sheet_cb.set(xls.sheet_names[0])

	df = xls.parse(xls.sheet_names[0])

	update_column_cb()

def sheet_name_selected(choice):
	global df
	df = xls.parse(choice)

	update_column_cb()

def column_name_selected(choice):
	global data
	data = df[choice]

def calculate_params():
	global shape, loc, scale, dist_for_params
	if len(data) > 0:
		for i in range(0, len(data)):
			if (pd.isna(data[i])):
				del data[i]

	if (len(data) <= 0):
		params_lb.configure(text="Столбец с данными пуст!")
		paramsFrame.configure(border_color="red")
		return False

	dist_for_params = distributions[distribution.get()]

	if (dist_for_params == "w3" or dist_for_params == "we" or dist_for_params == "w2"):
		params = weibull_min.fit(data, method=methods[calcMethod.get()])
	elif (dist_for_params == "g3" or dist_for_params == "g2"):
		params = gamma.fit(data, method=methods[calcMethod.get()])

	shape = params[0]
	loc = params[1]
	scale = params[2]

	if (dist_for_params == "we"):
		params = weibull_min.fit(data, method=methods[calcMethod.get()])
		shape = 1
		loc = 0
	if (dist_for_params == "w2" or dist_for_params == "g2"):
		params = weibull_min.fit(data, method=methods[calcMethod.get()])
		loc = 0

	params_lb.configure(text="Масштаб (a): " + str(round(scale, int(scale_commas_te.get()))) + "\n\nФорма (b): " + str(round(shape, int(shape_commas_te.get()))) + "\n\nСдвиг (c): " + str(round(loc, int(loc_commas_te.get()))))

	paramsFrame.configure(border_color="green")

def show_plot():
	if (len(data) <= 0):
		return

	fig, ax = plt.subplots(2)
	ax[0].hist(data, density=True, bins='auto', histtype='stepfilled', alpha=0.2, label='Экспериментальные данные')

	if (dist_for_params == "w3" or dist_for_params == "we" or dist_for_params == "w2"):
		x = np.linspace(weibull_min.ppf(0.0000001, shape, loc=loc, scale=scale), weibull_min.ppf(0.9999999, shape, loc=loc, scale=scale), 1000)
		ax[0].plot(x, weibull_min.pdf(x, shape, loc=loc, scale=scale), lw=2, alpha=0.6, color='red', label='Функция плотности вероятности')
	elif (dist_for_params == "g3" or dist_for_params == "g2"):
		x = np.linspace(gamma.ppf(0.0000001, shape, loc=loc, scale=scale), gamma.ppf(0.9999999, shape, loc=loc, scale=scale), 1000)
		ax[0].plot(x, gamma.pdf(x, shape, loc=loc, scale=scale), lw=2, alpha=0.6, color='red', label='Функция плотности вероятности')
	
	ax[0].legend(loc='upper right', frameon=False, fontsize='large')

	if (dist_for_params == "w3" or dist_for_params == "we" or dist_for_params == "w2"):
		x = np.linspace(weibull_min.ppf(0.0000001, shape, loc=loc, scale=scale), weibull_min.ppf(0.9999999, shape, loc=loc, scale=scale), 1000)
		ax[1].plot(x, weibull_min.cdf(x, shape, loc=loc, scale=scale), lw=2, alpha=0.6, color='green', label='Функция плотности распределения')
	elif (dist_for_params == "g3" or dist_for_params == "g2"):
		x = np.linspace(gamma.ppf(0.0000001, shape, loc=loc, scale=scale), gamma.ppf(0.9999999, shape, loc=loc, scale=scale), 1000)
		ax[1].plot(x, gamma.cdf(x, shape, loc=loc, scale=scale), lw=2, alpha=0.6, color='green', label='Функция плотности распределения')
	
	ax[1].legend(loc='lower right', frameon=False, fontsize='large')

	plt.show()

def check_commas(newval):
	if (newval == "" or newval.isdigit()):
		return True

	return False

def check_edge(newval):
    parts = newval.split('.')
    parts_number = len(parts)

    if parts_number > 2:
        return False

    if parts_number > 1 and parts[1]: # don't check empty string
        if not parts[1].isdecimal():
            return False

    if parts_number > 0 and parts[0]: # don't check empty string
        if not parts[0].isdecimal():
            return False

    return True

def calculate_probability():
	if (len(data) <= 0):
		return

	upper_edge = upper_edge_te.get()
	lower_edge = lower_edge_te.get()

	if (lower_edge == ""):
		lower_edge = 0

	if (upper_edge == ""):
		upper_edge = max(data) * 2

	if (dist_for_params == "w3" or dist_for_params == "we" or dist_for_params == "w2"):
		lower_cdf = weibull_min.cdf(float(lower_edge), shape, loc, scale)
		upper_cdf = weibull_min.cdf(float(upper_edge), shape, loc, scale)
	elif (dist_for_params == "g3" or dist_for_params == "g2"):
		lower_cdf = gamma.cdf(float(lower_edge), shape, loc, scale)
		upper_cdf = gamma.cdf(float(upper_edge), shape, loc, scale)

	probability_result_lb.configure(text="F(x) = " + str(round((upper_cdf - lower_cdf), 6))) 

def delete_file():
	global xls, df, data

	filepath_lb.configure(text="Выберите Excel файл!")
	fileFrame.configure(border_color="red")
	params_lb.configure(text='Excel файл не выбран!')
	paramsFrame.configure(border_color="red")

	sheet_cb.configure(values=[])
	sheet_cb.set("")
	column_cb.configure(values=[])
	column_cb.set("")

	df = pd.DataFrame()
	data = []

	xls.close()

choice_btn.configure(command=open_file)
delete_btn.configure(command=delete_file)
sheet_cb.configure(command=sheet_name_selected)
column_cb.configure(command=column_name_selected)
calculate_btn.configure(command=calculate_params)
plot_btn.configure(command=show_plot)
probability_btn.configure(command=calculate_probability)

commas_check = (window.register(check_commas), "%P")
shape_commas_te.configure(validate="key", validatecommand=commas_check)
loc_commas_te.configure(validate="key", validatecommand=commas_check)
scale_commas_te.configure(validate="key", validatecommand=commas_check)

edge_check = (window.register(check_edge), "%P")
lower_edge_te.configure(validate="key", validatecommand=edge_check)
upper_edge_te.configure(validate="key", validatecommand=edge_check)

window.mainloop()