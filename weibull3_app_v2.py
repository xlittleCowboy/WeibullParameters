import customtkinter
from os import path
from scipy.stats import weibull_min
import glob
import pandas as pd

xls = None
df = pd.DataFrame()
data = []

window = customtkinter.CTk()
window.title("Оценка параметров распределения Вейбулла")
window.geometry("800x450")
window.resizable(False, False)

MainFrame = customtkinter.CTkFrame(window)
MainFrame.pack(expand=True, anchor="center")

buttonsFrame = customtkinter.CTkFrame(MainFrame, border_width=2, width=350, height=150)
buttonsFrame.grid_rowconfigure(0, weight=1)
buttonsFrame.grid_columnconfigure(0, weight=1)
buttonsFrame.grid_rowconfigure(1, weight=1)
buttonsFrame.grid_columnconfigure(1, weight=1)
buttonsFrame.grid_propagate(0)
buttonsFrame.grid(sticky="nsew", row=1, column=1, padx=10, pady=10)

fileFrame = customtkinter.CTkFrame(MainFrame, border_color="red", border_width=2, width=350, height=150)
fileFrame.grid_rowconfigure(1, weight=1)
fileFrame.grid_columnconfigure(2, weight=1)
fileFrame.grid_propagate(0)
fileFrame.grid(sticky="nsew", row=1, column=2, padx=10, pady=10)

calculateFrame = customtkinter.CTkFrame(MainFrame, border_width=2, width=350, height=150)
calculateFrame.grid_rowconfigure(2, weight=1)
calculateFrame.grid_columnconfigure(1, weight=1)
calculateFrame.grid_propagate(0)
calculateFrame.grid(sticky="nsew", row=2, column=1, padx=10, pady=10)

paramsFrame = customtkinter.CTkFrame(MainFrame, border_color="red", border_width=2, width=350, height=150)
paramsFrame.grid_rowconfigure(2, weight=1)
paramsFrame.grid_columnconfigure(2, weight=1)
paramsFrame.grid_propagate(0)
paramsFrame.grid(sticky="nsew", row=2, column=2, padx=10, pady=10)

choice_btn = customtkinter.CTkButton(buttonsFrame, text="Выбрать файл с данными", font=("CTkDefaultFont", 15))
choice_btn.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

sheet_lb = customtkinter.CTkLabel(buttonsFrame, text="Лист: ", font=("CTkDefaultFont", 15))
sheet_lb.grid(sticky="e", row=1, column=0, padx=10, pady=10)

sheetName = customtkinter.StringVar(value="") 

sheet_cb = customtkinter.CTkComboBox(buttonsFrame, variable=sheetName, values=[], state="readonly", font=("CTkDefaultFont", 15))
sheet_cb.grid(sticky="w", row=1, column=1, padx=10, pady=10)

column_lb = customtkinter.CTkLabel(buttonsFrame, text="Столбец: ", font=("CTkDefaultFont", 15))
column_lb.grid(sticky="e", row=2, column=0, padx=10, pady=10)

columnName = customtkinter.StringVar(value="")

column_cb = customtkinter.CTkComboBox(buttonsFrame, variable=columnName, values=[], state="readonly", font=("CTkDefaultFont", 15))
column_cb.grid(sticky="w", row=2, column=1, padx=10, pady=10)

filepath_lb = customtkinter.CTkLabel(fileFrame, text="Выберите верный Excel файл!", font=("CTkDefaultFont", 15))
filepath_lb.grid(row=1, column=2, padx=10, pady=10)

calcMethod = customtkinter.StringVar(value="MLE") 

mle_btn = customtkinter.CTkRadioButton(calculateFrame, text="Метод максимального правдоподобия", value="MLE", variable=calcMethod, font=("CTkDefaultFont", 15))
mle_btn.grid(row=1, column=1, padx=10, pady=10)
  
mm_btn = customtkinter.CTkRadioButton(calculateFrame, text="Метод моментов", value="MM", variable=calcMethod, font=("CTkDefaultFont", 15))
mm_btn.grid(row=2, column=1, padx=10, pady=10)

calculate_btn = customtkinter.CTkButton(calculateFrame, text="Оценить параметры", font=("CTkDefaultFont", 15))
calculate_btn.grid(row=3, column=1, padx=10, pady=10)

params_lb = customtkinter.CTkLabel(paramsFrame, text='Excel файл не выбран!', font=("CTkDefaultFont", 15))
params_lb.grid(row=2, column=2, padx=10, pady=10)

def update_column_cb():
	global data
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
	global xls
	global df
	global data

	file = None
	filepath = customtkinter.filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
	if (filepath != ""):
		files = glob.glob(filepath)
		if (len(files) > 0):
			file = files[0]

	if (file is None):
		filepath_lb.configure(text="Выбран неверный файл!")
		fileFrame.configure(border_color="red")
		params_lb.configure(text='Выберите верный Excel файл!')
		paramsFrame.configure(border_color="red")

		sheet_cb.configure(values=[])
		sheet_cb.set("")
		column_cb.configure(values=[])
		column_cb.set("")

		df = pd.DataFrame()
		data = []

		return
	else:
		filepath_lb.configure(text="Выбранный файл: " + path.basename(filepath))
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
	if len(data) > 0:
		for i in range(0, len(data)):
			if (pd.isna(data[i])):
				del data[i]

	if (len(data) <= 0):
		params_lb.configure(text="Столбец с данными пуст!")
		paramsFrame.configure(border_color="red")
		return False

	params = weibull_min.fit(data, method=calcMethod.get())
	params_lb.configure(text="Форма: " + str(params[0]) + "\nСдвиг: " + str(params[1]) + "\nМасштаб: " + str(params[2]))

	paramsFrame.configure(border_color="green")

choice_btn.configure(command=open_file)
sheet_cb.configure(command=sheet_name_selected)
column_cb.configure(command=column_name_selected)
calculate_btn.configure(command=calculate_params)

window.mainloop()