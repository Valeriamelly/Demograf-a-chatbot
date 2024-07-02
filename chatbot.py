import tkinter as tk
from tkinter import scrolledtext
from tkinter import ttk
import openpyxl
from fuzzywuzzy import fuzz, process

class ExcelChatbot:
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.questions = []

    def combinar_encabezados(self, sheet):
        encabezados = []
        encabezado_actual = ""

        for col in range(1, sheet.max_column + 1):
            valor_fila_1 = sheet.cell(row=1, column=col).value
            valor_fila_2 = sheet.cell(row=2, column=col).value

            if valor_fila_1:
                encabezado_actual = valor_fila_1.lower()

            if valor_fila_2:
                encabezados.append(f"{encabezado_actual} ({valor_fila_2.lower()})")
            else:
                encabezados.append(encabezado_actual)

        return [encabezado.strip() for encabezado in encabezados]

    def obtener_respuesta(self, pregunta):
        wb = openpyxl.load_workbook(self.excel_file)
        sheet = wb.active

        encabezados = self.combinar_encabezados(sheet)
        encabezados = [encabezado.strip().lower() for encabezado in encabezados]
        pregunta = pregunta.lower().strip()

        respuesta = None
        categoria = None

        categorias_mapeadas = {
            "esperanza de vida hombre": "esperanza de vida al nacer (años) (hombre)",
            "esperanza de vida mujer": "esperanza de vida al nacer (años) (mujer)",
            "composición 0 a 14": "composición de la población en años (porcentaje) (0 a 14)",
            "composición 15 a 64": "composición de la población en años (porcentaje) (15 a 64)",
            "composición 65 a más": "composición de la población en años (porcentaje) (65 a más)",
            "fecundidad total": "tasa de fecundidad total",
            "poblacion": "población (miles)"
        }

        lista_paises = [row[0].lower() for row in sheet.iter_rows(min_row=3, max_row=77, min_col=1, max_col=1, values_only=True) if row[0]]

        pais_mencionado, _ = process.extractOne(pregunta, lista_paises, scorer=fuzz.partial_ratio)

        if "esperanza de vida" in pregunta:
            if "hombre" in pregunta:
                categoria = "esperanza de vida hombre"
            elif "mujer" in pregunta:
                categoria = "esperanza de vida mujer"
        elif "composición de la población" in pregunta:
            if "0 a 14" in pregunta:
                categoria = "composición 0 a 14"
            elif "15 a 64" in pregunta:
                categoria = "composición 15 a 64"
            elif "65 a más" in pregunta:
                categoria = "composición 65 a más"
        elif "fecundidad total" in pregunta or "tasa de fecundidad total" in pregunta:
            categoria = "fecundidad total"
        elif "poblacion" in pregunta or "población" in pregunta:
            categoria = "poblacion"

        if not pais_mencionado or not categoria:
            return "No pude identificar el país o la categoría en tu pregunta. Por favor, intenta de nuevo."

        for row in sheet.iter_rows(min_row=3, max_row=77, min_col=1, max_col=sheet.max_column, values_only=True):
            nombre_pais = row[0].lower()
            if pais_mencionado in nombre_pais:
                categoria_mapeada = categorias_mapeadas.get(categoria)
                if categoria_mapeada in encabezados:
                    indice_categoria = encabezados.index(categoria_mapeada)
                    respuesta = row[indice_categoria]
                    break

        if respuesta is not None:
            return f"La {categoria} en {pais_mencionado} es {respuesta}."
        else:
            return f"No encontré información sobre {categoria} para {pais_mencionado}."

    def answer_question(self, question):
        self.questions.append(question)
        return self.obtener_respuesta(question)

def submit_question(event=None):
    question = question_entry.get()
    if question:
        chatbot_response = chatbot.answer_question(question)
        chat_log.insert(tk.END, f"Usuario: {question}\n")
        chat_log.insert(tk.END, f"Chatbot: {chatbot_response}\n\n")
        question_entry.delete(0, tk.END)

def show_questions():
    chat_log.insert(tk.END, "Historial de preguntas:\n")
    for q in chatbot.questions:
        chat_log.insert(tk.END, f"- {q}\n")
    chat_log.insert(tk.END, "\n")

def show_help():
    help_text = (
        "Ejemplos de preguntas que puedes hacer:\n"
        "1. ¿Cuál es la población de [país]?\n"
        "2. ¿Cuál es la composición de la población de 0 a 14 años en [país]?\n"
        "3. ¿Cuál es la composición de la población de 15 a 64 años en [país]?\n"
        "4. ¿Cuál es la composición de la población de 65 a más años en [país]?\n"
        "5. ¿Cuál es la tasa de fecundidad total en [país]?\n"
        "6. ¿Cuál es la esperanza de vida al nacer para hombres en [país]?\n"
        "7. ¿Cuál es la esperanza de vida al nacer para mujeres en [país]?\n"
        "8. ¿Cuál es la población total a nivel mundial?\n"
    )
    chat_log.insert(tk.END, f"{help_text}\n")

def show_countries():
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    countries = [row[0] for row in sheet.iter_rows(min_row=3, max_row=77, min_col=1, max_col=1, values_only=True) if row[0]]
    chat_log.insert(tk.END, "Países disponibles:\n")
    for country in countries:
        chat_log.insert(tk.END, f"- {country}\n")
    chat_log.insert(tk.END, "\n")

def clear_console():
    chat_log.delete(1.0, tk.END)

# Ruta del archivo Excel
file_path = 'D:\dev\Chatbot\Demografia.xlsx'

# Inicializar el chatbot con los datos del Excel
chatbot = ExcelChatbot(file_path)

# Crear la interfaz de usuario
root = tk.Tk()
root.title("Chatbot de Excel")

mainframe = ttk.Frame(root, padding="10 10 10 10")
mainframe.grid(column=0, row=0, sticky=(tk.W, tk.E, tk.N, tk.S))

root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

question_label = ttk.Label(mainframe, text="Ingrese su pregunta:")
question_label.grid(column=1, row=1, sticky=(tk.W, tk.E))

question_entry = ttk.Entry(mainframe, width=100)
question_entry.grid(column=1, row=2, sticky=(tk.W, tk.E))
question_entry.bind("<Return>", submit_question)

submit_button = ttk.Button(mainframe, text="Enviar", command=submit_question)
submit_button.grid(column=2, row=2, sticky=(tk.W, tk.E))

help_button = ttk.Button(mainframe, text="Listado de preguntas a realizar", command=show_help)
help_button.grid(column=1, row=3, sticky=(tk.W, tk.E))

show_questions_button = ttk.Button(mainframe, text="Historial de consultas", command=show_questions)
show_questions_button.grid(column=2, row=3, sticky=(tk.W, tk.E))

show_countries_button = ttk.Button(mainframe, text="Listado de paises disponibles", command=show_countries)
show_countries_button.grid(column=1, row=4, sticky=(tk.W, tk.E))

clear_button = ttk.Button(mainframe, text="Limpiar consola", command=clear_console)
clear_button.grid(column=2, row=4, sticky=(tk.W, tk.E))

chat_log = scrolledtext.ScrolledText(mainframe, width=100, height=20)
chat_log.grid(column=1, row=5, columnspan=2, sticky=(tk.W, tk.E))

for child in mainframe.winfo_children(): 
    child.grid_configure(padx=5, pady=5)

root.mainloop()
