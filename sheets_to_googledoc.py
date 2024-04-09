from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import csv
import tkinter as tk
from tkinter import filedialog
import webbrowser
from PIL import Image, ImageTk

# line to install
# pyinstaller --onefile --noconsole main.py

# ------------------------------------------------------------- FUNÇÕES / Data
def select_file():
    global csv_data
    file_path = filedialog.askopenfilename(filetypes=[("Arquivos CSV", "*.csv")])
    if file_path:
        selected_file.set(file_path)
        csv_data = read_csv_file(file_path)
        display_csv_info(csv_data)
    else:
        selected_file.set("Nenhum arquivo selecionado")
        csv_data = None
def read_csv_file(file_path):
    global csv_data
    with open(file_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        columns = next(reader)
        data = list(reader)
    csv_data = {"columns": columns, "data": data}
    return csv_data
def display_csv_info(csv_data):
    try:
        columns = csv_data["columns"]
        values = csv_data["data"][0]  # Obtém os valores da primeira linha (após o cabeçalho)

        info_text.config(state=tk.NORMAL)  # Habilita a edição do texto
        info_text.delete(1.0, tk.END)  # Limpa o texto atual
        info_text.insert(tk.END, "Informações do Arquivo CSV:\n\n")

        for column, value in zip(columns, values):
            info_text.insert(tk.END, f"{column}: {value}\n")

        info_text.config(state=tk.DISABLED)  # Impede a edição do texto
    except Exception as e:
        print("Erro ao exibir informações do arquivo:", e)
def open_link():
    webbrowser.open("link here")
def cancel():
    global cancelado
    cancelado = True
    root.destroy()
def run_go():
    global run
    run = True

# ------------------------------------------------------------- FUNÇÕES / Text
def fontUpdate(end, font):
    requests = [
        {
            "updateTextStyle": {
                "range": {
                    "startIndex": 1,  # Início do texto a ser estilizado
                    "endIndex": end  # Fim do texto a ser estilizado
                },
                "textStyle": {
                    "weightedFontFamily": {
                        "fontFamily": font,  # Nome da família da fonte
                        "weight": 400  # Peso da fonte (400 é normal, 700 é negrito)
                    },
                    "fontSize": {
                        "magnitude": 12,  # Tamanho da fonte
                        "unit": "PT"  # Unidade do tamanho da fonte
                    }
                },
                "fields": "weightedFontFamily,fontSize"  # Campos a serem atualizados
            }
        }
    ]
    return requests
def find_line_start(text, line):
    lines = text.split('\n')
    for i, l in enumerate(lines):
        if l.strip().startswith(line):
            return sum(len(lines[j]) + 1 for j in range(i))
def textDeleteInsert(end_index):
    text = "\nExample Text Title\n" \
           f"{variable}\n\n\n" \
           "Pré-Evento:\n" \
           f"  \u2022 Field:{empty}{empty}{variable}\n" \
           f"  \u2022 Field:{empty}. . . . . . {variable}\n" \
           f"  \u2022 Field:{empty}. {variable}\n" \
           f"  \u2022 Field:. . . . . . . {variable}\n\n\n" \
           "Title:\n" \
           f"  \u2022 Field: _______\n" \
           f"  \u2022 Field: _______\n" \
           f"  \u2022 Field: _______\n" \
           f"  \u2022 Field: _______\n\n\n" \
           "Title:\n" \
           f"  \u2022 Field: \n     \u2022 text:{empty}{variable}\n\n" \
           f"  \u2022 Field:{empty}. . {variable}\n" \
           f"  \u2022 Field: {empty}{variable}\n\n" \
           f"  \u2022 Field:\n     \u2022 {variable}\n\n\n" \
           "Title:\n" \
           f"  \u2022 Field: {variable}\n" \
           f"  \u2022 Field: {variable}\n\n\n" \
           f"Title: \n{variable}\n"

    requests = [
        {'deleteContentRange': {'range': {
            'startIndex': 1,
            'endIndex': end_index}}},
        {
            'insertText': {
                'location': {
                    'index': 1,
                },
                'text': text
            }
        }, # Text insert

        # Formatação --------> "Text to search"
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': 1,
                    'endIndex': find_line_start(text, "Text to search") + len(
                        "Resumo Operacional | Evento") + 1
                },
                'textStyle': {
                    'fontSize': {
                        'magnitude': 16,  # Tamanho da fonte
                        'unit': 'PT'
                    },
                    'bold': True
                },
                'fields': 'fontSize, bold'
            }
        },  # Text style
        {
            'updateParagraphStyle': {
                'range': {
                    'startIndex': 1,
                    'endIndex': find_line_start(text, variable) + len(variable) + 1
                },
                'paragraphStyle': {
                    'alignment': 'CENTER'
                },
                'fields': 'alignment'
            }
        },  # Paragraph style

        # Formatação --------> "Text to search"
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': find_line_start(text, "Text to search"),
                    'endIndex': find_line_start(text, "Text to search") + len("Text to search") + 2
                },
                'textStyle': {
                    'bold': True
                },
                'fields': 'bold'
            }
        },# Text style

        # Formatação --------> "Text to search:"
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': find_line_start(text, "Text to search"),
                    'endIndex': find_line_start(text, "Text to search") + len(
                        "Equipamentos em comodato") + 2
                },
                'textStyle': {
                    'bold': True
                },
                'fields': 'bold'
            }
        },# Text style

        # Formatação --------> "Text to search"
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': find_line_start(text, "Text to search"),
                    'endIndex': find_line_start(text, "Text to search") + len(
                        "Text to search") + 2
                },
                'textStyle': {
                    'bold': True
                },
                'fields': 'bold'
            }
        }, # Text style

        # Formatação --------> "Text to search"
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': find_line_start(text, "Text to search"),
                    'endIndex': find_line_start(text, "Text to search") + len("Text to search") + 2
                },
                'textStyle': {
                    'bold': True
                },
                'fields': 'bold'
            }
        }, # Text style
    ]
    return requests

# ------------------------------------------------------------- USER INTERFACE
cancelado = False
run = False

# Cria a janela principal
root = tk.Tk()
root.title("Title")
root.geometry("600x490")

# Logo Musa
icone_png = Image.open("logo.png")
icone_tk = ImageTk.PhotoImage(icone_png)
root.iconphoto(True, icone_tk)

# Botão para selecionar o arquivo
file_button = tk.Button(root, text="Selecionar Arquivo", command=select_file)
file_button.pack(pady=10, padx=10)

# Botão para abrir o link específico
link_button = tk.Button(root, text="Abrir link", command=open_link, bg="#D3D6DB")
link_button.pack(pady=10, padx=10)

# Botão para cancelar
cancel_button = tk.Button(root, text="Cancelar", command=cancel, bg="#F05454", fg="White")
cancel_button.pack(pady=10, padx=10)

# Botão para realizar
run_button = tk.Button(root, text="Go!", command=run_go, bg="#006a39", fg="White")
run_button.pack(pady=10, padx=10)

# Variável para armazenar o caminho do arquivo selecionado
selected_file = tk.StringVar()
file_label = tk.Label(root, textvariable=selected_file)
file_label.pack(pady=10)

# Componente de texto para exibir as informações do arquivo
info_text = tk.Text(root, font=("Helvetica", 10), state=tk.DISABLED)
info_text.pack(expand=True, fill="both", padx=10, pady=10)

# Executa o loop principal da interface gráfica
root.mainloop()

# ------------------------------------------------------------- RUN
if cancelado:
    print("\nCancelado")
elif run:
    # ------------------------------------------------------------- VARIÁVEIS TEXTUAIS
    empty = " . . . . . . . . . . . "

    variable = csv_data["data"][0][1]

    # Title
    variable = csv_data["data"][0][4]
    variable = csv_data["data"][0][5]
    variable = csv_data["data"][0][5]
    variable = csv_data["data"][0][5]

    # Title
    variable = "__________"

    # Title
    variable = (f"{csv_data["data"][0][6]} - Entre __ até __")
    variable = csv_data["data"][0][7]
    variable = "Nome: __________ - CPF: __________"
    variable = "Dia __ - Entre __ até __"

    # Title
    variable = "Text"
    variable = ("text")

    # Title
    variable = "____________"
    variable = ("Text")

    # ------------------------------------------------------------- MAIN + GOOGLE AUTHS
    # ID do documento e credenciais do Google Cloud
    SERVICE_ACCOUNT_FILE = 'credentials.json'
    DOCUMENT_ID = 'document id'

    # Inicialização e credenciais para o Doc Service
    SCOPES = ['https://www.googleapis.com/auth/documents']
    CREDENTIALS = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('docs', 'v1', credentials=CREDENTIALS)

    # Abrir o Doc
    doc = service.documents().get(documentId=DOCUMENT_ID).execute()

    # Checar tamanho do conteúdo para limpeza das informações
    doc_content = doc.get('body').get('content')
    end_index = doc_content[-1].get('endIndex', 1) - 1

    # Requerimentos para atualização de texto
    txt = textDeleteInsert(end_index)
    font = fontUpdate(end=end_index, font="Helvetica")

    # Execução de requerimentos
    service.documents().batchUpdate(documentId=DOCUMENT_ID, body={'requests': font}).execute()
    service.documents().batchUpdate(documentId=DOCUMENT_ID, body={'requests': txt}).execute()

    print(f'\nDocumento atualizado') #
else:
    print("\nCancelado")
