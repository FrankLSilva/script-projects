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
    webbrowser.open("https://docs.google.com/document/d/1rPiZJzjoKoakPUGepZG0wHfQfJBqajC536CI-mCdprw/edit?pli=1")
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
    text = "\nResumo Operacional | Evento\n" \
           f"{nome_evento}\n\n\n" \
           "Pré-Evento:\n" \
           f"  \u2022 Visita Técnica:{empty}{empty}{visita_técnica}\n" \
           f"  \u2022 Entrega da Caçamba:{empty}. . . . . . {entrega_cacamba}\n" \
           f"  \u2022 Entrega dos Equipamentos:{empty}. {entrega_equips}\n" \
           f"  \u2022 Implementação dos Equipamentos:. . . . . . . {implem_equips}\n\n\n" \
           "Equipamentos em comodato:\n" \
           f"  \u2022 Sacos de lixo (pacote): _______\n" \
           f"  \u2022 Lixeira (unidade): _______\n" \
           f"  \u2022 Contêiner (unidade): _______\n" \
           f"  \u2022 Carro Coletor (unidade): _______\n\n\n" \
           "Resumo Coletas e Desmobilização:\n" \
           f"  \u2022 Coleta dos resíduos: \n     \u2022 1º Coleta no dia:{empty}{coleta_residuos}\n\n" \
           f"  \u2022 Retirada dos equipamentos:{empty}. . {retirada_equips}\n" \
           f"  \u2022 Horário de Acompanhamento: {empty}{horário_acomp}\n\n" \
           f"  \u2022 Acompanhamento de Coleta e Desmobilização:\n     \u2022 {coleta_desmob}\n\n\n" \
           "Pós-Evento:\n" \
           f"  \u2022 Relatório de impacto: {relatório_impact}\n" \
           f"  \u2022 Envio de faturamento: {envio_fatura}\n\n\n" \
           f"Responsável Operacional do Evento: \n{responsável}\n"

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

        # Formatação --------> "Resumo Operacional | Evento"
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': 1,
                    'endIndex': find_line_start(text, "Resumo Operacional | Evento") + len(
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
                    'endIndex': find_line_start(text, nome_evento) + len(nome_evento) + 1
                },
                'paragraphStyle': {
                    'alignment': 'CENTER'
                },
                'fields': 'alignment'
            }
        },  # Paragraph style

        # Formatação --------> "Pré-Evento:"
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': find_line_start(text, "Pré-Evento"),
                    'endIndex': find_line_start(text, "Pré-Evento") + len("Pré-Evento") + 2
                },
                'textStyle': {
                    'bold': True
                },
                'fields': 'bold'
            }
        },# Text style

        # Formatação --------> "Equipamentos em comodato:"
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': find_line_start(text, "Equipamentos em comodato"),
                    'endIndex': find_line_start(text, "Equipamentos em comodato") + len(
                        "Equipamentos em comodato") + 2
                },
                'textStyle': {
                    'bold': True
                },
                'fields': 'bold'
            }
        },# Text style

        # Formatação --------> "Resumo Coletas e Desmobilização:"
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': find_line_start(text, "Resumo Coletas e Desmobilização"),
                    'endIndex': find_line_start(text, "Resumo Coletas e Desmobilização") + len(
                        "Resumo Coletas e Desmobilização") + 2
                },
                'textStyle': {
                    'bold': True
                },
                'fields': 'bold'
            }
        }, # Text style

        # Formatação --------> "Pós-Evento:"
        {
            'updateTextStyle': {
                'range': {
                    'startIndex': find_line_start(text, "Pós-Evento"),
                    'endIndex': find_line_start(text, "Pós-Evento") + len("Pós-Evento") + 2
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
root.title("Musa | Resumo Operacional")
root.geometry("600x490")

# Logo Musa
icone_png = Image.open("musa_logo.png")
icone_tk = ImageTk.PhotoImage(icone_png)
root.iconphoto(True, icone_tk)

# Botão para selecionar o arquivo
file_button = tk.Button(root, text="Selecionar Arquivo", command=select_file)
file_button.pack(pady=10, padx=10)

# Botão para abrir o link específico
link_button = tk.Button(root, text="Abrir Resumo Operacional", command=open_link, bg="#D3D6DB")
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

    nome_evento = csv_data["data"][0][1]

    # Pré Evento
    visita_técnica = csv_data["data"][0][4]
    entrega_cacamba = csv_data["data"][0][5]
    entrega_equips = csv_data["data"][0][5]
    implem_equips = csv_data["data"][0][5]

    # Equipamentos em comodato
    sacos_lixo = "__________"

    # Resumo coletas e desmobilização
    coleta_residuos = (f"{csv_data["data"][0][6]} - Entre __ até __")
    retirada_equips = csv_data["data"][0][7]
    coleta_desmob = "Nome: __________ - CPF: __________"
    horário_acomp = "Dia __ - Entre __ até __"

    # Pós Evento
    relatório_impact = "Previsto para 7 dias úteis após o evento."
    envio_fatura = ("7 dias úteis após o evento. Caso haja necessidade de especificidades "
                    "na Nota Fiscal, informar previamente antes do envio das notas.")

    # Comunicado
    responsável = "____________"
    comunicado = ("Por estarmos operando em São Paulo, uma cidade dinâmica, "
                  "quaisquer alterações nos horários ou datas devem ser alinhadas com pelo menos "
                  "24 horas de antecedência para garantir a excelência na operação.")

    # ------------------------------------------------------------- MAIN + GOOGLE AUTHS
    # ID do documento e credenciais do Google Cloud
    SERVICE_ACCOUNT_FILE = 'credentials.json'
    DOCUMENT_ID = '1rPiZJzjoKoakPUGepZG0wHfQfJBqajC536CI-mCdprw'

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