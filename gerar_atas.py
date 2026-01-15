import pandas as pd
from docx import Document
import os
from docx2pdf import convert
from datetime import datetime
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk

EXCEL = 'modelos/lista_cliente.xlsx'
MODELO = 'modelos/modelo_ata.docx'

PASTA_BASE = 'ATAS_GERADAS'
PASTA_DOCX = os.path.join(PASTA_BASE, 'DOCX')
PASTA_PDF = os.path.join(PASTA_BASE, 'PDF')

os.makedirs(PASTA_DOCX, exist_ok=True)
os.makedirs(PASTA_PDF, exist_ok=True)

df = pd.read_excel(EXCEL)

def substituir_texto(doc, mapa):
    for p in doc.paragraphs:
        for chave, valor in mapa.items():
            if chave in p.text:
                p.text = p.text.replace(chave, str(valor))

data_atual = datetime.now()

meses_pt = {
    1: 'janeiro', 2: 'fevereiro', 3: 'março', 4: 'abril',
    5: 'maio', 6: 'junho', 7: 'julho', 8: 'agosto',
    9: 'setembro', 10: 'outubro', 11: 'novembro', 12: 'dezembro'
}

dia = data_atual.day
mes = meses_pt[data_atual.month]
ano = data_atual.year
data_extenso = f"{dia:02d}/{data_atual.month:02d}/{ano}"

def iniciar():
    total = len(df)
    barra['maximum'] = total

    for index, row in enumerate(df.itertuples(), start=1):
        empresa = row.EMPRESA
        cnpj = row.CNPJ
        responsavel = row.RESPONSAVEL

        doc = Document(MODELO)

        dados = {
            '{{EMPRESA}}': empresa,
            '{{CNPJ}}': cnpj,
            '{{RESPONSAVEL}}': responsavel,
            '{{DIA}}': dia,
            '{{MES}}': mes,
            '{{ANO}}': ano,
            '{{DATA_EXTENSO}}': data_extenso
        }

        substituir_texto(doc, dados)

        nome_base = f"ATA - {empresa}"

        caminho_docx = os.path.join(PASTA_DOCX, nome_base + '.docx')
        caminho_pdf = os.path.join(PASTA_PDF, nome_base + '.pdf')

        doc.save(caminho_docx)

        try:
            convert(caminho_docx, caminho_pdf)
        except:
            pass

        barra['value'] = index
        status.config(text=f"Gerando {index} de {total}")
        janela.update()

    status.config(text="✅ Processo finalizado!")

janela = tk.Tk()
janela.title("Gerador de ATAs")
janela.geometry("420x480")
janela.resizable(False, False)

style = ttk.Style()
style.theme_use('default')
style.configure(
    "Red.Horizontal.TProgressbar",
    troughcolor='#f2f2f2',
    background='#c62828'
)

img = Image.open("logo.png")

largura_base = 200
proporcao = largura_base / img.width
altura = int(img.height * proporcao)

img = img.resize((largura_base, altura))
logo = ImageTk.PhotoImage(img)

lbl_logo = tk.Label(janela, image=logo)
lbl_logo.pack(pady=20)

btn = tk.Button(
    janela,
    text="Iniciar Geração",
    command=iniciar,
    font=("Arial", 12, "bold"),
    width=18
)
btn.pack(pady=10)

barra = ttk.Progressbar(
    janela,
    length=300,
    style="Red.Horizontal.TProgressbar"
)
barra.pack(pady=25)

status = tk.Label(
    janela,
    text="Aguardando...",
    font=("Arial", 10)
)
status.pack()

janela.mainloop()