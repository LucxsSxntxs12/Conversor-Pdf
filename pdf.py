from os import path,listdir
from docx2pdf import convert
from tkinter import filedialog
import PySimpleGUI as sg

layout = [
    [sg.Text("Converta seu Word em PDF")],
    [sg.Button("Converter")]

]

janela = sg.Window('ConversorPDF',layout)


while True:
    event,values = janela.read()
    if event == sg.WIN_CLOSED:        
        break

    try:
        pasta = filedialog.askdirectory()
        caminho_pasta = path.join(pasta)
        lista_arquivos = listdir(pasta)

        if lista_arquivos and caminho_pasta:
            for item in range(len(lista_arquivos)):
                arquivo = lista_arquivos[item]
                convert(caminho_pasta +"/"+ arquivo)
            sg.popup("Conversao Concluida.")
        else:
            sg.popup(f"Nao existe arquivos .Docxs na Pasta{pasta}")
    except FileNotFoundError:
        pass
