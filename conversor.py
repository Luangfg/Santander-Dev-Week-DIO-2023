from os import path, listdir
from docx2pdf import convert
from tkinter import filedialog
import PySimpleGUI as sg



layout = [
    [sg.Text('Converta de Word para PDF')],
    [sg.Button('Converter')]

]    

janela = sg.Window('Conversor', layout)


while True:
    event, values = janela.read()
    if  event == sg.WIN_CLOSED:
        break
    
    try:
        pasta = filedialog.askdirectory()
        caminho_pasta = path.join(pasta)
        lista_arquivos = listdir(pasta)

        if lista_arquivos and caminho_pasta:
            for item in range(len(lista_arquivos)):
                arquivo = lista_arquivos[item]
                convert (caminho_pasta + '/' + arquivo)
            sg.popup('Conversão concluida.')
            
        else:
            sg.popup('Não existem arquivos na pasta indicada')
    
    except FileNotFoundError:
        pass