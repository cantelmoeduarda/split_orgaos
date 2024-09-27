#Split


#importações
import os
import tkinter as tk
from tkinter import messagebox
from tkinter import *
from PIL import Image,ImageTk
import openpyxl as xl
from openpyxl.styles import NamedStyle, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
import sys
import pandas as pd
from tkinter import simpledialog
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from tkinter import ttk

def main():

    caminho_pasta = var_caminho_pasta.get()
    caminho_valor_faturado = var_caminho_valor_faturado.get()


    #Lendo o excel em que gostaria de realizar a separação, fazendo com que o python leia apenas a aba "Apuração do Faturamento":
    try:
        df = pd.read_excel(caminho_valor_faturado, sheet_name='Valor Faturado')
    except FileNotFoundError:
        messagebox.showerror("Erro", "Selecione todos os arquivos solicitados.")
        raise Exception('Um dos arquivos não foram encontrados') #isso aqui vai parar o código
    except ValueError:
        messagebox.showerror("Erro","Verifique se os arquivos estão corretos")
        raise Exception('Arquivo provavelmente errado')
        
    #Iniciando a barra de progresso   
    pt=ttk.Progressbar(janela,variable=varBarra,maximum=100)
    pt.place(x=8,y=270, width=380,height=15)
    
    #Realizando o group by, que separará as linhas do excel por cada cliente:
    grouped = df.groupby('Órgão/Entidade')
    total_orgaos = len(grouped)
    
    count = 0
    novo_diretorio = os.path.join(caminho_pasta,'Valor Faturado por órgãos')
    if not os.path.exists(novo_diretorio):
        os.makedirs(novo_diretorio)

    for orgao, data in grouped:
        orgao_nome = f"{orgao}.xlsx".replace("/", "_").replace("\\", "_").replace(":", "_").replace("*", "_").replace("?", "_").replace('"', "_").replace("<", "_").replace(">", "_").replace("|", "_") #isso aqui substitui possiveis caracteres em que o windows não permite que salve os arquivos
        orgao_path = os.path.join(novo_diretorio, orgao_nome)
        data.to_excel(orgao_path, index=False, sheet_name='Valor Faturado')
        
        #barra progresso
        count += 1
        progresso = (count/total_orgaos)*100
        varBarra.set(progresso)
        janela.update_idletasks()

    
    
    messagebox.showinfo('Sucesso',f'Arquivo Valor Faturado por órgãos salvo na pasta "{os.path.basename(caminho_pasta)}" com sucesso \n\nTotal de arquivos criados: {count}')
    progress_bar.destroy()



# A partir daqui o código é referente à janela 

 
def selecionar_valor_faturado():
    caminho_valor_faturado = askopenfilename(title='Selecione o arquivo Valor Faturado')
    var_caminho_valor_faturado.set(caminho_valor_faturado)    
    if caminho_valor_faturado:
        label_valor_faturado_selecionado['text'] = f"* Arquivo selecionado: {os.path.basename(caminho_valor_faturado)}" 
    
def selecionar_pasta_final():
    caminho_pasta = askdirectory(title='Selecione o arquivo Valor Faturado')
    var_caminho_pasta.set(caminho_pasta)    
    if caminho_pasta:
        label_pasta_selecionada['text'] = f"* Pasta selecionada: {os.path.basename(caminho_pasta)}" 
        
janela = tk.Tk()
janela.geometry('400x300')
janela.title('Valor Faturado') 
janela.resizable(False,False) #pra não conseguirem mudar o tamanho da caixa
texto = tk.Label(janela, text='Selecione o arquivo Valor Faturado e a Pasta de Destino:')
texto.grid(column=1, row=0, padx=10, pady=10, sticky='w')


janela.grid_columnconfigure(0, weight=50)
janela.grid_columnconfigure(1, weight=50)

var_caminho_valor_faturado = tk.StringVar()
botao_selecionar_valor_faturado = tk.Button(janela, text="Selecionar Valor Faturado", command=selecionar_valor_faturado)
botao_selecionar_valor_faturado.grid(row=1, column=1, padx=10, pady=10, sticky='nsew')
label_valor_faturado_selecionado = tk.Label(janela, text='* Nenhum arquivo selecionado',fg='blue')
label_valor_faturado_selecionado.grid(row=2, column=1, sticky='nsew')

var_caminho_pasta = tk.StringVar()
botao_selecionararquivo = tk.Button(janela, text="Selecionar Pasta Destino", command=selecionar_pasta_final)
botao_selecionararquivo.grid(row=4, column=1, padx=10, pady=10, sticky='nsew')
label_pasta_selecionada = tk.Label(janela, text='* Nenhuma pasta selecionada',fg='blue')
label_pasta_selecionada.grid(row=5, column=1, sticky='nsew')



botao_processar = tk.Button(janela, text='Processar', command=main)
botao_processar.grid(column=1, row=6, columnspan=1, padx=10, pady=35, ipady=6, sticky='nsew')

#barra de progresso
varBarra = DoubleVar()
varBarra.set(0)




# Tratando as imagens que farão parte do botão:
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path,relative_path)

# icone triagulo
image_path = resource_path("triangulo.png")
imagem = Image.open(image_path)

imagem = imagem.resize((32,32),Image.LANCZOS)
icone = ImageTk.PhotoImage(imagem)
janela.iconphoto(True, icone)

# botao de ajuda
def ajuda():
    messagebox.showinfo("Informações importantes",
                        "A pasta destino é aquela que conterá o novo arquivo processado. \n\nAs Planilhas separadas pelos órgãos retornarão em uma pasta chamada 'Valor Faturado por órgãos', logo, não é necessária a criação de uma pasta nova. ")
    
icon_path = resource_path("botao_de_ajuda_transparente.png")
help_icon = Image.open(icon_path)
help_icon = help_icon.resize((22, 22), Image.LANCZOS)
help_icon = ImageTk.PhotoImage(help_icon)
help_button = tk.Button(janela, image=help_icon, command=ajuda, borderwidth=0, bg='#f0f0f0', activebackground='#f0f0f0')
help_button.grid(row=6, column=4, padx=10, sticky='e')


janela.mainloop()
