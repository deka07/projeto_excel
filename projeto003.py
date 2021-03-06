import os
import sys
import pandas
from tkinter import *
from tkinter import ttk, filedialog, messagebox

janela = Tk()
janela.geometry('750x600')
janela.title('Cactus Pandas')
janela.configure(bg='#f1f8e9')


# Frame da tabela: =================================================================================
frm = LabelFrame(janela, text='Dados do Excel', bg='#f1f8e9', width=700, height=300)
frm.place(x=20, y=20)

tabela = ttk.Treeview(frm)
tabela.place(width=650, height=250, x=20, y=10)

def load_arquivo():
    abrir = filedialog.askopenfilename(title='Selecione o arquivo:', filetypes=(('xlsx files','*xlsx*'),('All files','*.*')))
    carregar = pandas.read_excel(abrir)
    
    # Comando para carregar o arquivo excel na tabela:
    tabela['column'] = list(carregar)
    tabela['show'] = 'headings'
    for column in tabela['column']:
        tabela.heading(column, text=column)
    
    carregar_rows = carregar.to_numpy().tolist()
    for row in carregar_rows:
        tabela.insert('', 'end', value=row)

# Frame dos widget: ================================================================================
frm02 = LabelFrame(janela, text='Ferramentas', bg='#f1f8e9', width=700, height=250)
frm02.place(x=20, y=320)


btn = Button(frm02, text='Abrir', bg='#5efc82', command=load_arquivo)
btn.place(x=20, y=20)



janela.mainloop()
