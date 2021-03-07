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
frm.place(width=700, height=300, x=10)

tabela = ttk.Treeview(frm)
tabela.place(width=650, height=250, x=10, y=10)

rolagem_vertical = ttk.Scrollbar(frm, orient='vertical', command=tabela.yview)
rolagem_horizontal = ttk.Scrollbar(frm, orient='horizontal', command=tabela.xview)
tabela.configure(xscrollcommand=rolagem_horizontal.set, yscrollcommand=rolagem_vertical.set)
rolagem_horizontal.pack(side='bottom', fill='x')
rolagem_vertical.pack(side='right', fill='y')


# Frame dos widget: ================================================================================
frm02 = LabelFrame(janela, text='Ferramentas', bg='#f1f8e9', width=700, height=250)
frm02.place(x=20, y=320)

Label(frm02, text='Diretório do arquivo:', bg='#f1f8e9').place(x=20, y=20)
lbl_excel = Label(frm02, text='Nenhum arquivo selecionado', bg='#f1f8e9')
lbl_excel.place(x=20, y=50)

btn = Button(frm02, text='Arquivo', bg='#5efc82', command=lambda:folder())
btn.place(x=20, y=100)

btn2 = Button(frm02, text='Carregar', bg='#5efc82', command=lambda:load_arquivo())
btn2.place(x=150, y=100)


def folder():
    arquivo_excel = filedialog.askopenfilename(title='Selecione o arquivo:', filetypes=(('xlsx files','*xlsx*'),('All files','*.*')))

    lbl_excel['text'] = arquivo_excel
    return None
    

def load_arquivo():
    file_path = lbl_excel['text']
    try:
        excel_file = r"{}".format(file_path)
        carregar = pandas.read_excel(excel_file)
    
    # Erro se o arquivo escolido não for excel:
    except ValueError:
        messagebox.showerror('Aviso!', 'O arquivo escolhido é invalido!')
        return None
    
    # Erro se nenhum arquivo for selecionado:
    except FileNotFoundError:
        messagebox.showerror('Aviso!', 'Nenhum arquivo foi selecionado!')
        return None
    
    clear_data()
    # Comando para carregar o arquivo excel na tabela:
    tabela['column'] = list(excel_file)
    tabela['show'] = 'headings'

    for column in tabela['column']:
        tabela.heading(column, text=column)
   
    carregar_rows = carregar.to_numpy().tolist()
    for row in carregar_rows:
        tabela.insert('', 'end', value=row)

    return None

# Limpar a tabela a cada nova consulta
def clear_data():
    tabela.delete(*tabela.get_children())



janela.mainloop()
