###### Bibliotecas ######
import os
import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter import messagebox
import time
import datetime
import getpass
import sqlite3
from sqlite3 import Error
import pandas as pd

###### Parâmetros ######
db = 'db_cadastro_clientes.db'
hoje = datetime.datetime.now()
data = datetime.datetime(hoje.year, hoje.month, hoje.day).strftime('%d/%m/%Y')
mes = datetime.datetime(hoje.year, hoje.month, hoje.day).strftime('%B')

###### Config. App ######
app=tk.Tk()
app.title("Cadastro de Clientes")
app.geometry("300x450")
app.configure(background="Black")

###### Listas ######
motivo = ["Alteração Cadastral","Alteração de E-mail","Alteração de Telefone","Consulta","Credenciamento","Descredenciamento"]
login_user = getpass.getuser()

###### Funções ######
def criar_banco():
    conexao = sqlite3.connect('db_cadastro_clientes.db')

    c = conexao.cursor()

    c.execute("""CREATE TABLE Cadastro_Clientes (
        data text,
        mes text,
        codigo_cliente text,
        nome_cliente text,
        email text,
        telefone text,
        protocolo text,
        prazo_protocolo text,
        motivo text,
        login text
        )""")

    conexao.commit()
    conexao.close()
    tk.messagebox.showinfo(title="Confirmação",message="Banco de Dados Criado!")

def tabular_dados():
    telefoneValue = cb_telefone.get()
    if len(telefoneValue) !=11:
        tk.messagebox.showerror(title="Ops!!!",message="Número do telefone incorreto, ele deve conter 11 dígitos!!!")
    else:
        conexao = sqlite3.connect('db_cadastro_clientes.db')

        c = conexao.cursor()

        c.execute("INSERT INTO Cadastro_Clientes VALUES (:data,:mes,:codigo_cliente,:nome_cliente,:email,:telefone,:protocolo,:prazo_protocolo,:motivo,:login)",
                {
                    'data': cb_data.get(),
                    'mes': cb_mes.get(),
                    'codigo_cliente': cb_cod.get(),
                    'nome_cliente': cb_user.get(),
                    'email': cb_email.get(),
                    'telefone': cb_telefone.get(),
                    'protocolo': cb_prot.get(),
                    'prazo_protocolo': cb_prazo_prot.get(),
                    'motivo': cb_motivo_prot.get(),
                    'login': cb_login.get()
                })

        conexao.commit()
        conexao.close()
        tk.messagebox.showinfo(title="Confirmação",message="Dados Inseridos!")
        cb_cod.delete(0, END)
        cb_user.delete(0, END)
        cb_email.delete(0, END)
        cb_telefone.delete(0, END)
        cb_prot.delete(0, END)
        cb_prazo_prot.delete(0, END)
        cb_motivo_prot.delete(0, END)

def exporta_clientes():
    conexao = sqlite3.connect('db_cadastro_clientes.db')
    c = conexao.cursor()

    c.execute("SELECT *, oid FROM Cadastro_Clientes")
    dados_tabulados = c.fetchall()

    dados_tabulados=pd.DataFrame(dados_tabulados,columns=['data','mes','codigo_cliente','nome_cliente','email','telefone','protocolo','prazo_protocolo','motivo','login','Id_banco'])
    dados_tabulados.to_excel('cadastro_clientes.xlsx', index=False)

    conexao.commit()

    conexao.close()
    tk.messagebox.showinfo(title="Confirmação",message="Dados Exportados!")

###### Campos ######
frame = Frame(
    app,
    padx=5,
    pady=5
)
frame.pack(pady=5)

lb_data = tk.Label(frame,text="Data",background="Red",foreground="White")
lb_data.grid(row=1, column=0, padx=5, pady=5)
cb_data = ttk.Combobox(frame,values=data)
cb_data.grid(row=1, column=1, padx=5, pady=5)
cb_data.set(data)


lb_mes = tk.Label(frame,text="Mês",background="Red",foreground="White")
lb_mes.grid(row=2, column=0, padx=5, pady=5)

cb_mes = ttk.Combobox(frame,values=mes)
cb_mes.grid(row=2, column=1, padx=5, pady=5)
cb_mes.set(mes)


lb_cod = tk.Label(frame,text="Código do Cliente",background="Red",foreground="White")
lb_cod.grid(row=3, column=0, padx=5, pady=5)

cb_cod = tk.Entry(frame,width=23)
cb_cod.grid(row=3, column=1, padx=5, pady=5)


lb_user = tk.Label(frame,text="Nome do Cliente",background="Red",foreground="White")
lb_user.grid(row=4, column=0, padx=5, pady=5)

cb_user = tk.Entry(frame,width=23)
cb_user.grid(row=4, column=1, padx=5, pady=5)

lb_email = tk.Label(frame,text="E-mail",background="Red",foreground="White")
lb_email.grid(row=5, column=0, padx=5, pady=5)

cb_email = tk.Entry(frame,width=23)
cb_email.grid(row=5, column=1, padx=5, pady=5)


lb_telefone = tk.Label(frame,text="Telefone",background="Red",foreground="White")
lb_telefone.grid(row=6, column=0, padx=5, pady=5)

cb_telefone = tk.Entry(frame,width=23)
cb_telefone.grid(row=6, column=1, padx=5, pady=5)


lb_prot = tk.Label(frame,text="Número do Protocolo",background="Red",foreground="White")
lb_prot.grid(row=7, column=0, padx=5, pady=5)

cb_prot = tk.Entry(frame,width=23)
cb_prot.grid(row=7, column=1, padx=5, pady=5)


lb_prazo_prot = tk.Label(frame,text="Prazo do protocolo",background="Red",foreground="White")
lb_prazo_prot.grid(row=8, column=0, padx=5, pady=5)

cb_prazo_prot = tk.Entry(frame,width=23)
cb_prazo_prot.grid(row=8, column=1, padx=5, pady=5)


lb_motivo_prot = tk.Label(frame,text="Motivo",background="Red",foreground="White")
lb_motivo_prot.grid(row=9, column=0, padx=5, pady=5)

cb_motivo_prot = ttk.Combobox(frame,values=motivo)
cb_motivo_prot.grid(row=9, column=1, padx=5, pady=5)


lb_login = tk.Label(frame,text="Login do Agente",background="Red",foreground="White")
lb_login.grid(row=10, column=0, padx=5, pady=5)

cb_login = ttk.Combobox(frame,width=20)
cb_login.grid(row=10, column=1, padx=5, pady=5)
cb_login.set(login_user)


###### Botão ######
botao_tabular = tk.Button(frame,text='Tabular Dados', command=tabular_dados)
botao_tabular.grid(row=11, column=0,columnspan=2, padx=5, pady=5 , ipadx = 80)

botao_limpar_planilha = tk.Button(frame,text='Criar Banco de Dados', command=criar_banco)
botao_limpar_planilha.grid(row=12, column=0,columnspan=2, padx=5, pady=5 , ipadx = 61)

botao_limpar_planilha = tk.Button(frame,text='Exportar para Excel', command=exporta_clientes)
botao_limpar_planilha.grid(row=13, column=0,columnspan=2, padx=5, pady=5 , ipadx = 68)

app.mainloop()