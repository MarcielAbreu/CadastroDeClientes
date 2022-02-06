###### Bibliotecas ######
import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter import messagebox
import datetime
import getpass
import pandas as pd
from pathlib import Path

###### Parâmetros ######
file = r'C:/Users/' + getpass.getuser() + '/Desktop/Cadastro.xlsx'
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
def tabular_dados(): #Função tabular dados no excel
    telefoneValue = cb_telefone.get()
    if len(telefoneValue) !=11:
        tk.messagebox.showerror(title="Ops!!!",message="Número do telefone incorreto, ele deve conter 11 dígitos!!!")
    else:
        df = pd.read_excel('C:/Users/' + getpass.getuser() + '/Desktop/Cadastro.xlsx')
        df = df.append({'Data': cb_data.get(),'Mês':cb_mes.get(),'Cód Cliente':cb_cod.get(),'User':cb_user.get(),'e-mail':cb_email.get(),
                        'Telefone':cb_telefone.get(),'Número do Protocolo':cb_prot.get(),'Prazo do Protocolo':cb_prazo_prot.get(),
                        'Motivo':cb_motivo_prot.get(),'Login do Agente':cb_login.get()}, ignore_index=True)
        df.to_excel('C:/Users/' + getpass.getuser() + '/Desktop/Cadastro.xlsx', index=False)
        tk.messagebox.showinfo(title="Confirmação",message="Dados Inseridos!")
        cb_cod.delete(0, END)
        cb_user.delete(0, END)
        cb_email.delete(0, END)
        cb_telefone.delete(0, END)
        cb_prot.delete(0, END)
        cb_prazo_prot.delete(0, END)
        cb_motivo_prot.delete(0, END)

def limpar_planilha():
    df = pd.read_excel('C:/Users/' + getpass.getuser() + '/Desktop/Cadastro.xlsx')
    df.drop(["Data", "Mês", "Cód Cliente", "User", "e-mail", "Telefone", "Número do Protocolo", "Prazo do Protocolo", "Motivo", "Login do Agente"], axis=1, inplace = True)
    df.to_excel('C:/Users/' + getpass.getuser() + '/Desktop/Cadastro.xlsx', index=False)
    tk.messagebox.showinfo(title="Confirmação",message="Planilha Limpa!")
    
def criar_planilha():
    out_path = 'C:/Users/' + getpass.getuser() + '/Desktop/Cadastro.xlsx'
    df = pd.ExcelWriter(out_path, engine='xlsxwriter')
    df.save()
    tk.messagebox.showinfo(title="Confirmação",message="Planilha Criada!")

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

botao_limpar_planilha = tk.Button(frame,text='Criar Planilha', command=criar_planilha)
botao_limpar_planilha.grid(row=12, column=0,columnspan=2, padx=5, pady=5 , ipadx = 82)

botao_limpar_planilha = tk.Button(frame,text='Limpar Planilha', command=limpar_planilha)
botao_limpar_planilha.grid(row=13, column=0,columnspan=2, padx=5, pady=5 , ipadx = 77)

app.mainloop()