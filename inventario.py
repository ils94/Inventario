# coding=utf8

import csv
import json
import os
import pathlib
import sys
import threading
from datetime import date, datetime
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk

import psycopg2

inserir_query = """INSERT INTO INVENTARIO (DESCRICAO, QUANTIDADE, LOCAL, DATA_HORA, MODIFICADO) VALUES (%s, %s, %s, %s, %s)"""

alterar_query = """UPDATE INVENTARIO SET DESCRICAO = %s, QUANTIDADE = %s, LOCAL = %s, DATA_HORA = %s, MODIFICADO = %s WHERE ID = %s"""

carregar_query = """SELECT * FROM INVENTARIO ORDER BY ID DESC"""

pesquisar_query = "SELECT * FROM INVENTARIO WHERE DESCRICAO ILIKE %s OR QUANTIDADE ILIKE %s OR LOCAL ILIKE %s OR DATA_HORA ILIKE %s OR MODIFICADO ILIKE %s"

user_home = "Z:/" + str(os.getlogin())

json_arquivo = pathlib.Path(user_home + "/inventario/cfg.json")

banco = None
timer = None
credenciais = None

id = ""

root = Tk()

janela_width = 1024
janela_height = 764

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

x = (screen_width / 2) - (janela_width / 2)
y = (screen_height / 2) - (janela_height / 2)

root.geometry("1024x764+" + str(int(x)) + "+" + str(int(y)))
root.title("Inventário SEPAT")
root.iconbitmap('icones/inventario.ico')
root.resizable(False, False)


def data_hora():
    dia = date.today()
    dia_string = str(dia).split("-")
    dia_format = dia_string[2] + '/' + dia_string[1] + '/' + dia_string[0]
    dia_format_ext = dia_format.replace("/01/", "/jan/") \
        .replace("/02/", "/fev/") \
        .replace("/03/", "/mar/") \
        .replace("/04/", "/abr/") \
        .replace("/05/", "/mai/") \
        .replace("/06/", "/jun/") \
        .replace("/07/", "/jul/") \
        .replace("/08/", "/ago/") \
        .replace("/09/", "/set/") \
        .replace("/10/", "/out/") \
        .replace("/11/", "/nov/") \
        .replace("/12/", "/dez/")

    hora = datetime.now()
    hora_atual = hora.strftime("%H:%M:%S")

    return hora_atual, dia_format_ext.upper()


def mensagens_de_erro(e):
    messagebox.showerror("Erro", e)


def criar_json(dbName, dbUser, dbPass, dbHost, dbPort):
    global credenciais

    try:
        data = {}

        data['dbName'] = dbName
        data['dbUser'] = dbUser
        data['dbPass'] = dbPass
        data['dbHost'] = dbHost
        data['dbPort'] = dbPort

        json_data = json.dumps(data)

        salvar_json = open(user_home + "/inventario/cfg.json", "w")
        salvar_json.write(str(json_data))
        salvar_json.close()

        messagebox.showinfo("Salvo", "As credenciais foram salvas com sucesso.")

        credenciais.destroy()

        conectar()

    except Exception as e:
        mensagens_de_erro(e)


def salvar_credenciais():
    global credenciais

    credenciais = Toplevel(root)
    credenciais.geometry("300x160+" + str(int(x)) + "+" + str(int(y)))
    credenciais.resizable(False, False)
    credenciais.iconbitmap("icones/cadeado.ico")
    credenciais.title("Salvar Credenciais")
    credenciais.attributes("-topmost", True)

    def salvar():
        if entry_dbname.get() == "" or entry_dbuser.get() == "" or entry_dbpass.get() == "" or entry_dbhost.get() == "" or entry_dbport.get() == "":
            mensagens_de_erro("É necessário preencher todos os campos.")
        else:
            criar_json(entry_dbname.get(), entry_dbuser.get(), entry_dbpass.get(),
                       entry_dbhost.get(), entry_dbport.get())

    labelframe_credenciais = LabelFrame(credenciais, text="Inserir dados")
    labelframe_credenciais.pack(fill=X, side=TOP, padx=5)

    label_dbname = Label(labelframe_credenciais, text="DB NAME:", width=10, height=1, anchor=W)

    entry_dbname = Entry(labelframe_credenciais, width=30)

    label_dbuser = Label(labelframe_credenciais, text="DB USER:", width=10, height=1, anchor=W)

    entry_dbuser = Entry(labelframe_credenciais, width=30)

    label_dbpass = Label(labelframe_credenciais, text="DB PASS:", width=10, height=1, anchor=W)

    entry_dbpass = Entry(labelframe_credenciais, width=30)

    label_dbhost = Label(labelframe_credenciais, text="DB HOST:", width=10, height=1, anchor=W)

    entry_dbhost = Entry(labelframe_credenciais, width=30)

    label_dbport = Label(labelframe_credenciais, text="DB PORT:", width=10, height=1, anchor=W)

    entry_dbport = Entry(labelframe_credenciais, width=30)

    button_cre_salvar = Button(credenciais, text="Salvar", width=10, height=1, command=salvar)

    label_dbname.grid(row=1, column=0)
    label_dbuser.grid(row=2, column=0)
    label_dbpass.grid(row=3, column=0)
    label_dbhost.grid(row=4, column=0)
    label_dbport.grid(row=5, column=0)

    entry_dbname.grid(row=1, column=1)
    entry_dbuser.grid(row=2, column=1)
    entry_dbpass.grid(row=3, column=1)
    entry_dbhost.grid(row=4, column=1)
    entry_dbport.grid(row=5, column=1)

    button_cre_salvar.pack(side=LEFT, padx=5)

    try:
        if json_arquivo.exists():
            with open(json_arquivo) as js:
                dados = json.load(js)

                entry_dbname.insert(0, dados["dbName"])
                entry_dbuser.insert(0, dados["dbUser"])
                entry_dbpass.insert(0, dados["dbPass"])
                entry_dbhost.insert(0, dados["dbHost"])
                entry_dbport.insert(0, dados["dbPort"])
        else:
            entry_dbport.insert(0, "5432")
    except Exception as e:
        mensagens_de_erro(e)

    credenciais.mainloop()


def multithreading(funcao):
    threading.Thread(target=funcao).start()


def usuario_inativo():
    global banco

    if banco is not None:
        banco.close()

        pergunta = messagebox.askyesno("Desconectado",
                                       "Você foi desconectado por inatividade.\n\nDeseja se reconectar com o banco?")

        if pergunta:
            multithreading(conectar)
            reset_timer()
        else:
            sys.exit()
    else:
        reset_timer()


def reset_timer(event=None):
    global timer

    if timer is not None:
        root.after_cancel(timer)

    timer = root.after(60000, usuario_inativo)


def conectar():
    global banco

    try:
        if json_arquivo.exists():
            with open(json_arquivo) as js:
                dados = json.load(js)

                DB_NAME = dados["dbName"]
                DB_USER = dados["dbUser"]
                DB_PASS = dados["dbPass"]
                DB_HOST = dados["dbHost"]
                DB_PORT = dados["dbPort"]

            banco = psycopg2.connect(dbname=DB_NAME, user=DB_USER, password=DB_PASS, host=DB_HOST, port=DB_PORT)

            carregar_inventario()

            js.close()
    except Exception as e:
        mensagens_de_erro(e)


def banco_queries(**kwargs):
    global banco
    global id

    variaveis = kwargs.get("variaveis")
    modificar = kwargs.get("modificar")
    carregar = kwargs.get("carregar")
    pesquisar = kwargs.get("pesquisar")

    try:
        if modificar:
            cursor = banco.cursor()
            cursor.execute(modificar, variaveis)
            banco.commit()
            carregar_inventario()
        if carregar:
            cursor = banco.cursor()
            cursor.execute(carregar)
            id = ""
            return cursor
        if pesquisar:
            cursor = banco.cursor()
            cursor.execute(pesquisar, variaveis)
            id = ""
            return cursor
    except Exception as e:
        mensagens_de_erro(e)


def items(event):
    global id

    entry_descricao.delete(0, END)
    entry_quantidade.delete(0, END)
    entry_local.delete(0, END)

    id = tv.item(tv.selection())["values"][0]

    entry_descricao.insert(0, tv.item(tv.selection())["values"][1])
    entry_quantidade.insert(0, tv.item(tv.selection())["values"][2])
    entry_local.insert(0, tv.item(tv.selection())["values"][3])


def exportar_banco_para_planilha():
    try:
        arquivo = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Arquivo CSV", "*.csv")])

        with open(arquivo, "w", newline='', encoding="utf-8") as myfile:
            csvwriter = csv.writer(myfile, delimiter=",")

            csvwriter.writerow(
                ["DESCRIÇÃO", "QUANTIDADE", "LOCAL", "ÚLTIMA MODIFICAÇÃO", "MODIFICADO POR"])

            for row_id in tv.get_children():
                row = tv.item(row_id)["values"]
                csvwriter.writerow(row[1:6])

        pergunta = messagebox.showinfo("Concluído", "Abrir arquivo?")

        if pergunta:
            os.startfile(arquivo)
    except Exception as e:
        mensagens_de_erro(e)


def inserir_visualizador(cursor):
    try:
        tv.delete(*tv.get_children())
        for row in cursor:
            tv.insert("", index="end", values=(
                row[0], row[1], row[2], row[3], row[4], row[5]))
    except TypeError:
        pass


def inserir():
    if entry_descricao.get() == "" or entry_quantidade.get() == "" or entry_local.get() == "":
        mensagens_de_erro("Todos os campos precisam estar preenchidos.")
    else:

        dataHora = data_hora()

        pergunta = messagebox.askyesno("Inserir", "Inserir uma nova entrada no banco de dados?")

        if pergunta:
            variaveis = (entry_descricao.get().upper(), entry_quantidade.get().upper(), entry_local.get().upper(),
                         dataHora[1].upper() + " - " + dataHora[0], str(os.getlogin()))

            banco_queries(modificar=inserir_query, variaveis=variaveis)

            entry_descricao.delete(0, END)
            entry_quantidade.delete(0, END)
            entry_local.delete(0, END)

            carregar_inventario()


def alterar():
    global id

    if entry_descricao.get() == "" or entry_quantidade.get() == "" or entry_local.get() == "":
        mensagens_de_erro("Todos os campos precisam estar preenchidos.")
    elif id == "":
        mensagens_de_erro("Selecione um item na Inventario primeiro.")
    else:

        dataHora = data_hora()

        pergunta = messagebox.askyesno("Alterar", "Alterar a entrada com o ID: " + str(id) + '?')

        if pergunta:
            variaveis = (entry_descricao.get().upper(), entry_quantidade.get().upper(), entry_local.get().upper(),
                         dataHora[1].upper() + " - " + dataHora[0], str(os.getlogin()), id)

            banco_queries(modificar=alterar_query, variaveis=variaveis)

            carregar_inventario()


def carregar_inventario():
    cursor = banco_queries(carregar=carregar_query)
    inserir_visualizador(cursor)


def pesquisar_inventario(event):
    variaveis = (
        "%" + entry_pesquisar.get() + "%",
        "%" + entry_pesquisar.get() + "%",
        "%" + entry_pesquisar.get() + "%",
        "%" + entry_pesquisar.get() + "%",
        "%" + entry_pesquisar.get() + "%")

    cursor = banco_queries(pesquisar=pesquisar_query, variaveis=variaveis)
    inserir_visualizador(cursor)


menu_bar = Menu(root)

root_menu = Menu(menu_bar, tearoff=0)

root_menu.add_command(label="Exportar Inventário para .CSV", command=exportar_banco_para_planilha)
root_menu.add_command(label="Salvar Credenciais", command=salvar_credenciais)

menu_bar.add_cascade(label="Menu", menu=root_menu)

root.config(menu=menu_bar)

frame1 = LabelFrame(root)
frame1.pack(fill=X, padx=5, pady=5)

label_descricao = Label(frame1, text="Descrição: ", width=10, anchor=W)
entry_descricao = Entry(frame1, width=150)

label_quantidade = Label(frame1, text="Quantidade: ", width=10, anchor=W)
entry_quantidade = Entry(frame1, width=150)

label_local = Label(frame1, text="Local: ", width=10, anchor=W)
entry_local = Entry(frame1, width=150)

label_descricao.grid(row=0, column=0, padx=5, pady=2)
entry_descricao.grid(row=0, column=1, padx=5, pady=2)
label_quantidade.grid(row=1, column=0, padx=5, pady=2)
entry_quantidade.grid(row=1, column=1, padx=5, pady=2)
label_local.grid(row=2, column=0, padx=5, pady=2)
entry_local.grid(row=2, column=1, padx=5, pady=2)

frame2 = Frame(root)
frame2.pack(fill=X, padx=5, pady=5)

button_novo = Button(frame2, text="Cadastrar", width="15", height="1", command=inserir)
button_novo.pack(side=LEFT)

button_carregar = Button(frame2, text="Carregar", width="15", height="1", command=carregar_inventario)
button_carregar.pack(side=LEFT, padx=5)

button_alterar = Button(frame2, text="Alterar", width="15", height="1", command=alterar)
button_alterar.pack(side=RIGHT)

frame3 = LabelFrame(root, text="Visualização do Inventário")

xsb = ttk.Scrollbar(frame3, orient=HORIZONTAL)
xsb.pack(side=BOTTOM, fill=X)

ysb = ttk.Scrollbar(frame3, orient=VERTICAL)
ysb.pack(side=RIGHT, fill=Y)

tv = ttk.Treeview(frame3, height=25, selectmode='browse', show='headings', xscrollcommand=xsb.set,
                  yscrollcommand=ysb.set)

tv.bind("<<TreeviewSelect>>", items)

xsb.config(command=tv.xview)
ysb.config(command=tv.yview)

tv['columns'] = (
    "ID", "Descrição", "Quantidade", "Local", "Última Modificação", "Modificado Por")

tv.column("#0", width=2, minwidth=1)
tv.column("ID", width=50, minwidth=0)
tv.column("Descrição", width=500, minwidth=499)
tv.column("Quantidade", width=100, minwidth=99)
tv.column("Local", width=150, minwidth=149)
tv.column("Última Modificação", width=150, minwidth=149)
tv.column("Modificado Por", width=150, minwidth=149)

tv.heading("#0", text="", anchor=W)
tv.heading("ID", text="ID", anchor=W)
tv.heading("Descrição", text="Descrição", anchor=W)
tv.heading("Quantidade", text="Quantidade", anchor=W)
tv.heading("Local", text="Local", anchor=W)
tv.heading("Última Modificação", text="Última Modificação", anchor=W)
tv.heading("Modificado Por", text="Modificado Por", anchor=W)

tv.pack(padx=5, pady=5, fill=X)
frame3.pack(fill=X, padx=5, pady=5)

frame4 = Frame(root)

label_pesquisar = Label(frame4, text='Pesquisar:', width=10, height=1, anchor=W)
label_pesquisar.pack(side=LEFT)

entry_pesquisar = Entry(frame4, width=30)
entry_pesquisar.pack(side=LEFT)
entry_pesquisar.bind('<Return>', pesquisar_inventario)

frame4.pack(fill=X, padx=5, pady=5)

inventario_cfg = pathlib.Path(user_home + "/inventario")

if not inventario_cfg.exists():
    os.makedirs(user_home + "/inventario")

multithreading(conectar)

root.bind_all("<Any-KeyPress>", reset_timer)
root.bind_all("<Any-ButtonPress>", reset_timer)

root.mainloop()
