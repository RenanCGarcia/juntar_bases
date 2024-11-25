import os
import sys
import mysql.connector

import pandas as pd
import tkinter as tk
import customtkinter as ctk

from tkinter import ttk
from datetime import datetime


class Functions:
    correct_columns = []
    bases = list()
    n_bases = 0
    total_lines = 0

    def chave_mestra(aplicacao):
        chave_mestra = {
            'host': 'chave-mestra.cpugooo8wo0m.us-east-1.rds.amazonaws.com',
            'database': 'chave_mestra',
            'user': 'usuario',
            'password': 'rr915200',
        }

        try:
            # Estabelecer a conexão com o MySQL
            sql = mysql.connector.connect(**chave_mestra)
            if sql.is_connected():
                # Criação de um cursor
                cursor = sql.cursor()
                consulta = "SELECT ativo FROM aplicacoes WHERE nome_aplicacao = %s"
                cursor.execute(consulta, (aplicacao,))
                resultado = cursor.fetchone()
                
                # Verificar se 'ativo' é 1 e retornar True, caso contrário, retornar False
                if resultado and resultado[0] == 1:
                    ativo = True
                else:
                    ativo = False
                return ativo

        except mysql.connector.Error as err:
            print(err)
            tk.messagebox.showinfo(title='Erro no Banco de Dados', message=f"Erro: não foi possível se conectar com o banco de dados AWS, entre em contato com o suporte.") # PRO USUÁRIO FINAL
            return False
        finally:
            cursor.close()
            sql.close()

    def center_window(self, window, h, w):
        height = h
        width = w

        height_window = window.winfo_screenheight()
        width_window = window.winfo_screenwidth()

        x = (width_window - height) // 2
        y = (height_window - width) // 2

        position = f"{height}x{width}+{x}+{y}"
        window.geometry(position)

    def reset(self):
        Functions.correct_columns.clear()
        Functions.bases.clear()
        Functions.n_bases = 0
        Functions.total_lines = 0
        #Enviar diretório para entry
        self.diretory_box.configure(placeholder_text="")
        for item in self.view_clients.get_children():
            self.view_clients.delete(item)

    def select_table(self):
        try:
            # Função que abre tela de selecionar arquivo
            self.diretory_table = ctk.filedialog.askopenfilename(
                title="Selecionar arquivo Excel",
                filetypes=[("Arquivos Excel", "*.xlsx")]
            )
            # Trata o diretório
            directory = os.path.dirname(self.diretory_table)
            name_base = os.path.basename(self.diretory_table)
            lines = 0
            if Functions.n_bases == 0:
                first_table = pd.read_excel(self.diretory_table)
                self.first_table_directory = directory
                self.first_table_name = name_base
                print('Primeiro: ', self.first_table_directory, self.first_table_name)
                lines = len(first_table)
                Functions.correct_columns = list(first_table.columns)
                Functions.bases.append(self.diretory_table)
                Functions.n_bases += 1
                Functions.total_lines += lines
            elif Functions.n_bases > 0:
                table = pd.read_excel(self.diretory_table)
                lines = len(table)
                table_columns = list(table.columns)
                if table_columns != Functions.correct_columns:
                    directory = 'Planilha Inválida'
                    name_base = 'Planilha Inválida'
                    tk.messagebox.showinfo(title='Erro ao ler Planilha', message='As colunas dessa planilha são diferentes da primeira, verifique e tente novamente.')
                elif table_columns == Functions.correct_columns:
                    #Enviar diretório para lista de bases
                    Functions.bases.append(self.diretory_table)
                    Functions.n_bases += 1
                    Functions.total_lines += lines
                    pass
            #Enviar diretório para entry
            self.diretory_box.configure(placeholder_text=f"{self.diretory_table}")
            #Enviar diretório para viewtable
            full_directory_tuple = (directory, name_base, lines)
            self.view_clients.insert("", "end", values=full_directory_tuple)
        except Exception as e:
            tk.messagebox.showinfo(title='Erro ao ler planilha', message=f'ERRO:\n{e}')

    def join(self):
        try:
            # Lista para armazenar DataFrames
            dataframes = []
            # Lê cada planilha e adiciona ao DataFrame
            for base in Functions.bases:
                df = pd.read_excel(base)
                dataframes.append(df)
            # Concatena todos os DataFrames
            df_completo = pd.concat(dataframes, ignore_index=True)
            
            # Criando diretorio do arquivo de saída
            atual_directory = os.path.dirname(sys.argv[0])
            data_hora_atual = datetime.now()
            data_hora_formatada = data_hora_atual.strftime('%d-%m-%Y %H-%M-%S')
            nome_saida = f'Junção - {self.first_table_name} etc - {data_hora_formatada}.xlsx'
            if not os.path.exists(os.path.join(self.first_table_directory)):
                    os.makedirs(os.path.join(self.first_table_directory))
            arquivo_saida = os.path.join(self.first_table_directory, nome_saida)
            print(arquivo_saida)
            # Salva o DataFrame resultante em um novo arquivo Excel
            df_completo.to_excel(arquivo_saida, index=False)
            tk.messagebox.showinfo(title='Concluído', message=f'Planilha retornada com sucesso!\nArquivo com {Functions.total_lines} linhas.')
        except Exception as e:
            tk.messagebox.showinfo(title='Erro ao juntar planilha', message=f'ERRO:\n{e}')


class App(Functions):
    def __init__(self):
        self.window = ctk.CTk()
        self.window_Properties()
        self.main_Frame()
        self.elements()

    def window_Properties(self):
        self.window.iconbitmap(os.path.join(os.path.dirname(sys.argv[0]),'icon.ico'))
        self.window.title("Juntar Bases")
        self.window.resizable(width=False, height=False)
        self.center_window(self.window, 1024, 768)

    def main_Frame(self):
        # Criação do frame principal
        self.mainFrame = ctk.CTkFrame(master=self.window)
        self.mainFrame.place(relx=0, rely=0, relwidth=1, relheight=1)

    def elements(self):
        # Título
        self.label_Title = ctk.CTkLabel(master=self.mainFrame, text='JUNTAR BASES', font=("Helvetica", 25))

        # Box que mostra o diretório da planilha selecionada
        self.diretory_box = ctk.CTkEntry(master=self.mainFrame)
        # Botão de selecionar tabela
        self.btn_select_table = ctk.CTkButton(
            master=self.mainFrame, 
            text="Selecionar Planilha"
            ,command=self.select_table
            )
        # Botão de RESET
        self.btn_reset = ctk.CTkButton(
            master=self.mainFrame, 
            text="RESET"
            ,command=self.reset
            )
        
        # Tela que mostra os clientes que leu
        self.view_clients = ttk.Treeview(master=self.mainFrame, height=19, column=("col1", "col2", "col3", "col4"))
        # CONFIGURAÇÃO DAS COLUNAS
        self.view_clients.heading("#0", text="")
        self.view_clients.heading("#1", text="DIRETÓRIO")
        self.view_clients.heading("#2", text="BASE")
        self.view_clients.heading("#3", text="LINHAS")
        self.view_clients.column("#0", width=0)
        self.view_clients.column("#1", width=500)
        self.view_clients.column("#2", width=300)
        self.view_clients.column("#3", width=100)
        # SCROLL
        self.scroll_view_table = tk.Scrollbar(master=self.mainFrame, orient="vertical")
        self.view_clients.configure(yscrollcommand=self.scroll_view_table.set)

        # Botão de baixar planilha
        self.btn_start = ctk.CTkButton(
            master=self.mainFrame, 
            text="Iniciar Junção"
            , command=self.join
            )
        self.btn_start.place(relx=0.425, rely=0.925)

        # PLACES
        self.label_Title.place(relx=0.380, rely=0.05)
        self.diretory_box.place(relx=0.150, rely=0.150, relwidth=0.625)
        self.btn_select_table.place(relx=0.800, rely=0.150)
        self.btn_reset.place(relx=0.800, rely=0.050)
        self.view_clients.place(relx=0.025, rely=0.235, relwidth=0.945, relheight=0.675)
        self.scroll_view_table.place(relx=0.95, rely=0.236, relheight=0.673)

    def run(self):
        self.window.mainloop()

if Functions.chave_mestra('UPTECH'):
    window = App()
    window.run()