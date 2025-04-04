import os
import psycopg2
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
from tkinter import Tk, Label, StringVar, Frame, messagebox, ttk, filedialog, PhotoImage 
from tkinter.font import Font


load_dotenv()

class DataQueryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Consulta de Dados")
        self.root.geometry("850x750")
        self.root.resizable(False, False)
        self.root.configure(bg="#f5f5f5") 
        self.setup_styles()
        self.create_widgets()
        
    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("TLabel", background="#f5f5f5", font=("Segoe UI", 11))
        self.style.configure("TButton", font=("Segoe UI", 11), padding=8)
        self.style.configure("TEntry", padding=6, font=("Segoe UI", 11))
        self.style.configure("TRadiobutton", background="#f5f5f5", font=("Segoe UI", 11))
        self.style.configure("Accent.TButton", background="#2c3e50", foreground="white", 
                           font=("Segoe UI", 11, "bold"), padding=10)
        self.style.map("TButton",
              background=[('active', '#2c3e50')],  
              foreground=[('active', 'black')])
              
        self.style.map("Accent.TButton",
            background=[('active', '#2980b9')],  
            foreground=[('active', 'white')])

        
    def create_widgets(self):
        header_frame = Frame(self.root, bg="#2c3e50", height=100)
        header_frame.pack(fill="x")
        
        title_font = Font(family="Segoe UI", size=20, weight="bold")
        Label(header_frame, text="Sistema de Consulta de Dados", font=title_font, 
              bg="#2c3e50", fg="white").place(relx=0.5, rely=0.5, anchor="center")
        
    
        main_frame = Frame(self.root, bg="#ffffff", bd=1, relief="solid")
        main_frame.pack(pady=20, padx=40, fill="both", expand=True)
        

        content_frame = Frame(main_frame, bg="#ffffff", padx=30, pady=30)
        content_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Opções de consulta
        Label(content_frame, text="Selecione o Tipo de Relatório:", bg="#ffffff", 
             font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 15))
        
        self.query_var = StringVar()
        queries = [
            ("1. Dados Brutos", "1"),
            ("2. Dados Agrupados por Ano", "2"),
            ("3. Dados Agrupados por Meio de Comunicação", "3"),
            ("4. Dados Brutos por CEP", "4"),
            ("5. Todos os Status por Protocolo", "5"),
            ("6. Top 50 Serviços Mais Solicitados", "6")
        ]
        
        for i, (text, value) in enumerate(queries):
            ttk.Radiobutton(content_frame, text=text, variable=self.query_var, 
                          value=value).grid(row=i+1, column=0, sticky="w", pady=2)

        self.period_frame = Frame(content_frame, bg="#ffffff")
        self.period_frame.grid(row=len(queries)+1, column=0, sticky="w", pady=15)
        

        self.date_frame = Frame(self.period_frame, bg="#ffffff")
        self.date_frame.pack(fill="x", expand=True)
        
        Label(self.date_frame, text="Período de Consulta:", bg="#ffffff", 
             font=("Segoe UI", 12, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        
        Label(self.date_frame, text="Data Inicial (AAAA-MM-DD):", bg="#ffffff").grid(row=1, column=0, sticky="w")
        self.start_date_var = StringVar(value="2023-01-01")
        self.start_date_entry = ttk.Entry(self.date_frame, textvariable=self.start_date_var, width=15)
        self.start_date_entry.grid(row=1, column=1, sticky="w", padx=5)
        
        Label(self.date_frame, text="Data Final (AAAA-MM-DD):", bg="#ffffff").grid(row=2, column=0, sticky="w")
        self.end_date_var = StringVar(value="2023-12-31")
        self.end_date_entry = ttk.Entry(self.date_frame, textvariable=self.end_date_var, width=15)
        self.end_date_entry.grid(row=2, column=1, sticky="w", padx=5)
        

        self.year_frame = Frame(self.period_frame, bg="#ffffff")
        
        Label(self.year_frame, text="Período de Consulta:", bg="#ffffff", 
             font=("Segoe UI", 12, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        
        Label(self.year_frame, text="Ano Inicial (AAAA):", bg="#ffffff").grid(row=1, column=0, sticky="w")
        self.start_year_var = StringVar(value="2023")
        self.start_year_entry = ttk.Entry(self.year_frame, textvariable=self.start_year_var, width=10)
        self.start_year_entry.grid(row=1, column=1, sticky="w", padx=5)
        
        Label(self.year_frame, text="Ano Final (AAAA):", bg="#ffffff").grid(row=2, column=0, sticky="w")
        self.end_year_var = StringVar(value="2023")
        self.end_year_entry = ttk.Entry(self.year_frame, textvariable=self.end_year_var, width=10)
        self.end_year_entry.grid(row=2, column=1, sticky="w", padx=5)

        self.id_frame = Frame(content_frame, bg="#ffffff")
        self.id_frame.grid(row=len(queries)+2, column=0, sticky="w", pady=10)
        
        self.id_label = Label(self.id_frame, text="IDs (separados por vírgula):", bg="#ffffff")
        self.id_label.grid(row=0, column=0, sticky="w")
        
        self.id_var = StringVar()
        self.id_entry = ttk.Entry(self.id_frame, textvariable=self.id_var, width=40)
        self.id_entry.grid(row=0, column=1, sticky="w", padx=5)

        btn_frame = Frame(content_frame, bg="#ffffff")
        btn_frame.grid(row=len(queries)+3, column=0, pady=20)
        
        self.query_btn = ttk.Button(btn_frame, text="Gerar Relatório", 
                                  command=self.execute_query, style="Accent.TButton")
        self.query_btn.pack(pady=10, ipadx=20)

        self.query_var.trace_add("write", self.update_fields)
        self.update_fields()
        

        footer_frame = Frame(self.root, bg="#f5f5f5", height=30)
        footer_frame.pack(fill="x", pady=(0, 10))
        Label(footer_frame, text="© 2023 Sistema HVM - Desenvolvido para Prefeitura", 
             background="#f5f5f5", font=("Segoe UI", 9)).pack()

        self.root.eval('tk::PlaceWindow . center')



    def update_fields(self, *args):
        query_type = self.query_var.get()
        
        if query_type in ['1', '2', '3']:
            self.id_label.config(text="IDs (separados por vírgula):")
            self.id_entry.delete(0, 'end')
            self.id_entry.insert(0, "951, 952, 953")
            self.id_entry.config(state='normal')
        elif query_type == '4':
            self.id_label.config(text="CEPs (separados por vírgula, sem traço):")
            self.id_entry.delete(0, 'end')
            self.id_entry.insert(0, "05421090, 01010001, 015020003")
            self.id_entry.config(state='normal')
        elif query_type == '5':
            self.id_label.config(text="Protocolos (separados por vírgula):")
            self.id_entry.delete(0, 'end')
            self.id_entry.insert(0, "21654476, 13781361, 31433835")
            self.id_entry.config(state='normal')
        elif query_type == '6':
            self.id_label.config(text="(Não necessário para esta consulta)")
            self.id_entry.delete(0, 'end')
            self.id_entry.config(state='disabled')
        else:
            self.id_entry.config(state='normal')
        


        if query_type == '2':
            self.date_frame.pack_forget()
            self.year_frame.pack(fill="x", expand=True)
        elif query_type == '5' or  query_type == '4':
            self.year_frame.pack_forget()
            self.date_frame.pack_forget()  
        else:
            self.year_frame.pack_forget()
            self.date_frame.pack(fill="x", expand=True)



    def validate_inputs(self):
        query_type = self.query_var.get()
        
        if not query_type:
            messagebox.showwarning("Opção Vazia", "Por favor, selecione um tipo de relatório.")
            return False

        if query_type == '2':  
            try:
                start_year = int(self.start_year_var.get())
                end_year = int(self.end_year_var.get())
                
                if end_year < start_year:
                    messagebox.showwarning("Ano Inválido", "O ano final deve ser posterior ou igual ao ano inicial.")
                    return False
                
                self.start_date_var.set(f"{start_year}-01-01")
                self.end_date_var.set(f"{end_year}-12-31")
            except ValueError:
                messagebox.showerror("Erro de Ano", "Os anos devem ser números válidos (ex: 2023).")
                return False
        elif query_type in ['1', '3', '6']:
            try:
                start_date = datetime.strptime(self.start_date_var.get(), "%Y-%m-%d")
                end_date = datetime.strptime(self.end_date_var.get(), "%Y-%m-%d")
                
                if end_date < start_date:
                    messagebox.showwarning("Data Inválida", "A data final deve ser posterior à data inicial.")
                    return False
            except ValueError:
                messagebox.showerror("Erro de Data", "As datas devem estar no formato YYYY-MM-DD.")
                return False

        if query_type != '6':
            ids = self.id_var.get().strip()
            if not ids:
                messagebox.showwarning("Entrada Vazia", "Por favor, informe os valores solicitados.")
                return False
            
            try:
                if query_type in ['1', '2', '3']:
                    [int(id.strip()) for id in ids.split(',')]
                elif query_type == '4':
                    [int(cep.strip()) for cep in ids.split(',')]
                elif query_type == '5':
                    [int(proto.strip()) if proto.strip().isdigit() else proto.strip() for proto in ids.split(',')]
            except ValueError:
                messagebox.showerror("Erro de Formato", "Valores devem ser números inteiros separados por vírgula.")
                return False
        
        return True
    


    def conectar_banco(self):
        try:
            connection = psycopg2.connect(
                host=os.getenv('DB_HOST'),
                port=os.getenv('DB_PORT'),
                dbname=os.getenv('DB_NAME'),
                user=os.getenv('DB_USER'),
                password=os.getenv('DB_PASSWORD')
            )
            return connection
        except psycopg2.OperationalError as e:
            messagebox.showerror("Erro de Conexão", f"Erro ao conectar ao banco de dados: {e}")
            return None
        

    
    def executar_query(self, cursor, query):
        cursor.execute(query)
        rows = cursor.fetchall()
        colunas = [desc[0] for desc in cursor.description]
        return pd.DataFrame(rows, columns=colunas)
    
    
    
    def salvar_excel(self, dados, nome):
        default_dir = os.path.join(os.path.expanduser("~"), "Documents", "Dados_projeto_ascom")
        os.makedirs(default_dir, exist_ok=True)
        
        if self.query_var.get() == '2':
            start = self.start_year_var.get()
            end = self.end_year_var.get()
        else:
            start = self.start_date_var.get().replace("-", "")
            end = self.end_date_var.get().replace("-", "")
        
        filename = f"dados_{nome}_{start}_{end}.xlsx"
        initialfile = os.path.join(default_dir, filename)
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=filename,
            initialdir=default_dir
        )
        
        if filepath:
            try:
                dados.to_excel(filepath, index=False, engine="xlsxwriter")
                os.startfile(filepath)
                return True
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao salvar arquivo: {e}")
                return False
        return False
    


    def execute_query(self):
        if not self.validate_inputs():
            return
        
        query_type = self.query_var.get()
        connection = self.conectar_banco()
        
        if connection is None:
            return
        
        try:
            cursor = connection.cursor()
            dados = None
            nome = ""
            
            if query_type == '1':
                nome = "dados_brutos"
                query = f"""
                    SELECT  
                        tb001.id_servico AS "Id Servico",
                        INITCAP(tb001.descriçao_servico) AS "Serviço",
                        INITCAP(tb001.nome_prefeitura) AS "Solicitação Prefeitura",            
                        tb001.numero_do_protocolo AS "Protocolo",
                        TO_CHAR(tb001.data_cadastro, 'DD-MM-YYYY') AS "Data Cadastro",
                        TO_CHAR(tb001.data_finalizacao, 'DD-MM-YYYY') AS "Data Finalização",
                        CASE 
                            WHEN tb001.data_finalizacao IS NOT NULL THEN 
                                ROUND(EXTRACT(EPOCH FROM (tb001.data_finalizacao - tb001.data_cadastro)) / 86400) 
                            ELSE NULL 
                        END AS "Tempo Atendimento (dias)",
                        de_status,
                        id_status as "Id Status"
                    FROM extracoes.tabela1 tb001
                    WHERE 
                        tb001.data_cadastro BETWEEN '{self.start_date_var.get()} 00:00:00' AND '{self.end_date_var.get()} 23:59:59'
                        AND tb001.id_servico IN ({self.id_var.get()})
                        AND tb001.id_status IN (0, 1, 2, 5);
                """
                dados = self.executar_query(cursor, query)



            elif query_type == '2':
                nome = "dados_agrupado_ano"
                query = f"""
                    SELECT 
                        EXTRACT(YEAR FROM data_cadastro) AS Ano,
                        INITCAP(descriçao_servico) AS Servico,
                        COUNT(DISTINCT CASE WHEN id_status = 1 THEN numero_do_protocolo  AS "Quantidade Recebidas",
                        COUNT(DISTINCT CASE WHEN id_status in (2, 5)  THEN numero_do_protocolo END) AS "Quantidade Finalizadas"
                    FROM 
                        extracoes.tabela1
                    WHERE 
                        id_servico IN ({self.id_var.get()})
                        AND data_cadastro BETWEEN '{self.start_date_var.get()} 00:00:00' AND '{self.end_date_var.get()} 23:59:59'
                    GROUP BY 
                        EXTRACT(YEAR FROM data_cadastro), descriçao_servico;
                """
                dados = self.executar_query(cursor, query)



            elif query_type == '3':
                nome = "dados_agrupados_meio_comunicacao"
                query = f"""
                    SELECT           
                        CASE
                            WHEN meio_comunicacao = 'CRM' THEN 'Central Telefonica'
                            WHEN meio_comunicacao = 'PORTAL' THEN 'Portal 156'
                            ELSE INITCAP(meio_comunicacao)
                        END AS "Meio_Comunicacao",
                        INITCAP(descriçao_servico) AS Servico,
                        COUNT(DISTINCT CASE WHEN id_status = 1 THEN numero_do_protocolo  AS "Quantidade Recebidas",
                        COUNT(DISTINCT CASE WHEN id_status in (2, 5) THEN numero_do_protocolo END) AS "Quantidade Finalizadas"
                    FROM 
                        extracoes.tabela1
                    WHERE 
                        id_servico IN ({self.id_var.get()})
                        AND data_cadastro BETWEEN '{self.start_date_var.get()} 00:00:00' AND '{self.end_date_var.get()} 23:59:59'
                        AND meio_comunicacao IN ('CRM', 'PORTAL', 'CHATBOT', 'MOBILE')
                    GROUP BY 
                        meio_comunicacao, descriçao_servico;
                """
                dados = self.executar_query(cursor, query)



            elif query_type == '4':
                nome = "dados_brutos_cep"
                ceps_formatados = ", ".join([f"'{cep.strip()}'" for cep in self.id_var.get().split(",")])
                query = f"""
                    SELECT  
                        tb001.numero_cep as CEP,
                        tb001.numero_protocolo AS "Protocolo",
                        tb001.id_servico AS "Id Servico",
                        INITCAP(tb001.descriçao_servico) AS "Serviço",
                        INITCAP(tb001.nome_prefeitura) AS "Solicitação Prefeitura",                        
                        TO_CHAR(tb001.data_cadastro, 'DD-MM-YYYY') AS "Data Cadastro",
                        TO_CHAR(tb001.data_finalizacao, 'DD-MM-YYYY') AS "Data Finalização",
                        ROUND(EXTRACT(EPOCH FROM (tb001.data_finalizacao - tb001.data_cadastro)) / 86400) AS "TMA tempo médio atendimento (em dias)",
                        tb001.de_status AS "Status"
                    FROM extracoes.tabela1 tb001
                    WHERE 
                        tb001.numero_cep IN ({ceps_formatados})
                    ORDER BY
                    tb001.numero_cep DESC;
                """
                dados = self.executar_query(cursor, query)

            elif query_type == '5':
                nome = "dados_protocolo"
                query = f"""
                    SELECT 
                        tb001.numero_do_protocolo AS Protocolo,
                        tb001.descriçao_servico AS Servico,
                        tb001.meio_comunicacao AS "Meio de Comunicação",
                        TO_CHAR(data_cadastro, 'DD-MM-YYYY') AS "Data de Abertura",
                        TO_CHAR(data_finalizacao, 'DD-MM-YYYY') AS "Data de Finalização",
                        de_status AS Status,
                        tb001.nome_prefeitura AS "Prefeitura",
                        tb002.descricao_reposta AS Resposta
                    FROM extracoes.tabela1 tb001
                    LEFT JOIN extracoes.tabela2 tb002
                        ON tb001.numero_do_protocolo = tb002.numero_do_protocolo
                    WHERE tb001.numero_do_protocolo IN ({self.id_var.get()})
                    ORDER BY
                        tb001.numero_do_protocolo
                """
                dados = self.executar_query(cursor, query)

            elif query_type == '6':
                nome = "top50_servicos"
                query = f"""
                    SELECT 
                        INITCAP(tb001.descriçao_servico) AS "Serviço",
                        COUNT(DISTINCT tb001.numero_do_protocolo) AS "Solicitações Abertas",
                        COUNT(DISTINCT CASE WHEN tb001.id_status in (2, 5) THEN tb001.numero_do_protocolo END) AS "Solicitações Finalizadas",
                        ROUND(AVG(EXTRACT(EPOCH FROM (tb001.data_finalizacao - tb001.data_cadastro)) / 86400)) 
                            AS "TMA tempo médio atendimento (em dias)"
                    FROM 
                        extracoes.tabela1 tb001
                    WHERE 
                        tb001.data_cadastro BETWEEN '{self.start_date_var.get()} 00:00:00' AND '{self.end_date_var.get()} 23:59:59'
                    GROUP BY 
                        tb001.descriçao_servico
                    ORDER BY 
                        "Solicitações Abertas" DESC
                    LIMIT 50;
                """
                dados = self.executar_query(cursor, query)

            if dados is not None:
                if dados.empty:
                    messagebox.showinfo("Sem Dados", "Nenhum dado encontrado para os critérios informados.")
                else:
                    self.salvar_excel(dados, nome)

        except Exception as error:
            messagebox.showerror("Erro", f"Ocorreu um erro durante a consulta:\n{error}")
        finally:
            if 'cursor' in locals():
                cursor.close()
            if 'connection' in locals():
                connection.close()


if __name__ == "__main__":
    load_dotenv()
    root = Tk()
    app = DataQueryApp(root)
    root.mainloop()