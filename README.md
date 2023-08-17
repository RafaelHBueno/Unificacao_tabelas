import os
import pandas as pd
import openpyxl
import shutil
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
from threading import Thread

def copiar_numeros_de_serie(arquivo_entrada, arquivo_saida):
    cards_df = pd.read_excel(arquivo_entrada, sheet_name='Cards')

    linha_destino = 0  # Linha inicial para colar os números de série na aba Cards

    for linha_cards in range(cards_df.shape[0]):
        aba_origem = f'Tabela_{2 * linha_cards + 1}'  # Nome da aba ímpar correspondente
        tabela_origem_df = pd.read_excel(arquivo_entrada, sheet_name=aba_origem, dtype=str, engine='openpyxl')

        if 'Nº de Série' in tabela_origem_df.columns and tabela_origem_df.shape[0] > 0:
            numero_serie = tabela_origem_df['Nº de Série'].iloc[0]
            cards_df.at[linha_destino, 'Nº de Série'] = numero_serie  # Mantém como string
            linha_destino += 1

    cards_df.to_excel(arquivo_saida, index=False)




def excluir_planilha_em_pasta(pasta, nome_planilha):
    for arquivo in os.listdir(pasta):
        if arquivo.endswith('.xlsx'):
            caminho_arquivo = os.path.join(pasta, arquivo)

            wb = openpyxl.load_workbook(caminho_arquivo)

            if nome_planilha in wb.sheetnames:
                planilha = wb[nome_planilha]
                wb.remove(planilha)
                wb.save(caminho_arquivo)

class ExcelToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Cabeçalho Completo com Nº Série")
        self.root.geometry("500x250")

        self.label_caminho = tk.Label(root, text="Caminho da Pasta de Origem:")
        self.label_caminho.pack(pady=10)

        self.entry_caminho = tk.Entry(root, width=70)
        self.entry_caminho.pack(pady=5)

        self.button_caminho = tk.Button(root, text="Selecionar", command=self.select_folder)
        self.button_caminho.pack()

        self.progress_bar = ttk.Progressbar(root, length=300, mode='determinate')
        self.progress_bar.pack(pady=20)

        self.run_button = tk.Button(root, text="Executar Script", command=self.run_script)
        self.run_button.pack(pady=20)

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        self.entry_caminho.delete(0, tk.END)
        self.entry_caminho.insert(0, folder_path)

    def run_script(self):
        caminho = self.entry_caminho.get()

        if not os.path.exists(caminho):
            messagebox.showerror("Erro", "Caminho inválido.")
            return

        self.run_button.config(state=tk.DISABLED)
        self.run_button.update()

        thread = Thread(target=self.run_excel_tool, args=(caminho,))
        thread.start()

    def run_excel_tool(self, caminho):
        try:
            pasta_destino_intermediario = os.path.join(caminho, "Intermediarios")
            caminho_arquivo_final = os.path.join(caminho, "arquivo_final.xlsx")

            os.makedirs(pasta_destino_intermediario, exist_ok=True)
            total_progress = 128 * 2

            for i in range(1, 129):
                nome_arquivo_entrada = f'Pagina-{i:02d}.xlsx'
                caminho_entrada = os.path.join(caminho, nome_arquivo_entrada)
                nome_arquivo_saida = f'intermediario_Pagina-{i:02d}.xlsx'
                caminho_saida = os.path.join(pasta_destino_intermediario, nome_arquivo_saida)

                copiar_numeros_de_serie(caminho_entrada, caminho_saida)
                self.progress_bar['value'] = (i * 2) / total_progress * 100
                self.progress_bar.update()

            excluir_planilha_em_pasta(caminho, 'Cards')
            self.progress_bar['value'] = 50

            dados_compilados = pd.DataFrame()


            for i in range(1, 129):
                nome_arquivo_saida = f'intermediario_Pagina-{i:02d}.xlsx'
                caminho_arquivo_saida = os.path.join(pasta_destino_intermediario, nome_arquivo_saida)

                df = pd.read_excel(caminho_arquivo_saida, sheet_name='Sheet1', dtype=str, engine='openpyxl')
                dados_compilados = pd.concat([dados_compilados, df], ignore_index=True)
                self.progress_bar['value'] = 50 + (i * 2) / total_progress * 50
                self.progress_bar.update()

            dados_compilados.to_excel(caminho_arquivo_final, index=False)

            shutil.rmtree(pasta_destino_intermediario)

            messagebox.showinfo("Finalizado", "Planilha Gerada com Sucesso!")



        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

        self.run_button.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToolApp(root)
    root.mainloop()
