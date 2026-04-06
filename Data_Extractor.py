import win32com.client
import pythoncom
import os
import pandas as pd
from datetime import datetime
import pdfplumber
import re
import pytesseract
import cv2
import numpy as np
from PIL import Image
import fitz
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import unicodedata
import sqlite3

# --- CONFIGURAÇÃO OCR ---
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\esantan3\OneDrive - The Mosaic Company\Área de Trabalho\Projetos\tesseract-main\tesseract.exe'

# --- CAMINHOS ---
PASTA_ANEXOS = r"C:\Users\esantan3\OneDrive - The Mosaic Company\Área de Trabalho\Projetos\Exchange Tax Invoice\Folder XML_PDF"
ARQUIVO_FINAL = r"C:\Users\esantan3\OneDrive - The Mosaic Company\Área de Trabalho\Projetos\Exchange Tax Invoice\Final Report\Relatorio_Controle_Trocas.xlsx"
BANCO_SQLITE = r"C:\Users\esantan3\OneDrive - The Mosaic Company\Área de Trabalho\Projetos\Exchange Tax Invoice\Final Report\Banco_Controle_Trocas.db"

# --- LISTA DE ASSISTENTES ---
LISTA_ASSISTENTES = [
    "Todos os Assistentes",
    "Fernando Rodrigues",
    "Gustavo Chaves",
    "João Teixeira",
    "José Viana",
    "Vitória Nunes"
]

class AppMosaicMaster:
    def __init__(self, root):
        self.root = root
        self.root.title("Mosaic Tool v10.0 - Motor SQLite (Anti-Duplicação)")
        self.root.geometry("750x600")
        self.root.configure(bg="#f4f6f9")
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TLabel", background="#f4f6f9", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10, "bold"), background="#004a8d", foreground="white")
        style.map("TButton", background=[("active", "#003366")])

        titulo = tk.Label(root, text="Monitoramento de Troca de Notas", font=("Segoe UI", 14, "bold"), bg="#f4f6f9", fg="#004a8d")
        titulo.pack(pady=(20, 10))
        
        frame_filtros = tk.Frame(root, bg="#f4f6f9")
        frame_filtros.pack(pady=10)

        ttk.Label(frame_filtros, text="Data Início:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.ent_ini = ttk.Entry(frame_filtros, width=15)
        self.ent_ini.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.ent_ini.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(frame_filtros, text="Data Fim:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.ent_fim = ttk.Entry(frame_filtros, width=15)
        self.ent_fim.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.ent_fim.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(frame_filtros, text="Filtrar Assistente:").grid(row=1, column=0, padx=5, pady=15, sticky="e")
        self.combo_assistente = ttk.Combobox(frame_filtros, values=LISTA_ASSISTENTES, state="readonly", width=30)
        self.combo_assistente.current(0)
        self.combo_assistente.grid(row=1, column=1, columnspan=3, padx=5, pady=15, sticky="w")

        self.btn = tk.Button(root, text="EXTRAIR DADOS", command=self.start, bg="#004a8d", fg="white", font=("Segoe UI", 11, "bold"), width=25, relief="flat")
        self.btn.pack(pady=15)

        self.progress = ttk.Progressbar(root, length=650, mode="determinate")
        self.progress.pack(pady=5)

        self.log_txt = tk.Text(root, height=12, width=85, font=("Consolas", 9), bg="#1e1e1e", fg="#4af626", relief="flat")
        self.log_txt.pack(pady=15, padx=20)

    def log(self, msg):
        self.log_txt.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_txt.see(tk.END)

    def remover_acentos(self, texto):
        if not texto: return ""
        return unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8').lower()

    def extrair_placa(self, texto):
        if not texto: return None
        busca = re.search(r'[A-Za-z]{3}[-\s]?[0-9][A-Za-z0-9][0-9]{2}', texto)
        if busca:
            return busca.group(0).replace("-", "").replace(" ", "").upper()
        return None

    def ler_xml(self, caminho):
        d = {"NF": "N/D", "Material": "N/D", "Vol": "0", "Emitente": "N/D", "CFOP": "N/D"}
        try:
            with open(caminho, 'r', encoding='utf-8', errors='ignore') as f:
                conteudo = f.read()

            conteudo_limpo = re.sub(r'(</?)[a-zA-Z0-9\-]+:', r'\1', conteudo)

            nf_m = re.search(r'<nNF>(\d+)</nNF>', conteudo_limpo)
            if nf_m: d["NF"] = nf_m.group(1)

            mat_m = re.search(r'<xProd>([^<]+)</xProd>', conteudo_limpo)
            if mat_m: 
                d["Material"] = mat_m.group(1).replace('&quot;', '"').replace('&amp;', '&').strip()

            emit_m = re.search(r'<emit>.*?<xNome>([^<]+)</xNome>', conteudo_limpo, re.DOTALL)
            if emit_m: 
                d["Emitente"] = emit_m.group(1).replace('&quot;', '"').replace('&amp;', '&').strip()
                
            cfop_m = re.search(r'<CFOP>\s*([^<]+)\s*</CFOP>', conteudo_limpo, re.IGNORECASE)
            if cfop_m:
                d["CFOP"] = cfop_m.group(1).strip()

            pesol_m = re.search(r'<pesoL>\s*([\d\.,]+)\s*</pesoL>', conteudo_limpo, re.IGNORECASE)
            if pesol_m:
                val = pesol_m.group(1).replace(',', '.') 
                try:
                    d["Vol"] = str(int(float(val)))
                except ValueError:
                    d["Vol"] = val.split('.')[0]
            else:
                qvol_m = re.search(r'<qVol>\s*([\d\.,]+)\s*</qVol>', conteudo_limpo, re.IGNORECASE)
                if qvol_m:
                    d["Vol"] = qvol_m.group(1).split('.')[0]

            return d
        except Exception as e:
            return d

    def ler_pdf_limpo(self, caminho):
        d = {"NF": "N/D", "Material": "N/D", "Vol": "0", "Emitente": "N/D", "CFOP": "N/D"}
        texto_completo = ""
        try:
            with pdfplumber.open(caminho) as pdf:
                pagina = pdf.pages[0]
                topo_area = pagina.within_bbox((0, 0, pagina.width, 200))
                texto_topo = topo_area.extract_text() or ""
                texto_completo = pagina.extract_text() or ""
            
            if len(texto_completo.strip()) < 50:
                doc = fitz.open(caminho)
                pix = doc[0].get_pixmap()
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                texto_topo = pytesseract.image_to_string(cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY), lang='por')
                texto_completo = texto_topo

            nf_m = re.search(r'(?:N[º°]|Nota Fiscal N[º°]?)\s*(\d+)', texto_completo, re.I)
            if nf_m: d["NF"] = nf_m.group(1)
            
            cfop_pdf_m = re.search(r'(?:CFOP)\s*[:\.\-]?\s*(\d{4})', texto_completo, re.I)
            if cfop_pdf_m: d["CFOP"] = cfop_pdf_m.group(1)

            lixo = ["SÉRIE", "EMISSÃO", "VALOR", "TOTAL", "DANFE", "FOLHA", "CHAVE", "PROTOCOLO", "CNPJ", "INSCR", "FONE", "TEL", "UF", "DATA"]
            linhas = [l.strip() for l in texto_topo.split('\n') if l.strip()]
            for linha in linhas:
                if len(linha) > 5 and not any(x in linha.upper() for x in lixo):
                    if not re.match(r'^\d', linha):
                        d["Emitente"] = linha[:60].strip()
                        break

            vol_m = re.search(r'(?:QUANTIDADE|QTD|VOLUMES)\s*(\d+)', texto_completo, re.I)
            if vol_m: d["Vol"] = vol_m.group(1)
            
            mat_m = re.search(r'(?:DESCRIÇÃO DO PRODUTO)\s*\n.*?\d+\s+([^\n\r]+)', texto_completo, re.I)
            if mat_m: d["Material"] = mat_m.group(1).strip()[:40]

            return d
        except: return d

    def start(self):
        threading.Thread(target=self.run, daemon=True).start()

    def run(self):
        pythoncom.CoInitialize()
        self.btn.config(state="disabled")
        assistente_escolhido = self.combo_assistente.get()
        self.log(f"Iniciando rastreio v10.0 (SQLite Ativo). Filtro: {assistente_escolhido}")
        
        # --- INICIALIZAÇÃO DO BANCO DE DADOS SQLITE ---
        try:
            conn = sqlite3.connect(BANCO_SQLITE)
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS trocas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    assunto TEXT,
                    armazem TEXT,
                    data_hora_solicitacao TEXT,
                    nf_entrada TEXT,
                    emitente TEXT,
                    material TEXT,
                    cfop TEXT,
                    volume TEXT,
                    status TEXT,
                    data_hora_conclusao TEXT,
                    assistente TEXT,
                    nf_saida TEXT,
                    observacoes TEXT
                )
            ''')
            conn.commit()
            self.log("Conexão com o Banco de Dados estabelecida.")
        except Exception as e:
            self.log(f"Erro ao conectar no banco: {e}")
            messagebox.showerror("Erro de BD", f"Falha no Banco de Dados:\n{e}")
            self.btn.config(state="normal")
            return

        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.GetDefaultFolder(6)
            try: pasta = inbox.Folders("Troca de Notas")
            except: pasta = inbox.Parent.Folders("Troca de Notas")

            filtro = f"[ReceivedTime] >= '{self.ent_ini.get()} 00:00' AND [ReceivedTime] <= '{self.ent_fim.get()} 23:59'"
            mensagens = pasta.Items.Restrict(filtro)
            mensagens.Sort("[ReceivedTime]", False) 
            total = mensagens.Count
            
            fluxo = {}
            
            palavras_cancelamento = ["cancelamento", "cancelar", "desconsiderar", "cancela"]
            separadores_historico = ["\r\nde: ", "\nde: ", "\r\nfrom: ", "\nfrom: ", "________________________________", "-----mensagem original-----", "\nem ", "\r\nem "]

            for i, msg in enumerate(mensagens):
                self.progress['value'] = ((i + 1) / total) * 100
                cid = msg.ConversationID
                
                email = ""
                try: email = msg.SenderEmailAddress.lower() if msg.Sender.Type != "EX" else msg.Sender.GetExchangeUser().PrimarySmtpAddress.lower()
                except: pass

                corpo_cru = ""
                try: corpo_cru = msg.Body.lower()
                except: pass

                for sep in separadores_historico:
                    if sep in corpo_cru:
                        corpo_cru = corpo_cru.split(sep)[0]
                
                corpo_limpo = self.remover_acentos(corpo_cru)
                
                assunto_cru = ""
                try: assunto_cru = msg.Subject
                except: pass
                
                texto_busca = assunto_cru + " " + corpo_limpo
                placa_email = self.extrair_placa(texto_busca)

                if cid not in fluxo: 
                    fluxo[cid] = []

                ultima_transacao = fluxo[cid][-1] if fluxo[cid] else None
                eh_assistente = "mosaicco.com" in email

                nomes_anexos = []
                anexo_alvo = None
                if msg.Attachments.Count > 0:
                    for j in range(1, msg.Attachments.Count + 1):
                        a = msg.Attachments.Item(j)
                        nome_arq = a.FileName.lower()
                        nomes_anexos.append(nome_arq)
                        if nome_arq.endswith(".xml") and not anexo_alvo: 
                            anexo_alvo = a
                    if not anexo_alvo:
                        for j in range(1, msg.Attachments.Count + 1):
                            a = msg.Attachments.Item(j)
                            if a.FileName.lower().endswith(".pdf"): 
                                anexo_alvo = a; break

                if eh_assistente:
                    if anexo_alvo and ultima_transacao and ultima_transacao['concl'] is None and not ultima_transacao['cancelado']: 
                        path = os.path.join(PASTA_ANEXOS, anexo_alvo.FileName)
                        anexo_alvo.SaveAsFile(path)
                        dados = self.ler_xml(path) if path.lower().endswith(".xml") else self.ler_pdf_limpo(path)
                        ultima_transacao['concl'] = {
                            'DataHora': msg.SentOn.strftime("%d/%m/%Y %H:%M:%S"), 
                            'Assistente': msg.SenderName,
                            'NF': dados['NF'],
                            'Vol_Saida': dados['Vol'],
                            'CFOP_Saida': dados['CFOP']
                        }
                    elif anexo_alvo:
                        achou_link = False
                        for cid_buscado, lista_t in fluxo.items():
                            for t in lista_t:
                                if t['concl'] is None and not t['cancelado']:
                                    bateu_placa = (placa_email is not None and t['placa_solic'] is not None and placa_email == t['placa_solic'])
                                    bateu_anexo = False
                                    for arq in nomes_anexos:
                                        if arq in t['anexos_iniciais'] and len(arq) > 10 and not arq.startswith("image"):
                                            bateu_anexo = True
                                            break
                                    if bateu_placa or bateu_anexo:
                                        path = os.path.join(PASTA_ANEXOS, anexo_alvo.FileName)
                                        anexo_alvo.SaveAsFile(path)
                                        dados = self.ler_xml(path) if path.lower().endswith(".xml") else self.ler_pdf_limpo(path)
                                        t['concl'] = {
                                            'DataHora': msg.SentOn.strftime("%d/%m/%Y %H:%M:%S"),
                                            'Assistente': msg.SenderName,
                                            'NF': dados['NF'],
                                            'Vol_Saida': dados['Vol'],
                                            'CFOP_Saida': dados['CFOP']
                                        }
                                        motivo = "Placa Encontrada" if bateu_placa else "Anexo Específico Idêntico"
                                        t['obs'].add(f"Vinculado via Radar Global ({motivo})")
                                        achou_link = True
                                        break
                            if achou_link: break
                else:
                    eh_cancelamento = any(p in corpo_limpo for p in palavras_cancelamento)

                    if anexo_alvo:
                        path = os.path.join(PASTA_ANEXOS, anexo_alvo.FileName)
                        anexo_alvo.SaveAsFile(path)
                        dados = self.ler_xml(path) if path.lower().endswith(".xml") else self.ler_pdf_limpo(path)

                        nova_transacao = {
                            'solic': {
                                'DataHora': msg.SentOn.strftime("%d/%m/%Y %H:%M:%S"), 
                                'Armazem': msg.SenderName,
                                'NF': dados['NF'], 'Material': dados['Material'], 
                                'Vol': dados['Vol'], 'Emitente': dados['Emitente'], 'Assunto': msg.Subject,
                                'CFOP': dados['CFOP']
                            },
                            'concl': None,
                            'cancelado': False,
                            'obs': set(),
                            'anexos_iniciais': nomes_anexos,
                            'placa_solic': placa_email 
                        }
                        fluxo[cid].append(nova_transacao)

                    elif eh_cancelamento and ultima_transacao and ultima_transacao['concl'] is None:
                        ultima_transacao['cancelado'] = True

            # --- GRAVAÇÃO NO SQLITE (UPSERT NATIVO) ---
            novos_registros = 0
            atualizados = 0

            for cid, lista_transacoes in fluxo.items():
                for f in lista_transacoes:
                    s = f['solic']
                    c = f['concl']
                    nome_conclusao = c['Assistente'] if c else "-"
                    
                    if assistente_escolhido != "Todos os Assistentes":
                        primeiro_nome = self.remover_acentos(assistente_escolhido.split()[0])
                        nome_conclusao_limpo = self.remover_acentos(nome_conclusao)
                        if not c or primeiro_nome not in nome_conclusao_limpo: continue

                    status_final = "PENDENTE"
                    if f['cancelado']: status_final = "CANCELADO"
                    elif c: status_final = "CONCLUÍDO"

                    observacoes = " | ".join(f['obs']) if f['obs'] else "-"
                    if f['cancelado']: observacoes = "Cancelamento identificado no e-mail isolado."

                    vol_saida = c['Vol_Saida'] if c else "0"
                    vol_solic = s['Vol'] if s else "0"
                    volume_final = vol_saida if vol_saida != "0" else vol_solic
                    
                    cfop_saida = c['CFOP_Saida'] if c else "N/D"
                    cfop_solic = s['CFOP'] if s else "N/D"
                    cfop_final = cfop_saida if cfop_saida != "N/D" else cfop_solic

                    # Variáveis para o banco
                    assunto = s['Assunto']
                    armazem = s['Armazem']
                    dh_solic = s['DataHora']
                    nf_ent = s['NF']
                    emit = s['Emitente']
                    mat = s['Material']
                    dh_concl = c['DataHora'] if c else "-"
                    nf_sai = c['NF'] if c else "-"

                    # Checa se a nota já existe cruzando Data + Assunto + NF (Impede duplicar 'N/D')
                    cursor.execute('''
                        SELECT id, status FROM trocas
                        WHERE data_hora_solicitacao = ? AND assunto = ? AND nf_entrada = ?
                    ''', (dh_solic, assunto, nf_ent))
                    linha_banco = cursor.fetchone()

                    if linha_banco:
                        row_id, row_status = linha_banco
                        # Se estava pendente no banco e agora foi concluída, faz o UPDATE
                        if row_status == 'PENDENTE' and status_final == 'CONCLUÍDO':
                            cursor.execute('''
                                UPDATE trocas
                                SET status = ?, data_hora_conclusao = ?, assistente = ?, nf_saida = ?, cfop = ?, volume = ?, observacoes = ?
                                WHERE id = ?
                            ''', (status_final, dh_concl, nome_conclusao, nf_sai, cfop_final, volume_final, observacoes, row_id))
                            atualizados += 1
                    else:
                        # Se não existir, faz o INSERT (Inserção nova)
                        cursor.execute('''
                            INSERT INTO trocas (
                                assunto, armazem, data_hora_solicitacao, nf_entrada, emitente,
                                material, cfop, volume, status, data_hora_conclusao, assistente,
                                nf_saida, observacoes
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (assunto, armazem, dh_solic, nf_ent, emit, mat, cfop_final, volume_final, status_final, dh_concl, nome_conclusao, nf_sai, observacoes))
                        novos_registros += 1

            conn.commit()

            # --- GERA O ESPELHO EM EXCEL PARA VISUALIZAÇÃO ---
            df_export = pd.read_sql_query("SELECT * FROM trocas", conn)
            
            if not df_export.empty:
                # Remove a coluna de ID do banco para o Excel ficar limpo
                df_export.drop(columns=['id'], inplace=True, errors='ignore')
                
                # Renomeia as colunas para o formato antigo amigável
                df_export.columns = [
                    "Assunto", "Armazém", "Data/Hora Solicitação", "NF Entrada",
                    "Emitente (Topo da Nota)", "Material", "CFOP", "Volume (Peso Líquido)",
                    "Status", "Data/Hora Conclusão", "Assistente", "NF Saída (Nova)", "Observações"
                ]
                
                # Tratamento final do caractere '&'
                if 'Emitente (Topo da Nota)' in df_export.columns:
                    df_export['Emitente (Topo da Nota)'] = df_export['Emitente (Topo da Nota)'].astype(str).str.replace('&amp;', '&', regex=False)
                
                df_export.to_excel(ARQUIVO_FINAL, index=False)
                messagebox.showinfo("Sucesso SQLite", f"Banco de Dados atualizado!\nNovos registros: {novos_registros}\nPendentes concluídos: {atualizados}\n\nEspelho Excel atualizado com sucesso.")
                os.startfile(os.path.dirname(ARQUIVO_FINAL))
            else:
                messagebox.showwarning("Aviso", "O Banco de Dados está vazio. Nenhum processo encontrado.")

        except Exception as e: messagebox.showerror("Erro Crítico", str(e))
        finally:
            if 'conn' in locals(): conn.close()
            pythoncom.CoUninitialize()
            self.btn.config(state="normal")

if __name__ == "__main__":
    root = tk.Tk(); AppMosaicMaster(root); root.mainloop()