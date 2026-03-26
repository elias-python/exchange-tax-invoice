import win32com.client
import pythoncom
import os
import xml.etree.ElementTree as ET
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

# --- CONFIGURAÇÃO OCR ---
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\esantan3\OneDrive - The Mosaic Company\Área de Trabalho\Projetos\tesseract-main\tesseract.exe'

# --- CAMINHOS ---
PASTA_ANEXOS = r"C:\Users\esantan3\OneDrive - The Mosaic Company\Área de Trabalho\Projetos\Exchange Tax Invoice\Folder XML_PDF"
ARQUIVO_FINAL = r"C:\Users\esantan3\OneDrive - The Mosaic Company\Área de Trabalho\Projetos\Exchange Tax Invoice\Final Report\Relatorio_Controle_Trocas.xlsx"

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
        self.root.title("Mosaic Tool v7.8 - Radar Duplo (Texto e Anexo)")
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

    def limpar_assunto_radar(self, texto):
        # Essa função arranca "RES:", "ENC:" e espaços extras para podermos comparar os textos cruamente
        if not texto: return ""
        t = texto.lower()
        t = re.sub(r'^(re|res|enc|fwd|fw)[\s:]+', '', t).strip()
        return self.remover_acentos(t)

    def ler_xml(self, caminho):
        ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
        d = {"NF": "N/D", "Material": "N/D", "Vol": "0", "Emitente": "N/D"}
        try:
            tree = ET.parse(caminho); root = tree.getroot()
            inf = root.find('.//nfe:infNFe', ns) if root.tag.endswith('nfeProc') else root
            if inf is not None:
                d["NF"] = inf.find('.//nfe:ide/nfe:nNF', ns).text
                d["Material"] = inf.find('.//nfe:det/nfe:prod/nfe:xProd', ns).text
                vol = inf.find('.//nfe:transp/nfe:vol/nfe:qVol', ns)
                d["Vol"] = vol.text if vol is not None else "0"
                emit = inf.find('.//nfe:emit/nfe:xNome', ns)
                d["Emitente"] = emit.text if emit is not None else "N/D"
            return d
        except: return d

    def ler_pdf_limpo(self, caminho):
        d = {"NF": "N/D", "Material": "N/D", "Vol": "0", "Emitente": "N/D"}
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
        self.log(f"Iniciando rastreio v7.8. Filtro: {assistente_escolhido}")
        
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
                
                # Prepara o Assunto Limpo para o nosso novo Radar
                assunto_cru = ""
                try: assunto_cru = msg.Subject
                except: pass
                
                assunto_limpo_radar = self.limpar_assunto_radar(assunto_cru)
                texto_busca = assunto_limpo_radar + " " + corpo_limpo

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
                    # 1. TENTA AMARRAR PELA CONVERSA
                    if anexo_alvo and ultima_transacao and ultima_transacao['concl'] is None and not ultima_transacao['cancelado']: 
                        path = os.path.join(PASTA_ANEXOS, anexo_alvo.FileName)
                        anexo_alvo.SaveAsFile(path)
                        dados = self.ler_xml(path) if path.lower().endswith(".xml") else self.ler_pdf_limpo(path)
                        ultima_transacao['concl'] = {
                            'DataHora': msg.ReceivedTime.strftime("%d/%m/%Y %H:%M:%S"),
                            'Assistente': msg.SenderName,
                            'NF': dados['NF']
                        }
                    # 2. RADAR DUPLO: Procura por Texto (Assunto) OU por Anexo!
                    elif anexo_alvo:
                        achou_link = False
                        for cid_buscado, lista_t in fluxo.items():
                            for t in lista_t:
                                if t['concl'] is None and not t['cancelado']:
                                    
                                    # Condição A: O Assunto é igual? (ignorando RES: e ENC:)
                                    bateu_assunto = (assunto_limpo_radar == t['assunto_radar'] and len(assunto_limpo_radar) > 5)
                                    
                                    # Condição B: O nome do anexo é igual?
                                    bateu_anexo = any(arq in t['anexos_iniciais'] for arq in nomes_anexos)

                                    if bateu_assunto or bateu_anexo:
                                        path = os.path.join(PASTA_ANEXOS, anexo_alvo.FileName)
                                        anexo_alvo.SaveAsFile(path)
                                        dados = self.ler_xml(path) if path.lower().endswith(".xml") else self.ler_pdf_limpo(path)
                                        t['concl'] = {
                                            'DataHora': msg.ReceivedTime.strftime("%d/%m/%Y %H:%M:%S"),
                                            'Assistente': msg.SenderName,
                                            'NF': dados['NF']
                                        }
                                        motivo = "Texto do Assunto" if bateu_assunto else "Nome do Anexo"
                                        t['obs'].add(f"Vinculado via Radar Global ({motivo})")
                                        achou_link = True
                                        break
                            if achou_link: break
                else:
                    eh_cancelamento = any(p in texto_busca for p in palavras_cancelamento)

                    if anexo_alvo:
                        path = os.path.join(PASTA_ANEXOS, anexo_alvo.FileName)
                        anexo_alvo.SaveAsFile(path)
                        dados = self.ler_xml(path) if path.lower().endswith(".xml") else self.ler_pdf_limpo(path)

                        nova_transacao = {
                            'solic': {
                                'DataHora': msg.ReceivedTime.strftime("%d/%m/%Y %H:%M:%S"),
                                'Armazem': msg.SenderName,
                                'NF': dados['NF'], 'Material': dados['Material'], 
                                'Vol': dados['Vol'], 'Emitente': dados['Emitente'], 'Assunto': msg.Subject
                            },
                            'concl': None,
                            'cancelado': False,
                            'obs': set(),
                            'anexos_iniciais': nomes_anexos,
                            'assunto_radar': assunto_limpo_radar # Salva o Assunto Limpo pra ser procurado depois!
                        }
                        fluxo[cid].append(nova_transacao)

                    elif eh_cancelamento and ultima_transacao and ultima_transacao['concl'] is None:
                        ultima_transacao['cancelado'] = True

            final = []
            for cid, lista_transacoes in fluxo.items():
                for f in lista_transacoes:
                    s = f['solic']
                    c = f['concl']
                    nome_conclusao = c['Assistente'] if c else "-"
                    
                    if assistente_escolhido != "Todos os Assistentes":
                        primeiro_nome = self.remover_acentos(assistente_escolhido.split()[0])
                        nome_conclusao_limpo = self.remover_acentos(nome_conclusao)
                        
                        if not c or primeiro_nome not in nome_conclusao_limpo:
                            continue

                    status_final = "PENDENTE"
                    if f['cancelado']: status_final = "CANCELADO"
                    elif c: status_final = "CONCLUÍDO"

                    observacoes = " | ".join(f['obs']) if f['obs'] else "-"
                    if f['cancelado']: observacoes = "Cancelamento identificado no e-mail isolado."

                    final.append({
                        "Assunto": s['Assunto'],
                        "Armazém": s['Armazem'],
                        "Data/Hora Solicitação": s['DataHora'],
                        "NF Entrada": s['NF'],
                        "Emitente (Topo da Nota)": s['Emitente'],
                        "Material": s['Material'],
                        "Volume": s['Vol'],
                        "Status": status_final,
                        "Data/Hora Conclusão": c['DataHora'] if c else "-",
                        "Assistente": nome_conclusao,
                        "NF Saída (Nova)": c['NF'] if c else "-",
                        "Observações": observacoes
                    })

            if final:
                df = pd.DataFrame(final)
                if 'Emitente (Topo da Nota)' in df.columns:
                    df['Emitente (Topo da Nota)'] = df['Emitente (Topo da Nota)'].str.replace('&amp;', '&', regex=False)
                df.to_excel(ARQUIVO_FINAL, index=False)
                messagebox.showinfo("Sucesso", f"Relatório gerado!\n{len(final)} processos (Radar de Texto/Anexo ativado).")
                os.startfile(os.path.dirname(ARQUIVO_FINAL))
            else:
                messagebox.showwarning("Aviso", "Nenhum processo encontrado.")

        except Exception as e: messagebox.showerror("Erro", str(e))
        finally:
            pythoncom.CoUninitialize(); self.btn.config(state="normal")

if __name__ == "__main__":
    root = tk.Tk(); AppMosaicMaster(root); root.mainloop()