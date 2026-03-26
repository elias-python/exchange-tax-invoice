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

# --- CONFIGURAÇÃO OCR ---
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\esantan3\OneDrive - The Mosaic Company\Área de Trabalho\Projetos\tesseract-main\tesseract.exe'

# --- CAMINHOS ---
PASTA_ANEXOS = r"C:\Users\esantan3\OneDrive - The Mosaic Company\Área de Trabalho\Projetos\Exchange Tax Invoice\Folder XML_PDF"
ARQUIVO_FINAL = r"C:\Users\esantan3\OneDrive - The Mosaic Company\Área de Trabalho\Projetos\Exchange Tax Invoice\Final Report\Relatorio_Controle_Trocas.xlsx"

class AppMosaicMaster:
    def __init__(self, root):
        self.root = root
        self.root.title("Mosaic Tool v5.0 - Paranaguá")
        self.root.geometry("700x550")
        
        tk.Label(root, text="Controle de Fluxo: Emitente Real vs Conclusão", font=("Arial", 11, "bold")).pack(pady=10)
        
        frame = tk.Frame(root)
        frame.pack(pady=5)
        self.ent_ini = tk.Entry(frame); self.ent_ini.insert(0, "20/03/2026"); self.ent_ini.grid(row=0, column=0, padx=5)
        self.ent_fim = tk.Entry(frame); self.ent_fim.insert(0, "25/03/2026"); self.ent_fim.grid(row=0, column=1, padx=5)

        self.btn = tk.Button(root, text="SINCRONIZAR E FILTRAR EMITENTES", command=self.start, bg="#004a8d", fg="white", font=("Arial", 10, "bold"))
        self.btn.pack(pady=10)

        self.progress = ttk.Progressbar(root, length=600, mode="determinate")
        self.progress.pack(pady=10)

        self.log_txt = tk.Text(root, height=15, width=90, font=("Consolas", 8))
        self.log_txt.pack(pady=10, padx=10)

    def log(self, msg):
        self.log_txt.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_txt.see(tk.END)

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
                # Pegamos apenas a parte superior da primeira página (primeiros 200 pixels de altura)
                pagina = pdf.pages[0]
                topo_area = pagina.within_bbox((0, 0, pagina.width, 200))
                texto_topo = topo_area.extract_text() or ""
                texto_completo = pagina.extract_text() or ""
            
            if len(texto_completo.strip()) < 50: # OCR se necessário
                doc = fitz.open(caminho)
                pix = doc[0].get_pixmap()
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                texto_topo = pytesseract.image_to_string(cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY), lang='por')
                texto_completo = texto_topo

            # 1. Busca NF
            nf_m = re.search(r'(?:N[º°]|Nota Fiscal N[º°]?)\s*(\d+)', texto_completo, re.I)
            if nf_m: d["NF"] = nf_m.group(1)

            # 2. Busca Emitente (Limpeza de Ruído do Topo)
            # Lista de palavras que NUNCA são o nome da empresa
            lixo = ["SÉRIE", "EMISSÃO", "VALOR", "TOTAL", "DANFE", "FOLHA", "CHAVE", "PROTOCOLO", "CNPJ", "INSCR", "FONE", "TEL", "UF", "DATA"]
            
            linhas = [l.strip() for l in texto_topo.split('\n') if l.strip()]
            for linha in linhas:
                # Regras para validar se é o nome da empresa:
                # - Ter mais de 5 caracteres
                # - Não conter palavras da lista 'lixo'
                # - Não começar com números (evita datas e séries)
                if len(linha) > 5 and not any(x in linha.upper() for x in lixo):
                    if not re.match(r'^\d', linha): # Se não começar com dígito
                        d["Emitente"] = linha[:60].strip()
                        break

            # 3. Volume e Material (Busca simples no texto completo)
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
        self.log("Limpando ruídos e sincronizando notas...")
        
        try:
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.GetDefaultFolder(6)
            try: pasta = inbox.Folders("Troca de Notas")
            except: pasta = inbox.Parent.Folders("Troca de Notas")

            filtro = f"[ReceivedTime] >= '{self.ent_ini.get()} 00:00' AND [ReceivedTime] <= '{self.ent_fim.get()} 23:59'"
            mensagens = pasta.Items.Restrict(filtro)
            total = mensagens.Count
            
            fluxo = {}

            for i, msg in enumerate(mensagens):
                self.progress['value'] = ((i + 1) / total) * 100
                cid = msg.ConversationID
                
                email = ""
                try: email = msg.SenderEmailAddress.lower() if msg.Sender.Type != "EX" else msg.Sender.GetExchangeUser().PrimarySmtpAddress.lower()
                except: pass

                if cid not in fluxo: fluxo[cid] = {'solic': None, 'concl': None}

                if msg.Attachments.Count > 0:
                    anexo_alvo = None
                    for j in range(1, msg.Attachments.Count + 1):
                        a = msg.Attachments.Item(j)
                        if a.FileName.lower().endswith(".xml"): anexo_alvo = a; break
                    
                    if not anexo_alvo:
                        for j in range(1, msg.Attachments.Count + 1):
                            a = msg.Attachments.Item(j)
                            if a.FileName.lower().endswith(".pdf"): anexo_alvo = a; break

                    if anexo_alvo:
                        path = os.path.join(PASTA_ANEXOS, anexo_alvo.FileName)
                        anexo_alvo.SaveAsFile(path)
                        
                        # Usa a nova função ler_pdf_limpo
                        dados = self.ler_xml(path) if path.lower().endswith(".xml") else self.ler_pdf_limpo(path)

                        if "mosaicco.com" in email:
                            fluxo[cid]['concl'] = {
                                'DataHora': msg.ReceivedTime.strftime("%d/%m/%Y %H:%M:%S"),
                                'Assistente': msg.SenderName,
                                'NF': dados['NF']
                            }
                        else:
                            fluxo[cid]['solic'] = {
                                'DataHora': msg.ReceivedTime.strftime("%d/%m/%Y %H:%M:%S"),
                                'Armazem': msg.SenderName,
                                'NF': dados['NF'], 'Material': dados['Material'], 
                                'Vol': dados['Vol'], 'Emitente': dados['Emitente'], 'Assunto': msg.Subject
                            }

            final = []
            for cid, f in fluxo.items():
                if f['solic']:
                    s = f['solic']; c = f['concl']
                    final.append({
                        "Assunto": s['Assunto'],
                        "Armazém": s['Armazem'],
                        "Data/Hora Solicitação": s['DataHora'],
                        "NF Entrada": s['NF'],
                        "Emitente (Topo da Nota)": s['Emitente'],
                        "Material": s['Material'],
                        "Volume": s['Vol'],
                        "Status": "CONCLUÍDO" if c else "PENDENTE",
                        "Data/Hora Conclusão": c['DataHora'] if c else "-",
                        "Assistente": c['Assistente'] if c else "-",
                        "NF Saída (Nova)": c['NF'] if c else "-"
                    })

            if final:
                # Remove duplicatas de texto como 'P&amp;K' vindas do XML
                df = pd.DataFrame(final)
                if 'Emitente (Topo da Nota)' in df.columns:
                    df['Emitente (Topo da Nota)'] = df['Emitente (Topo da Nota)'].str.replace('&amp;', '&', regex=False)
                
                df.to_excel(ARQUIVO_FINAL, index=False)
                messagebox.showinfo("Sucesso", "Relatório atualizado e limpo!")
                os.startfile(os.path.dirname(ARQUIVO_FINAL))

        except Exception as e: messagebox.showerror("Erro", str(e))
        finally:
            pythoncom.CoUninitialize(); self.btn.config(state="normal")

if __name__ == "__main__":
    root = tk.Tk(); AppMosaicMaster(root); root.mainloop()