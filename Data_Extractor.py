import win32com.client
import pythoncom
import os
import pandas as pd
from datetime import datetime, timedelta
import re
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import unicodedata
import sqlite3
import shutil

# --- CAMINHOS ---
PASTA_ANEXOS = r"C:\Users\esantan3\OneDrive - The Mosaic Company\Área de Trabalho\Projetos\Exchange Tax Invoice\Folder XML_PDF"
ARQUIVO_FINAL = r"C:\Users\esantan3\OneDrive - The Mosaic Company\Área de Trabalho\Projetos\Exchange Tax Invoice\Final Report\Relatorio_Controle_Trocas.xlsx"
BANCO_SQLITE = r"C:\Users\esantan3\OneDrive - The Mosaic Company\Área de Trabalho\Projetos\Exchange Tax Invoice\Final Report\Banco_Controle_Trocas.db"
BACKUP_BANCO_ATIVO = True
MODO_CONSERVADOR_POWER_BI = True
MODO_ESTRITO_HISTORICO = True
JANELA_MINIMA_RECONCILIACAO_DIAS = 30
HORAS_ALERTA_PENDENTE = 24

# --- LISTA DE ASSISTENTES ---
LISTA_ASSISTENTES = [
    "Todos os Assistentes",
    "Fernando Rodrigues",
    "Gustavo Chaves",
    "João Teixeira",
    "José Viana",
    "Vitória Nunes",
    "João Costa",
]


class AppMosaicMaster:
    def to_naive(self, dt):
        if dt is not None and hasattr(dt, "tzinfo") and dt.tzinfo is not None:
            return dt.replace(tzinfo=None)
        return dt

    def __init__(self, root):
        self.root = root
        self.root.title("Monitor de Trocas - The Mosaic Company")
        self.root.geometry("800x670")
        self.root.configure(bg="#F4F6F9")

        self.is_running = False
        self.auto_mode = tk.BooleanVar(value=True)
        self.intervalo_min = tk.IntVar(value=15)
        self.tempo_restante = self.intervalo_min.get() * 60

        # --- ESTILIZAÇÃO CORPORATIVA ---
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TProgressbar", thickness=15, background="#004B87")
        style.configure(
            "Title.TLabel",
            font=("Segoe UI", 18, "bold"),
            foreground="#FFFFFF",
            background="#004B87",
        )
        style.configure(
            "Subtitle.TLabel",
            font=("Segoe UI", 10),
            foreground="#CCE0FF",
            background="#004B87",
        )

        style.configure("Card.TFrame", background="#FFFFFF", relief="flat")
        style.configure(
            "Clock.TLabel",
            font=("Segoe UI Light", 36),
            foreground="#004B87",
            background="#FFFFFF",
        )
        style.configure(
            "Status.TLabel",
            font=("Segoe UI", 11),
            foreground="#333333",
            background="#FFFFFF",
        )

        style.configure(
            "TButton",
            font=("Segoe UI", 10, "bold"),
            background="#004B87",
            foreground="white",
            borderwidth=0,
        )
        style.map("TButton", background=[["active", "#003366"]])

        # --- CABEÇALHO ---
        frame_top = tk.Frame(root, bg="#004B87", height=80)
        frame_top.pack(fill="x")
        ttk.Label(frame_top, text="THE MOSAIC COMPANY", style="Title.TLabel").pack(
            pady=(15, 2)
        )
        ttk.Label(
            frame_top,
            text="Monitoramento Automatizado de Troca de Notas Fiscais",
            style="Subtitle.TLabel",
        ).pack(pady=(0, 15))

        # --- PAINEL CENTRAL (CARD) ---
        frame_center = ttk.Frame(root, style="Card.TFrame", padding=30)
        frame_center.pack(pady=20, padx=40, fill="both", expand=True)

        self.lbl_clock = ttk.Label(frame_center, text="15:00", style="Clock.TLabel")
        self.lbl_clock.pack(pady=(5, 2))

        # --- SELEÇÃO DE DIAS ---
        frame_dias = tk.Frame(frame_center, bg="#FFFFFF")
        frame_dias.pack(pady=(2, 10))

        lbl_dias = tk.Label(
            frame_dias,
            text="Analisar e-mails dos últimos:",
            font=("Segoe UI", 9, "bold"),
            fg="#555555",
            bg="#FFFFFF",
        )
        lbl_dias.pack(side="left", padx=5)

        self.janela_dias = tk.IntVar(value=7)
        self.cb_dias = ttk.Combobox(
            frame_dias,
            textvariable=self.janela_dias,
            values=[1, 3, 7, 15, 30, 60, 90],
            width=5,
            state="readonly",
        )
        self.cb_dias.pack(side="left", padx=5)

        lbl_dias_sufixo = tk.Label(
            frame_dias,
            text="dias corridos.",
            font=("Segoe UI", 9),
            fg="#555555",
            bg="#FFFFFF",
        )
        lbl_dias_sufixo.pack(side="left")

        self.lbl_etapa = ttk.Label(
            frame_center,
            text="Sistema em *standby* (aguardando execução).",
            style="Status.TLabel",
        )
        self.lbl_etapa.pack(pady=(5, 15))

        self.progress = ttk.Progressbar(frame_center, length=600, mode="determinate")
        self.progress.pack(pady=10)

        # --- BOTÃO DE AÇÃO MANUAL ---
        self.btn = ttk.Button(
            frame_center, text="SINCRONIZAR AGORA", command=self.start_manual, width=25
        )
        self.btn.pack(pady=15)

        # --- TERMINAL DE LOG ---
        frame_terminal = tk.Frame(root, bg="#FFFFFF", bd=1, relief="solid")
        frame_terminal.pack(padx=40, pady=(0, 10), fill="x")

        self.log_txt = tk.Text(
            frame_terminal,
            height=5,
            font=("Consolas", 9),
            bg="#FAFAFA",
            fg="#555555",
            bd=0,
            padx=10,
            pady=10,
        )
        self.log_txt.pack(fill="x")

        # --- RODAPÉ ---
        self.lbl_rodape = tk.Label(
            root,
            text="Automação D&A | Janela de processamento contínuo ajustável.",
            font=("Segoe UI", 8),
            fg="#888888",
            bg="#F4F6F9",
        )
        self.lbl_rodape.pack(side="bottom", pady=5)

        self.log("Sistema Iniciado. Monitoramento em background ativado.")
        self.root.after(1000, self.tick_relogio)

    # ================= LOG E FERRAMENTAS =================
    def log(self, msg):
        self.log_txt.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_txt.see(tk.END)
        self.root.update_idletasks()

    def extrair_placa(self, texto):
        if not texto:
            return None
        busca = re.search(r"[A-Za-z]{3}[-\s]?[0-9][A-Za-z0-9][0-9]{2}", texto)
        return (
            busca.group(0).replace("-", "").replace(" ", "").upper() if busca else None
        )

    def extrair_nf_texto(self, texto):
        if not texto:
            return None
        busca = re.search(
            r"(?:n\s*f\s*e?|n\s*[ºo°]?\s*[:\-]?)\s*(\d{4,10})",
            texto,
            re.IGNORECASE,
        )
        return busca.group(1) if busca else None

    def parse_data_hora(self, texto):
        t = (texto or "").strip()
        if not t or t == "-":
            return None
        for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M"):
            try:
                return datetime.strptime(t, fmt)
            except ValueError:
                pass
        return None

    def registrar_evento_diagnostico(
        self,
        cursor,
        chave,
        tipo,
        assunto,
        remetente,
        conversation_id,
        nf_ref,
        placa_ref,
        detalhe,
    ):
        cursor.execute(
            """
            INSERT OR IGNORE INTO diagnostico_emails
            (chave_evento, data_evento, tipo_evento, assunto, remetente, conversation_id, nf_ref, placa_ref, detalhe)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                chave,
                datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                tipo,
                assunto or "-",
                remetente or "-",
                conversation_id or "-",
                nf_ref or "-",
                placa_ref or "-",
                detalhe or "-",
            ),
        )

    def buscar_pendente_no_banco(self, cursor, nf_resp, placa_resp, sent_on):
        candidatos = []

        if nf_resp:
            cursor.execute(
                """
                SELECT id, assunto, data_hora_solicitacao, nf_entrada
                FROM trocas
                WHERE status = 'PENDENTE' AND nf_entrada = ?
                ORDER BY id DESC
                """,
                (nf_resp,),
            )
            candidatos.extend(cursor.fetchall())

        if placa_resp:
            cursor.execute("""
                SELECT id, assunto, data_hora_solicitacao, nf_entrada
                FROM trocas
                WHERE status = 'PENDENTE' AND assunto IS NOT NULL
                ORDER BY id DESC
                """)
            for row in cursor.fetchall():
                placa_assunto = self.extrair_placa(row[1] or "")
                if placa_assunto and placa_assunto == placa_resp:
                    candidatos.append(row)

        melhor = None
        melhor_dt = None
        vistos = set()
        for row in candidatos:
            if row[0] in vistos:
                continue
            vistos.add(row[0])
            dt_solic = self.parse_data_hora(row[2])
            dt_solic_naive = self.to_naive(dt_solic)
            sent_on_naive = self.to_naive(sent_on)
            if not dt_solic_naive:
                continue
            if dt_solic_naive <= sent_on_naive and (
                melhor_dt is None or dt_solic_naive > melhor_dt
            ):
                melhor = row
                melhor_dt = dt_solic_naive

        return melhor

    def remover_acentos(self, texto):
        if not texto:
            return ""
        return (
            unicodedata.normalize("NFKD", texto)
            .encode("ASCII", "ignore")
            .decode("utf-8")
            .lower()
        )

    def nome_eh_assistente(self, nome_remetente):
        nome_norm = re.sub(r"[^a-z0-9\s]", " ", self.remover_acentos(nome_remetente))
        nome_norm = re.sub(r"\s+", " ", nome_norm).strip()
        for ast in LISTA_ASSISTENTES[1:]:
            ast_norm = re.sub(r"[^a-z0-9\s]", " ", self.remover_acentos(ast))
            ast_tokens = [t for t in ast_norm.split() if len(t) > 2]
            if ast_tokens and all(tok in nome_norm for tok in ast_tokens):
                return True
        return False

    def fazer_backup_banco(self):
        if not BACKUP_BANCO_ATIVO:
            return
        if not os.path.exists(BANCO_SQLITE):
            return
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = BANCO_SQLITE.replace(".db", f"_backup_{stamp}.db")
        shutil.copy2(BANCO_SQLITE, backup_path)
        self.log(f"Backup do banco criado: {os.path.basename(backup_path)}")

    # ================= LEITURA XML =================
    def ler_xml(self, caminho):
        d = {
            "nf": "N/D",
            "material": "N/D",
            "volume": "0",
            "qvol": "0",
            "emitente": "N/D",
            "cfop": "N/D",
            "cnpj_emitente": "N/D",
            "cnpj_destinatario": "N/D",
            "nome_destinatario": "N/D",
            "transportadora": "N/D",
        }
        try:
            with open(caminho, "r", encoding="utf-8", errors="ignore") as f:
                conteudo = f.read()
            conteudo_limpo = re.sub(r"(</?)[a-zA-Z0-9\-]+:", r"\1", conteudo)

            nf_m = re.search(r"<nNF>(\d+)</nNF>", conteudo_limpo)
            if nf_m:
                d["nf"] = nf_m.group(1)
            mat_m = re.search(r"<xProd>([^<]+)</xProd>", conteudo_limpo)
            if mat_m:
                d["material"] = (
                    mat_m.group(1).replace("&quot;", '"').replace("&amp;", "&").strip()
                )
            cfop_m = re.search(
                r"<CFOP>\s*([^<]+)\s*</CFOP>", conteudo_limpo, re.IGNORECASE
            )
            if cfop_m:
                d["cfop"] = cfop_m.group(1).strip()

            pesol_m = re.search(
                r"<pesoL>\s*([\d\.,]+)\s*</pesoL>", conteudo_limpo, re.IGNORECASE
            )
            pesol_val = 0
            if pesol_m:
                val_p = pesol_m.group(1).replace(",", ".")
                if val_p.count(".") > 1:
                    val_p = val_p.replace(".", "", val_p.count(".") - 1)
                pesol_val = int(float(val_p))
                d["volume"] = str(pesol_val)

            qvol_m = re.search(
                r"<qVol>\s*([\d\.,]+)\s*</qVol>", conteudo_limpo, re.IGNORECASE
            )
            qvol_val = 0
            if qvol_m:
                val_q = qvol_m.group(1).replace(",", ".")
                if val_q.count(".") > 1:
                    val_q = val_q.replace(".", "", val_q.count(".") - 1)
                qvol_val = int(float(val_q))
                d["qvol"] = str(qvol_val)

            emit_b = re.search(r"<emit>(.*?)</emit>", conteudo_limpo, re.S)
            if emit_b:
                d["cnpj_emitente"] = (
                    (re.search(r"<CNPJ>(\d+)", emit_b.group(1)) or re.search(r"", ""))
                    .group(0)
                    .replace("<CNPJ>", "")
                    .strip()
                )
                d["emitente"] = (
                    (
                        re.search(r"<xNome>([^<]+)", emit_b.group(1))
                        or re.search(r"", "")
                    )
                    .group(0)
                    .replace("<xNome>", "")
                    .strip()
                )

            dest_b = re.search(r"<dest>(.*?)</dest>", conteudo_limpo, re.S)
            if dest_b:
                d["nome_destinatario"] = (
                    (
                        re.search(r"<xNome>([^<]+)", dest_b.group(1))
                        or re.search(r"", "")
                    )
                    .group(0)
                    .replace("<xNome>", "")
                    .strip()
                )
                d["cnpj_destinatario"] = (
                    (
                        re.search(r"<(?:CNPJ|CPF)>(\d+)", dest_b.group(1))
                        or re.search(r"", "")
                    )
                    .group(0)
                    .replace("<CNPJ>", "")
                    .replace("<CPF>", "")
                    .strip()
                )

            transp_b = re.search(
                r"<transporta>(.*?)</transporta>", conteudo_limpo, re.S
            )
            if transp_b:
                d["transportadora"] = (
                    (
                        re.search(r"<xNome>([^<]+)", transp_b.group(1))
                        or re.search(r"", "")
                    )
                    .group(0)
                    .replace("<xNome>", "")
                    .strip()
                )

            esp_m = re.search(r"<esp>([^<]+)</esp>", conteudo_limpo, re.IGNORECASE)
            esp_text = esp_m.group(1).upper() if esp_m else ""

            if "EMBALAGEM" in esp_text:
                if qvol_val > 0:
                    d["volume"] = str(qvol_val)

            return d
        except Exception:
            return d

    # ================= MOTOR DE AUTOMAÇÃO =================
    def tick_relogio(self):
        if self.auto_mode.get() and not self.is_running:
            if self.tempo_restante <= 0:
                self.log("Iniciando rotina automática de extração.")
                self.start_processo()
            else:
                self.tempo_restante -= 1
                mins, secs = divmod(self.tempo_restante, 60)
                self.lbl_clock.config(
                    text=f"{mins:02d}:{secs:02d}", foreground="#004B87"
                )
        self.root.after(1000, self.tick_relogio)

    def start_manual(self):
        if self.is_running:
            return
        self.log("Execução manual solicitada pelo usuário.")
        self.start_processo()

    def start_processo(self):
        self.is_running = True
        self.btn.config(state="disabled")
        self.lbl_clock.config(text="PROCESSANDO", foreground="#0078D4")

        self.dias_para_busca = self.janela_dias.get()

        threading.Thread(target=self.run, daemon=True).start()

    # ================= PROCESSO PRINCIPAL =================
    def run(self):
        pythoncom.CoInitialize()
        self.lbl_etapa.config(
            text="Estabelecendo conexão com o Banco de Dados e Outlook..."
        )
        self.root.update_idletasks()

        # --- LISTA PARA ARMAZENAR OS ALERTAS ---
        alertas_justificativas = []

        try:
            self.fazer_backup_banco()
            conn = sqlite3.connect(BANCO_SQLITE)
            cursor = conn.cursor()

            cursor.execute(
                """CREATE TABLE IF NOT EXISTS trocas (
                id INTEGER PRIMARY KEY AUTOINCREMENT, assunto TEXT, armazem TEXT, data_hora_solicitacao TEXT, 
                nf_entrada TEXT, emitente TEXT, material TEXT, cfop TEXT, volume TEXT, qvol TEXT, 
                status TEXT, data_hora_conclusao TEXT, assistente TEXT, nf_saida TEXT, cnpj_emitente TEXT, 
                cnpj_destinatario TEXT, justificativa TEXT, observacoes TEXT, nome_destinatario TEXT,
                transportadora_saida TEXT, cfop_saida TEXT, padronizado_xml TEXT, entry_id TEXT)"""
            )
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS diagnostico_emails (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    chave_evento TEXT UNIQUE,
                    data_evento TEXT,
                    tipo_evento TEXT,
                    assunto TEXT,
                    remetente TEXT,
                    conversation_id TEXT,
                    nf_ref TEXT,
                    placa_ref TEXT,
                    detalhe TEXT
                )
                """)
            conn.commit()

            try:
                cursor.execute("ALTER TABLE trocas ADD COLUMN entry_id TEXT")
            except sqlite3.OperationalError:
                pass

            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace(
                "MAPI"
            )
            inbox = outlook.GetDefaultFolder(6)

            # Tenta encontrar a pasta com diferentes variações
            pasta = None
            nomes_pasta = ["Troca de Notas", "Troca De Notas", "trocas", "Trocas"]

            for nome in nomes_pasta:
                try:
                    pasta = inbox.Folders(nome)
                    self.log(f"Pasta encontrada: '{nome}'")
                    break
                except:
                    try:
                        pasta = inbox.Parent.Folders(nome)
                        self.log(f"Pasta encontrada (nível superior): '{nome}'")
                        break
                    except:
                        continue

            if pasta is None:
                # Se ainda não encontrou, listar todas as pastas disponíveis
                pastas_disponiveis = [f.Name for f in inbox.Folders]
                self.log(f"Pastas disponíveis: {', '.join(pastas_disponiveis)}")
                self.lbl_etapa.config(
                    text="ERRO: Pasta 'Troca de Notas' não encontrada.",
                    foreground="#D13438",
                )
                return

            hoje = datetime.now()
            dias_efetivos = max(self.dias_para_busca, JANELA_MINIMA_RECONCILIACAO_DIAS)
            sete_dias_atras = hoje - timedelta(days=dias_efetivos)
            data_fim_str = hoje.strftime("%d/%m/%Y")
            data_ini_str = sete_dias_atras.strftime("%d/%m/%Y")

            self.log(
                f"Processando registros no período de {data_ini_str} a {data_fim_str} ({dias_efetivos} dias efetivos para reconciliação)."
            )
            filtro = f"[ReceivedTime] >= '{data_ini_str} 00:00' AND [ReceivedTime] <= '{data_fim_str} 23:59'"

            mensagens = pasta.Items.Restrict(filtro)
            mensagens.Sort("[ReceivedTime]", False)

            total_msgs = mensagens.Count
            transacoes_ativas = []
            respostas_sem_match = []

            for i, msg in enumerate(mensagens):
                percent = int(((i + 1) / total_msgs) * 100) if total_msgs > 0 else 100
                self.progress["value"] = percent
                self.lbl_etapa.config(
                    text=f"Lendo base de e-mails ({i+1}/{total_msgs}) - {percent}%"
                )
                self.root.update_idletasks()

                cid = msg.ConversationID
                email = ""
                try:
                    email = (
                        msg.SenderEmailAddress.lower()
                        if msg.Sender.Type != "EX"
                        else msg.Sender.GetExchangeUser().PrimarySmtpAddress.lower()
                    )
                except:
                    pass

                nome_rem = self.remover_acentos(msg.SenderName)
                eh_assistente = "mosaicco.com" in email or any(
                    self.remover_acentos(ast) in nome_rem
                    for ast in LISTA_ASSISTENTES[1:]
                )
                if not eh_assistente:
                    eh_assistente = self.nome_eh_assistente(msg.SenderName)

                anexo_alvo = None
                tem_xml = False

                if msg.Attachments.Count > 0:
                    for j in range(1, msg.Attachments.Count + 1):
                        at = msg.Attachments.Item(j)
                        if at.FileName.lower().endswith(".xml"):
                            anexo_alvo = at
                            tem_xml = True
                            break
                    if not anexo_alvo:
                        for j in range(1, msg.Attachments.Count + 1):
                            at = msg.Attachments.Item(j)
                            if at.FileName.lower().endswith(".pdf"):
                                anexo_alvo = at
                                break

                if not eh_assistente:
                    # Processa emails de não-assistentes (armazéns/clientes) como solicitações
                    dados = {
                        "nf": "N/D",
                        "material": "N/D",
                        "volume": "0",
                        "qvol": "0",
                        "emitente": "N/D",
                        "cfop": "N/D",
                        "cnpj_emitente": "N/D",
                        "cnpj_destinatario": "N/D",
                        "nome_destinatario": "N/D",
                        "transportadora": "N/D",
                    }

                    if anexo_alvo:
                        nome_arq_unico = f"{msg.EntryID[:15]}_{anexo_alvo.FileName}"
                        path = os.path.join(PASTA_ANEXOS, nome_arq_unico)
                        anexo_alvo.SaveAsFile(path)
                        dados = self.ler_xml(path)

                    placa_solic = self.extrair_placa(
                        msg.Subject + " " + (msg.Body or "")
                    )
                    nf_texto = self.extrair_nf_texto(
                        msg.Subject + " " + (msg.Body or "")
                    )
                    nf_referencia = (
                        dados["nf"] if dados["nf"] != "N/D" else (nf_texto or "N/D")
                    )

                    duplicada = None
                    for t in reversed(transacoes_ativas):
                        mesmo_cid = t["cid"] == cid
                        mesmo_nf = (
                            nf_referencia != "N/D"
                            and t.get("nf_ref")
                            and t.get("nf_ref") == nf_referencia
                        )
                        janela_curta = (
                            abs((msg.SentOn - t["data_solic_dt"]).total_seconds())
                            <= 300
                        )
                        if (
                            t["concl"] is None
                            and janela_curta
                            and (mesmo_cid or mesmo_nf)
                        ):
                            duplicada = t
                            break

                    if duplicada:
                        # Mantém um único registro para envios duplicados (ex.: Nota + DANFE)
                        if (not duplicada["solic"]["tem_xml"]) and tem_xml:
                            duplicada["solic"].update(
                                {
                                    "nf": dados["nf"],
                                    "mat": dados["material"],
                                    "vol": dados["volume"],
                                    "qvol": dados["qvol"],
                                    "emit": dados["emitente"],
                                    "cfop": dados["cfop"],
                                    "cnpj_e": dados["cnpj_emitente"],
                                    "cnpj_d": dados["cnpj_destinatario"],
                                    "nome_d": dados["nome_destinatario"],
                                    "tem_xml": tem_xml,
                                }
                            )
                        if not duplicada.get("placa") and placa_solic:
                            duplicada["placa"] = placa_solic
                        continue

                    # Registra mesmo sem anexo, para capturar todas as solicitações
                    transacoes_ativas.append(
                        {
                            "cid": cid,
                            "placa": placa_solic,
                            "nf_ref": nf_referencia,
                            "data_solic_dt": msg.SentOn,
                            "solic": {
                                "entry_id": msg.EntryID,
                                "assunto": msg.Subject,
                                "armazem": msg.SenderName,
                                "data": msg.SentOn.strftime("%d/%m/%Y %H:%M:%S"),
                                "nf": dados["nf"],
                                "mat": dados["material"],
                                "vol": dados["volume"],
                                "qvol": dados["qvol"],
                                "emit": dados["emitente"],
                                "cfop": dados["cfop"],
                                "cnpj_e": dados["cnpj_emitente"],
                                "cnpj_d": dados["cnpj_destinatario"],
                                "nome_d": dados["nome_destinatario"],
                                "tem_xml": tem_xml,
                            },
                            "concl": None,
                        }
                    )
                else:
                    target = None
                    placa_resp = self.extrair_placa(
                        msg.Subject + " " + (msg.Body or "")
                    )
                    nf_resp = self.extrair_nf_texto(
                        msg.Subject + " " + (msg.Body or "")
                    )

                    for t in reversed(transacoes_ativas):
                        t_data_naive = self.to_naive(t["data_solic_dt"])
                        msg_sent_naive = self.to_naive(msg.SentOn)
                        if (
                            t["concl"] is None
                            and t["cid"] == cid
                            and t_data_naive <= msg_sent_naive
                        ):
                            target = t
                            break

                    if not target and placa_resp:
                        for t in reversed(transacoes_ativas):
                            t_data_naive = self.to_naive(t["data_solic_dt"])
                            msg_sent_naive = self.to_naive(msg.SentOn)
                            if (
                                t["concl"] is None
                                and t["placa"]
                                and t["placa"] == placa_resp
                                and t_data_naive <= msg_sent_naive
                            ):
                                target = t
                                self.log(
                                    f"Correspondência por placa encontrada: {placa_resp}"
                                )
                                break

                    if not target and nf_resp:
                        for t in reversed(transacoes_ativas):
                            t_data_naive = self.to_naive(t["data_solic_dt"])
                            msg_sent_naive = self.to_naive(msg.SentOn)
                            if (
                                t["concl"] is None
                                and t.get("nf_ref") == nf_resp
                                and t_data_naive <= msg_sent_naive
                            ):
                                target = t
                                self.log(
                                    f"Correspondência por NF encontrada: {nf_resp}"
                                )
                                break

                    if not target:
                        pend_db = self.buscar_pendente_no_banco(
                            cursor, nf_resp, placa_resp, msg.SentOn
                        )

                        if pend_db:
                            pend_id, _pend_assunto, _pend_data, _pend_nf = pend_db
                            cursor.execute(
                                """
                                UPDATE trocas
                                SET status = 'CONCLUÍDO',
                                    data_hora_conclusao = ?,
                                    assistente = ?,
                                    justificativa = ?
                                WHERE id = ? AND status = 'PENDENTE'
                                """,
                                (
                                    msg.SentOn.strftime("%d/%m/%Y %H:%M:%S"),
                                    msg.SenderName,
                                    "-",
                                    pend_id,
                                ),
                            )
                            self.registrar_evento_diagnostico(
                                cursor,
                                f"RECONCILIADA_DB_{msg.EntryID}",
                                "RESPOSTA_RECONCILIADA_DB",
                                msg.Subject,
                                msg.SenderName,
                                cid,
                                nf_resp,
                                placa_resp,
                                f"Resposta reconciliada com pendência histórica ID {pend_id}",
                            )
                            self.log(
                                f"Resposta reconciliada no banco para pendência histórica (ID {pend_id})."
                            )
                            continue

                        respostas_sem_match.append(
                            {
                                "cid": cid,
                                "placa": placa_resp,
                                "nf": nf_resp,
                                "sent_on": msg.SentOn,
                                "subject": msg.Subject,
                                "assistente": msg.SenderName,
                            }
                        )
                        self.registrar_evento_diagnostico(
                            cursor,
                            f"SEM_VINCULO_{msg.EntryID}",
                            "RESPOSTA_SEM_VINCULO",
                            msg.Subject,
                            msg.SenderName,
                            cid,
                            nf_resp,
                            placa_resp,
                            "Resposta de assistente sem correspondência automática",
                        )
                        self.log(
                            f"Sem vínculo imediato para resposta do assistente: {msg.Subject}"
                        )
                        continue

                    # Captura conclusão de assistente mesmo sem anexo ou sem XML na solicitação
                    corpo = msg.Body or ""
                    match_just = re.search(
                        r"(?:justificativa|motivo)s?[\s\-\:]*(.*)",
                        corpo,
                        re.IGNORECASE | re.DOTALL,
                    )
                    if match_just:
                        linhas_texto = [
                            linha.strip()
                            for linha in match_just.group(1).split("\n")
                            if linha.strip()
                        ]
                        texto_just = linhas_texto[0][:150] if linhas_texto else "-"
                    else:
                        texto_just = "-"

                    if anexo_alvo:
                        nome_arq_unico = f"{msg.EntryID[:15]}_{anexo_alvo.FileName}"
                        path = os.path.join(PASTA_ANEXOS, nome_arq_unico)
                        anexo_alvo.SaveAsFile(path)
                        dados_saida = self.ler_xml(path)
                    else:
                        dados_saida = {
                            "nf": "N/D",
                            "cnpj_destinatario": "N/D",
                            "nome_destinatario": "N/D",
                            "transportadora": "N/D",
                            "cfop": "N/D",
                        }

                    target["concl"] = {
                        "data": msg.SentOn.strftime("%d/%m/%Y %H:%M:%S"),
                        "assistente": msg.SenderName,
                        "nf": dados_saida["nf"],
                        "just": texto_just,
                        "cnpj_d_sai": dados_saida["cnpj_destinatario"],
                        "nome_d_sai": dados_saida["nome_destinatario"],
                        "transp_sai": dados_saida["transportadora"],
                        "cfop_sai": dados_saida["cfop"],
                        "tem_xml": tem_xml,
                    }

            # Segunda tentativa de vínculo para respostas que chegaram fora da ordem esperada
            if respostas_sem_match:
                reconciliadas = 0
                for r in respostas_sem_match:
                    target = None
                    for t in reversed(transacoes_ativas):
                        t_data_naive = self.to_naive(t["data_solic_dt"])
                        r_sent_naive = self.to_naive(r["sent_on"])
                        if (
                            t["concl"] is None
                            and t["cid"] == r["cid"]
                            and t_data_naive <= r_sent_naive
                        ):
                            target = t
                            break

                    if not target and r["placa"]:
                        for t in reversed(transacoes_ativas):
                            t_data_naive = self.to_naive(t["data_solic_dt"])
                            r_sent_naive = self.to_naive(r["sent_on"])
                            if (
                                t["concl"] is None
                                and t["placa"]
                                and t["placa"] == r["placa"]
                                and t_data_naive <= r_sent_naive
                            ):
                                target = t
                                break

                    if not target and r["nf"]:
                        for t in reversed(transacoes_ativas):
                            t_data_naive = self.to_naive(t["data_solic_dt"])
                            r_sent_naive = self.to_naive(r["sent_on"])
                            if (
                                t["concl"] is None
                                and t.get("nf_ref") == r["nf"]
                                and t_data_naive <= r_sent_naive
                            ):
                                target = t
                                break

                    if target:
                        target["concl"] = {
                            "data": r["sent_on"].strftime("%d/%m/%Y %H:%M:%S"),
                            "assistente": (
                                r["assistente"] if r.get("assistente") else "-"
                            ),
                            "nf": r["nf"] if r["nf"] else "N/D",
                            "just": "-",
                            "cnpj_d_sai": "N/D",
                            "nome_d_sai": "N/D",
                            "transp_sai": "N/D",
                            "cfop_sai": "N/D",
                            "tem_xml": False,
                        }
                        reconciliadas += 1

                self.log(
                    f"Reconciliação pós-leitura: {reconciliadas} de {len(respostas_sem_match)} respostas sem vínculo foram associadas."
                )

            self.lbl_etapa.config(text="Sincronizando registros no Banco de Dados...")
            self.root.update_idletasks()

            registros_historico_preservados = 0

            for t in transacoes_ativas:
                s = t["solic"]
                c = t["concl"]
                cnpj_d = (
                    c["cnpj_d_sai"] if (c and c["cnpj_d_sai"] != "N/D") else s["cnpj_d"]
                )
                nome_d = (
                    c["nome_d_sai"] if (c and c["nome_d_sai"] != "N/D") else s["nome_d"]
                )

                if not c:
                    padrao = "SIM" if s["tem_xml"] else "NÃO (Armazém)"
                else:
                    if s["tem_xml"] and c["tem_xml"]:
                        padrao = "SIM"
                    elif not s["tem_xml"] and not c["tem_xml"]:
                        padrao = "NÃO (Ambos)"
                    elif not s["tem_xml"]:
                        padrao = "NÃO (Armazém)"
                    else:
                        padrao = "NÃO (Assistente)"

                cursor.execute(
                    "SELECT id, status FROM trocas WHERE entry_id = ?", (s["entry_id"],)
                )
                row = cursor.fetchone()

                if row:
                    db_id, db_status = row

                    if MODO_ESTRITO_HISTORICO and db_status == "CONCLUÍDO":
                        registros_historico_preservados += 1
                        continue

                    if c:
                        # --- CAPTURA DE ALERTA: SE O STATUS ERA PENDENTE, ENTÃO É UMA JUSTIFICATIVA INÉDITA ---
                        if db_status == "PENDENTE" and c["just"] != "-":
                            alertas_justificativas.append(
                                f"• NF: {c['nf']} ({c['assistente']}) - Motivo: {c['just']}"
                            )

                        cursor.execute(
                            """UPDATE trocas SET 
                                          assunto=?, armazem=?, data_hora_solicitacao=?, nf_entrada=?, emitente=?, material=?, cfop=?, volume=?, qvol=?,
                                          status="CONCLUÍDO", data_hora_conclusao=?, assistente=?, nf_saida=?, 
                                          cnpj_emitente=?, cnpj_destinatario=?, justificativa=?, nome_destinatario=?, 
                                          transportadora_saida=?, cfop_saida=?, padronizado_xml=? 
                                          WHERE id=?""",
                            (
                                s["assunto"],
                                s["armazem"],
                                s["data"],
                                s["nf"],
                                s["emit"],
                                s["mat"],
                                s["cfop"],
                                s["vol"],
                                s["qvol"],
                                c["data"],
                                c["assistente"],
                                c["nf"],
                                s["cnpj_e"],
                                cnpj_d,
                                c["just"],
                                nome_d,
                                c["transp_sai"] if c else "N/D",
                                c["cfop_sai"] if c else "N/D",
                                padrao,
                                db_id,
                            ),
                        )
                    else:
                        cursor.execute(
                            """UPDATE trocas SET 
                                          assunto=?, armazem=?, data_hora_solicitacao=?, nf_entrada=?, emitente=?, material=?, cfop=?, volume=?, qvol=?,
                                          cnpj_emitente=?, cnpj_destinatario=?, nome_destinatario=?
                                          WHERE id=?""",
                            (
                                s["assunto"],
                                s["armazem"],
                                s["data"],
                                s["nf"],
                                s["emit"],
                                s["mat"],
                                s["cfop"],
                                s["vol"],
                                s["qvol"],
                                s["cnpj_e"],
                                s["cnpj_d"],
                                s["nome_d"],
                                db_id,
                            ),
                        )
                else:
                    # --- CAPTURA DE ALERTA: SE A LINHA É NOVA E JÁ TEM CONCLUSÃO ---
                    if c and c["just"] != "-":
                        alertas_justificativas.append(
                            f"• NF: {c['nf']} ({c['assistente']}) - Motivo: {c['just']}"
                        )

                    cursor.execute(
                        """INSERT INTO trocas (assunto, armazem, data_hora_solicitacao, nf_entrada, emitente, material, cfop, volume, qvol, status, data_hora_conclusao, assistente, nf_saida, cnpj_emitente, cnpj_destinatario, justificativa, nome_destinatario, transportadora_saida, cfop_saida, padronizado_xml, entry_id) 
                                      VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (
                            s["assunto"],
                            s["armazem"],
                            s["data"],
                            s["nf"],
                            s["emit"],
                            s["mat"],
                            s["cfop"],
                            s["vol"],
                            s["qvol"],
                            "CONCLUÍDO" if c else "PENDENTE",
                            c["data"] if c else "-",
                            c["assistente"] if c else "-",
                            c["nf"] if c else "-",
                            s["cnpj_e"],
                            cnpj_d,
                            c["just"] if c else "-",
                            nome_d,
                            c["transp_sai"] if c else "N/D",
                            c["cfop_sai"] if c else "N/D",
                            padrao,
                            s["entry_id"],
                        ),
                    )

            if MODO_ESTRITO_HISTORICO and registros_historico_preservados > 0:
                self.log(
                    f"Modo estrito ativo: {registros_historico_preservados} registro(s) já concluído(s) foram preservados sem alteração."
                )

            if not MODO_CONSERVADOR_POWER_BI:
                # Consolida pendências duplicadas (mesma NF e mesmo dia) quando já há um registro concluído
                cursor.execute(
                    """
                    SELECT id, nf_entrada, substr(data_hora_solicitacao, 1, 10)
                    FROM trocas
                    WHERE status = 'PENDENTE' AND nf_entrada IS NOT NULL AND nf_entrada <> 'N/D'
                    """,
                )
                pendentes_nf = cursor.fetchall()

                for pend_id, pend_nf, pend_data in pendentes_nf:
                    cursor.execute(
                        """
                        SELECT data_hora_conclusao, assistente, nf_saida, justificativa
                        FROM trocas
                        WHERE status = 'CONCLUÍDO'
                          AND nf_entrada = ?
                          AND substr(data_hora_solicitacao, 1, 10) = ?
                        ORDER BY id DESC
                        LIMIT 1
                        """,
                        (pend_nf, pend_data),
                    )
                    concluida_ref = cursor.fetchone()

                    if concluida_ref:
                        data_conc, assistente_conc, nf_saida_conc, just_conc = (
                            concluida_ref
                        )
                        cursor.execute(
                            """
                            UPDATE trocas
                            SET status = 'CONCLUÍDO',
                                data_hora_conclusao = ?,
                                assistente = ?,
                                nf_saida = ?,
                                justificativa = ?
                            WHERE id = ?
                            """,
                            (
                                data_conc if data_conc else "-",
                                assistente_conc if assistente_conc else "-",
                                nf_saida_conc if nf_saida_conc else "-",
                                just_conc if just_conc else "-",
                                pend_id,
                            ),
                        )

            # Diagnóstico de pendências sem retorno por tempo elevado
            cursor.execute("""
                SELECT entry_id, assunto, armazem, data_hora_solicitacao, nf_entrada
                FROM trocas
                WHERE status = 'PENDENTE'
                """)
            for (
                pend_entry_id,
                pend_assunto,
                pend_armazem,
                pend_data,
                pend_nf,
            ) in cursor.fetchall():
                dt_pend = self.parse_data_hora(pend_data)
                if not dt_pend:
                    continue
                if (hoje - dt_pend).total_seconds() >= HORAS_ALERTA_PENDENTE * 3600:
                    self.registrar_evento_diagnostico(
                        cursor,
                        f"PENDENTE_SEM_RETORNO_{pend_entry_id}",
                        "PENDENTE_SEM_RETORNO",
                        pend_assunto,
                        pend_armazem,
                        "-",
                        pend_nf,
                        self.extrair_placa(pend_assunto or ""),
                        f"Pendente há mais de {HORAS_ALERTA_PENDENTE} horas",
                    )

            conn.commit()

            self.lbl_etapa.config(text="Atualizando base de dados Excel...")
            self.root.update_idletasks()

            df = pd.read_sql_query("SELECT * FROM trocas", conn)
            df.drop(columns=["id", "entry_id"], inplace=True, errors="ignore")
            df.to_excel(ARQUIVO_FINAL, index=False)

            self.lbl_etapa.config(
                text="Rotina concluída com sucesso.", foreground="#0078D4"
            )
            self.log("Processo finalizado. Arquivo Excel atualizado.")

            # ================== EXIBIÇÃO DOS ALERTAS ==================
            if alertas_justificativas:
                qtd_alertas = len(alertas_justificativas)
                self.log(
                    f"AVISO: {qtd_alertas} nova(s) justificativa(s) detectada(s) nesta rodada."
                )

                # Monta a mensagem formatada para o Pop-up
                msg_alerta = f"Atenção, {qtd_alertas} nova(s) justificativa(s) de troca de nota foi/foram detectada(s):\n\n"

                # Limita a exibição para não quebrar a tela se houver 50 justificativas de uma vez
                msg_alerta += "\n".join(alertas_justificativas[:15])
                if qtd_alertas > 15:
                    msg_alerta += f"\n\n... e mais {qtd_alertas - 15} justificativas não listadas aqui. Verifique o Excel."

                # Dispara o pop-up com segurança na Thread principal (evita travamento do Tkinter)
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Novas Justificativas Detectadas", msg_alerta
                    ),
                )
            # =========================================================

        except Exception as e:
            self.lbl_etapa.config(
                text="Falha durante a execução.", foreground="#D13438"
            )
            self.log(f"ERRO DE SISTEMA: {str(e)}")
        finally:
            if "conn" in locals():
                conn.close()
            pythoncom.CoUninitialize()
            self.is_running = False

            self.tempo_restante = self.intervalo_min.get() * 60
            self.btn.config(state="normal")
            self.progress["value"] = 0


if __name__ == "__main__":
    root = tk.Tk()
    AppMosaicMaster(root)
    root.mainloop()
