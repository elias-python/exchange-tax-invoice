import win32com.client

def raio_x_outlook():
    print(">>> INICIANDO RAIO-X DO OUTLOOK...")
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # 1. Onde estou conectado?
    inbox = outlook.GetDefaultFolder(6) # Pasta padrão
    print(f"\n1. PASTA ATUAL: {inbox.Name}")
    print(f"2. CONTA (DONO DA PASTA): {inbox.Parent.Name}")
    print(f"3. TOTAL DE E-MAILS NA PASTA: {inbox.Items.Count}")
    
    # 2. Vamos ler os últimos 5 e-mails SEM FILTRO NENHUM
    print("\n>>> Lendo os 5 últimos e-mails que chegaram nesta pasta (SEM FILTROS):")
    
    mensagens = inbox.Items
    mensagens.Sort("[ReceivedTime]", True) # Do mais novo para o mais velho
    
    for i in range(min(5, mensagens.Count)):
        msg = mensagens[i]
        try:
            print(f"--------------------------------------------------")
            print(f"ASSUNTO: {msg.Subject}")
            print(f"DATA: {msg.ReceivedTime}")
            print(f"TEM ANEXO? {msg.Attachments.Count > 0}")
            
            # Tenta ver quem mandou
            try:
                print(f"DE: {msg.SenderName} ({msg.SenderEmailAddress})")
            except:
                print("DE: Desconhecido")
                
        except Exception as e:
            print(f"Erro ao ler item: {e}")

    print("\n>>> FIM DO RAIO-X")

if __name__ == "__main__":
    raio_x_outlook()