import win32com.client
import pythoncom

pythoncom.CoInitialize()

try:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)

    print("=" * 60)
    print("PASTAS NA INBOX:")
    print("=" * 60)

    for i, pasta in enumerate(inbox.Folders):
        print(f"{i+1}. {pasta.Name}")

    print("\n" + "=" * 60)
    print("PASTAS NO NÍVEL SUPERIOR:")
    print("=" * 60)

    try:
        for i, pasta in enumerate(inbox.Parent.Folders):
            print(f"{i+1}. {pasta.Name}")
    except Exception as e:
        print(f"Erro ao listar pastas do nível superior: {e}")

except Exception as e:
    print(f"Erro: {e}")
finally:
    pythoncom.CoUninitialize()
