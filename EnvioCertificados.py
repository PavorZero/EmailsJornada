import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import pandas as pd
import time
import requests
from msal import PublicClientApplication
import base64
import mimetypes

# ==============================
# CONFIG MICROSOFT GRAPH
# ==============================
CLIENT_ID = "SEU ID"
TENANT_ID = "SEU ID"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Mail.Send"]
GRAPH_URL = "https://graph.microsoft.com/v1.0"

app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
access_token = None

def obter_token():
    global access_token
    result = None

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise Exception("Falha ao criar device flow")
        messagebox.showinfo("Login Microsoft", f"Acesse {flow['verification_uri']} e digite o código: {flow['user_code']}")
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        access_token = result["access_token"]
        return True
    else:
        messagebox.showerror("Erro", f"Não foi possível obter token: {result.get('error_description')}")
        return False

# ==============================
# FUNÇÃO ENVIO DE EMAILS
# ==============================
def enviar_email_graph(graph_client, remetente, destinatario, assunto, mensagem, arquivos_anexos, assinatura_html, caminho_logo):
    anexos = []

    # Loop pelos anexos normais (certificados e outros arquivos)
    for caminho in arquivos_anexos:
        if pd.isna(caminho) or not caminho or not os.path.exists(caminho):
            continue  # Pula se o caminho for inválido ou vazio
            
        nome_arquivo = os.path.basename(caminho)
        content_type, _ = mimetypes.guess_type(caminho)
        if content_type is None:
            content_type = "application/octet-stream"

        try:
            with open(caminho, "rb") as f:
                arquivo_base64 = base64.b64encode(f.read()).decode("utf-8")

            anexos.append({
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": nome_arquivo,
                "contentType": content_type,
                "contentBytes": arquivo_base64
            })
        except Exception as e:
            print(f"Erro ao processar anexo {caminho}: {e}")
            continue

    # Adiciona a imagem da assinatura como inline
    try:
        with open(caminho_logo, "rb") as f:
            logo_base64 = base64.b64encode(f.read()).decode("utf-8")

        anexos.append({
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": "assinatura.png",
            "contentType": "image/png",
            "isInline": True,
            "contentId": "logoassinatura",
            "contentBytes": logo_base64
        })
    except Exception as e:
        print(f"Erro ao processar assinatura: {e}")

    # Converter quebras de linha para HTML
    mensagem_html = mensagem.replace('\n', '<br>')
    
    # Adicionar formatação HTML básica
    mensagem_html = f"""
    <div style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.6; color: #333;">
        {mensagem_html}
    </div>
    """

    # Corpo do e-mail com assinatura + logo embutida
    corpo_html = f"""
    <html>
        <head>
            <meta charset="UTF-8">
        </head>
        <body style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.6; color: #333; margin: 0; padding: 20px;">
            {mensagem_html}
            <br>
            <div style="margin-top: 20px; padding-top: 20px; border-top: 1px solid #eee;">
                {assinatura_html}
            </div>
            <br>
            <img src="cid:logoassinatura" style="max-width: 200px; height: auto;">
        </body>
    </html>
    """

    # Monta o payload
    email_msg = {
        "message": {
            "subject": assunto,
            "body": {
                "contentType": "HTML",
                "content": corpo_html
            },
            "from": {
                "emailAddress": {
                    "address": remetente
                }
            },
            "toRecipients": [
                {"emailAddress": {"address": destinatario}}
            ],
            "attachments": anexos
        },
        "saveToSentItems": "true"
    }

    # Envia o e-mail usando a URL completa
    response = graph_client.post(f"{GRAPH_URL}/me/sendMail", json=email_msg)
    return response

# ==============================
# FUNÇÃO PRINCIPAL DE ENVIO
# ==============================
def iniciar_envio():
    # Verificar se todos os campos obrigatórios estão preenchidos
    if not entry_remetente.get().strip():
        messagebox.showerror("Erro", "Por favor, informe o remetente!")
        return
        
    if not assinatura_path:
        messagebox.showerror("Erro", "Selecione uma imagem para assinatura!")
        return
        
    # Obter token de acesso
    if not obter_token():
        return
    
    # Configurar cliente Graph
    graph_client = requests.Session()
    graph_client.headers.update({
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    })
    
    # Selecionar arquivo Excel
    arquivo_excel = filedialog.askopenfilename(
        title="Selecione a planilha Excel",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    
    if not arquivo_excel:
        messagebox.showwarning("Aviso", "Nenhuma planilha selecionada!")
        return
    
    try:
        # Ler dados do Excel
        df = pd.read_excel(arquivo_excel)
        
        # Verificar colunas necessárias
        colunas_necessarias = ['Email', 'Nome', 'Certificado']
        for coluna in colunas_necessarias:
            if coluna not in df.columns:
                messagebox.showerror("Erro", f"Coluna '{coluna}' não encontrada na planilha!")
                return
        
        # Verificar se há coluna de assunto, caso contrário usar padrão
        if 'Assunto' not in df.columns:
            df['Assunto'] = 'Certificado'  # Assunto padrão
        
        # Obter dados da interface
        mensagem = text_corpo.get("1.0", tk.END).strip()
        remetente = entry_remetente.get().strip()
        assinatura_html = text_assinatura.get("1.0", tk.END).strip()
        
        # Criar uma nova janela para mostrar o progresso
        progress_window = tk.Toplevel(root)
        progress_window.title("Progresso do Envio")
        progress_window.geometry("600x500")
        
        progress_label = tk.Label(progress_window, text="Preparando para enviar...")
        progress_label.pack(pady=10)
        
        progress_bar = ttk.Progressbar(progress_window, length=550, mode='determinate')
        progress_bar.pack(pady=10)
        
        log_text = scrolledtext.ScrolledText(progress_window, width=70, height=20)
        log_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        log_text.insert(tk.END, "Iniciando envio de e-mails...\n")
        
        progress_window.update()
        
        # Processar cada linha da planilha
        total_emails = len(df)
        enviados = 0
        erros = 0
        
        for index, row in df.iterrows():
            destinatario = row['Email']
            nome = row['Nome']
            assunto = row['Assunto']
            certificado_path = row['Certificado']
            
            # Verificar se o caminho do certificado existe
            if pd.isna(certificado_path) or not os.path.exists(certificado_path):
                log_text.insert(tk.END, f"ERRO: Certificado não encontrado para {destinatario}\n")
                erros += 1
                continue
            
            # Personalizar mensagem (substituir placeholders)
            mensagem_personalizada = mensagem.replace("{nome}", nome)
            
            # Atualizar progresso
            progresso_atual = (index + 1) / total_emails * 100
            progress_bar['value'] = progresso_atual
            progress_label.config(text=f"Enviando {index + 1} de {total_emails} - {nome}")
            
            # Adicionar log
            log_text.insert(tk.END, f"Enviando para {destinatario} ({nome})... ")
            progress_window.update()
            
            # Preparar anexos: certificado específico + anexos gerais
            anexos_do_email = [certificado_path] + anexos_gerais
            
            # Enviar email
            try:
                response = enviar_email_graph(
                    graph_client=graph_client,
                    remetente=remetente,
                    destinatario=destinatario,
                    assunto=assunto,
                    mensagem=mensagem_personalizada,
                    arquivos_anexos=anexos_do_email,
                    assinatura_html=assinatura_html,
                    caminho_logo=assinatura_path
                )
                
                if response.status_code == 202:
                    log_text.insert(tk.END, "OK\n")
                    enviados += 1
                else:
                    log_text.insert(tk.END, f"ERRO: {response.status_code} - {response.text}\n")
                    erros += 1
                
                # Pequena pausa entre envios
                time.sleep(1)
                
            except Exception as e:
                log_text.insert(tk.END, f"ERRO: {str(e)}\n")
                erros += 1
                continue
        
        # Finalizar
        log_text.insert(tk.END, f"\nProcesso concluído!\nEnviados: {enviados}\nErros: {erros}")
        progress_bar['value'] = 100
        
        # Adicionar botão para fechar a janela de progresso
        btn_fechar = tk.Button(progress_window, text="Fechar", command=progress_window.destroy)
        btn_fechar.pack(pady=10)
        
        # Mostrar resumo
        messagebox.showinfo("Concluído", f"Processo de envio finalizado!\n\nE-mails enviados: {enviados}\nErros: {erros}")
        
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar planilha: {str(e)}")

# ==============================
# INTERFACE TKINTER
# ==============================
root = tk.Tk()
root.title("Envio Automático de Certificados (Microsoft Graph)")
root.geometry("800x900")

# Frame principal
main_frame = tk.Frame(root, padx=10, pady=10)
main_frame.pack(fill=tk.BOTH, expand=True)

# Remetente
frame_remetente = tk.Frame(main_frame)
frame_remetente.pack(fill=tk.X, pady=5)
tk.Label(frame_remetente, text="Remetente:").pack(side=tk.LEFT)
entry_remetente = tk.Entry(frame_remetente, width=50)
entry_remetente.pack(side=tk.LEFT, padx=5)
entry_remetente.insert(0, "seu-email@dominio.com")

# Corpo do email
frame_corpo = tk.Frame(main_frame)
frame_corpo.pack(fill=tk.BOTH, expand=True, pady=5)

tk.Label(frame_corpo, text="Corpo do E-mail (use {nome} para personalizar):").pack(anchor=tk.W)
text_corpo = scrolledtext.ScrolledText(frame_corpo, width=85, height=10)
text_corpo.pack(fill=tk.BOTH, padx=5, pady=5, expand=True)
text_corpo.insert(tk.END, "Prezado(a) {nome},\n\nÉ com grande satisfação que entregamos seu certificado de participação.\n\nAgradecemos sua presença e contribuição para o sucesso do nosso evento.\n\nAtenciosamente,")

# Assinatura em HTML
frame_assinatura_html = tk.Frame(main_frame)
frame_assinatura_html.pack(fill=tk.BOTH, expand=True, pady=5)

tk.Label(frame_assinatura_html, text="Assinatura HTML (use tags HTML para formatação):").pack(anchor=tk.W)
text_assinatura = scrolledtext.ScrolledText(frame_assinatura_html, width=85, height=6)
text_assinatura.pack(fill=tk.BOTH, padx=5, pady=5, expand=True)
text_assinatura.insert(tk.END, """<p style="color: #2c3e50;">
<strong>Equipe de Organização</strong><br>
<em>Evento Corporativo 2024</em><br>
📞 (11) 9999-9999<br>
📧 contato@empresa.com
</p>""")

# Exemplos de formatação HTML
frame_exemplos = tk.Frame(main_frame)
frame_exemplos.pack(fill=tk.X, pady=5)
tk.Label(frame_exemplos, text="Dicas de formatação HTML:", font=("Arial", 9, "bold"), fg="blue").pack(anchor=tk.W)
tk.Label(frame_exemplos, text="<strong>Negrito</strong> • <em>Itálico</em> • <u>Sublinhado</u> • <span style='color: red;'>Cor</span> • <br>Quebra de linha", 
         font=("Arial", 8), fg="gray").pack(anchor=tk.W)

# Anexos gerais (para todos os emails)
frame_anexos = tk.Frame(main_frame)
frame_anexos.pack(fill=tk.X, pady=5)
tk.Label(frame_anexos, text="Anexos Gerais (para todos os emails):").pack(anchor=tk.W)

frame_botoes_anexos = tk.Frame(main_frame)
frame_botoes_anexos.pack(fill=tk.X, pady=5)
btn_anexo = tk.Button(frame_botoes_anexos, text="Adicionar Anexo Geral", command=lambda: adicionar_anexo())
btn_anexo.pack(side=tk.LEFT)
btn_limpar_anexos = tk.Button(frame_botoes_anexos, text="Limpar Anexos", command=lambda: limpar_anexos())
btn_limpar_anexos.pack(side=tk.LEFT, padx=5)

listbox_anexos = tk.Listbox(main_frame, height=4)
listbox_anexos.pack(fill=tk.X, padx=5, pady=5)

# Assinatura (imagem)
frame_assinatura_img = tk.Frame(main_frame)
frame_assinatura_img.pack(fill=tk.X, pady=5)
btn_assinatura = tk.Button(frame_assinatura_img, text="Carregar Imagem de Assinatura", command=lambda: carregar_assinatura())
btn_assinatura.pack(side=tk.LEFT)
lbl_assinatura = tk.Label(frame_assinatura_img, text="Nenhuma assinatura carregada")
lbl_assinatura.pack(side=tk.LEFT, padx=10)

# Informações sobre a planilha
frame_info = tk.Frame(main_frame)
frame_info.pack(fill=tk.X, pady=10)
tk.Label(frame_info, text="A planilha Excel deve conter as colunas: Email, Nome, Certificado (caminho do arquivo) e Assunto (opcional)", 
         font=("Arial", 9), fg="gray").pack()

# Botão de envio
btn_enviar = tk.Button(main_frame, text="Selecionar Planilha e Enviar", command=iniciar_envio, 
                      font=("Arial", 12), padx=20, pady=10, bg="#4CAF50", fg="white")
btn_enviar.pack(pady=20)

# Variáveis globais
anexos_gerais = []  # Anexos que serão enviados para todos
assinatura_path = None

# Funções auxiliares
def adicionar_anexo():
    arquivos = filedialog.askopenfilenames()
    for arq in arquivos:
        anexos_gerais.append(arq)
        listbox_anexos.insert(tk.END, arq)

def limpar_anexos():
    anexos_gerais.clear()
    listbox_anexos.delete(0, tk.END)

def carregar_assinatura():
    global assinatura_path
    assinatura_path = filedialog.askopenfilename(
        filetypes=[("Imagens", "*.png;*.jpg;*.jpeg")]
    )
    if assinatura_path:
        lbl_assinatura.config(text=os.path.basename(assinatura_path))

# Iniciar aplicação

root.mainloop()
