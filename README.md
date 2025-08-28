📧 Envio de E-mails via Microsoft Graph com Tkinter
Aplicativo em Python + Tkinter para envio de e-mails usando a Microsoft Graph API. Permite selecionar destinatário(s), preencher assunto, corpo da mensagem em HTML, anexar arquivos de qualquer formato (PDF, PNG, PPTX, DOCX etc.), adicionar assinatura personalizada e incluir logo inline.

✨ Funcionalidades
- Login via Microsoft Account (Azure AD / Microsoft 365).
- Envio de e-mails HTML com suporte a assinatura.
- Suporte a múltiplos anexos de qualquer formato.
- Inserção de logo inline na assinatura (cid:logo).
- Interface gráfica simples e funcional com Tkinter.
- Salvamento automático do e-mail em Itens Enviados.
  
🖼️ Interface
- Campo para destinatário e assunto.
- Editor de mensagem e assinatura HTML.
- Seleção de anexos múltiplos.
- Seleção de logo para assinatura.
- Botão para enviar via Graph API.
  
⚙️ Requisitos
Python 3.9+
Dependências:
pip install msal requests

🔑 Configuração no Azure
1. Acesse o Azure Portal (https://portal.azure.com).
2. Registre um novo App Registration.
3. Copie o Application (client) ID e substitua no código:

CLIENT_ID = "SEU_CLIENT_ID_AQUI"
4. Em API Permissions, adicione:
- Mail.Send → Delegated
5. Dê consentimento ao aplicativo.
  
▶️ Executando
Clone o repositório:
git clone https://github.com/PavorZero/EmailsJornada.git
cd envio-emails-graph
Rode o aplicativo:
python app.py
Na primeira execução, será exibido um Device Code. Acesse a URL fornecida, insira o código e autorize.

📦 Gerar Executável (opcional)
Para gerar um .exe e distribuir sem precisar de Python:
pip install pyinstaller
pyinstaller --noconsole --onefile app.py
O executável estará em dist/app.exe.

🛠️ Tecnologias
- Python
- Tkinter
- Microsoft Graph API
- MSAL Python
  
📄 Licença
Este projeto está sob a licença MIT - veja o arquivo LICENSE para detalhes.

🤝 Contribuição
Sinta-se à vontade para abrir Issues ou enviar Pull Requests 🚀.
