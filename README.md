# 📧 Envio de E-mails via Microsoft Graph com Tkinter

Aplicativo em **Python + Tkinter** para envio de e-mails usando a **Microsoft Graph API**.  
Permite selecionar destinatário(s), preencher assunto, corpo da mensagem em HTML, anexar arquivos de qualquer formato (PDF, PNG, PPTX, DOCX etc.), adicionar assinatura personalizada e incluir logo inline.

---

## ✨ Funcionalidades

- Login via **Microsoft Account** (Azure AD / Microsoft 365).
- Envio de **e-mails HTML** com suporte a assinatura.
- Suporte a **múltiplos anexos** de qualquer formato.
- Inserção de **logo inline** na assinatura (`cid:logo`).
- Interface gráfica simples e funcional com **Tkinter**.
- Salvamento automático do e-mail em **Itens Enviados**.

---

## 🖼️ Interface

- Campo para destinatário e assunto.  
- Editor de mensagem e assinatura HTML.  
- Seleção de anexos múltiplos.  
- Seleção de logo para assinatura.  
- Botão para enviar via Graph API.

---

## ⚙️ Requisitos

- Python **3.9+**
- Dependências:
  ```bash
  pip install msal requests
