import re
import pandas as pd
import smtplib
from smtplib import SMTPException


def cadastrar_usario(nome, email):
    # Cria um DataFrame com os dados do usuário
    df_usuario = pd.DataFrame({"Nome": [nome.lower()],
                               "Email": [email.lower()]})

    # Carrega a planilha existente ou cria um novo arquivo
    try:
        df_planilha = pd.read_excel("usuarios.xlsx")
        df_planilha = pd.concat([df_planilha, df_usuario], ignore_index=True)
    except FileNotFoundError:
        df_planilha = df_usuario

    # Salva os dados na planilha
    df_planilha.to_excel("usuarios.xlsx", index=False)


def validar_email(email):
    # Expressão regular para validar o formato do e-mail
    padrao_email = r"[^@]+@[^@]+\.[^@]+"
    return re.match(padrao_email, email) is not None


def verificar_usuario(nome, email):
    # Função para verificar cadastro
    try:
        # Carregar planilha
        df_planilha = pd.read_excel("usuarios.xlsx")
        # Verificar se o usuario está presente na planilha
        usuario_encontrado = df_planilha[(df_planilha["Nome"] == nome.lower())
                                         & (df_planilha["Email"] ==
                                            email.lower())]
        if not usuario_encontrado.empty:
            return True
    except FileNotFoundError:
        pass

    return False  # Se o usuário não for encontrado


def verificar_usuario_cadastrado(nome, email):
    try:
        df_planilha = pd.read_excel("usuarios.xlsx")
        usuario_encontrado = df_planilha[
            (df_planilha["Nome"] == nome.lower()) &
            (df_planilha["Email"] == email.lower())
        ]
        return not usuario_encontrado.empty
    except FileNotFoundError:
        return False


# Função para enviar e-mails
def enviar_email(destinatario, assunto, corpo):
    # Configurações do remetente e servidor SMTP
    remetente = "projetofinal.python@gmail.com"
    senha = "ozcgrtuowtzcnuka"
    servidor_smtp = "smtp.gmail.com"
    porta_smtp = 587

    # Constrói a mensagem de e-mail
    mensagem = (f"Subject: {assunto}\n\n{corpo}")

    try:
        # Estabelece uma conexão com o servidor SMTP
        with smtplib.SMTP(servidor_smtp, porta_smtp) as servidor:
            # Inicia a criptografia da conexão
            servidor.starttls()

            # Realiza a autenticação no servidor
            servidor.login(remetente, senha)

            # Codifica a mensagem para utf-8
            mensagem = mensagem.encode("utf-8")

            # Envia o e-mail
            servidor.sendmail(remetente, destinatario, mensagem)

            # Exibe mensagem de sucesso
            print("E-mail enviado com sucesso!")

    except SMTPException as erro:
        # Trata exceções que possam ocorrer durante o envio
        print(f"Erro ao enviar e-mail: {erro}")
