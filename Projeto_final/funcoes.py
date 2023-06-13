import PySimpleGUI as sg
import sys
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


def exibir_janela_cadastro():

    botoes = sg.Column([
        [sg.Button("Entrar", size=(10, 1)),
         sg.Button("Cadastrar", size=(10, 1)),
         sg.Button("Devolução", size=(10, 1)),
         sg.Button("Cancelar", size=(10, 1))]],
         justification='center', )

    layout = [
        [sg.Text("Nome Completo:", font=("SegoeUI", 12)),
         sg.Input(key="-NOME-")],
        [sg.Text("E-mail Válido:", font=("SegoeUI", 12)),
         sg.Input(key="-EMAIL-")],
        [botoes]
    ]

    janela_cadastro = sg.Window("Cadastro Biblioteca Lisboa", layout,
                                resizable=False, size=(500, 150),
                                finalize=True)

    while True:
        event, values = janela_cadastro.read()
        if event in (sg.WIN_CLOSED, "Cancelar"):
            sys.exit()
        if event == "Cadastrar":
            nome_cadastro = values["-NOME-"].strip()
            email_cadastro = values["-EMAIL-"].strip()
            if not nome_cadastro or not email_cadastro:
                sg.popup("Preencha todos os campos!")
            elif not validar_email(email_cadastro):
                sg.popup("Email inválido!")
            elif verificar_usuario_cadastrado(nome_cadastro, email_cadastro):
                sg.popup("Usuário já cadastrado.")
            else:
                cadastrar_usario(nome_cadastro, email_cadastro)
                sg.popup("Usuário cadastrado com sucesso!")
        elif event == "Entrar":
            nome_verificar = values["-NOME-"].strip()
            email_verificar = values["-EMAIL-"].strip()
            if not nome_verificar or not email_verificar:
                sg.popup("Preencha todos os campos!")
            elif verificar_usuario(nome_verificar, email_verificar):
                janela_cadastro.close()
                break
            else:
                sg.popup("Usuário não encontrado!")

        janela_cadastro["-NOME-"].update("")
        janela_cadastro["-EMAIL-"].update("")
        janela_cadastro.finalize()  # Finaliza a janela para permitir operações
        janela_cadastro.un_hide()  # Exibe a janela novamente


def realizar_devolucao(titulo, autor, data, genero):
    layout_devolucao = [
        [sg.Text("Livro a ser devolvido:", font="SegoeUI")],
        [sg.Text(f"Título: {titulo}")],
        [sg.Text(f"Autor: {autor}")],
        [sg.Text(f"Data: {data}")],
        [sg.Text(f"Gênero: {genero}")],
        [sg.Button("Devolver", size=(10, 1)), sg.Button("Cancelar",
                                                        size=(10, 1))]
    ]

    janela_devolucao = sg.Window("Devolução de Livros", layout_devolucao,
                                 finalize=True)

    while True:
        event_devolucao, _ = janela_devolucao.read()

        if event_devolucao == sg.WINDOW_CLOSED or \
           event_devolucao == "Cancelar":
            break
        if event_devolucao == "Devolver":
            # Aqui adiciona a lógica para realizar a devolução do livro,
            # como atualizar o status do livro, enviar um email de confirmação
            sg.popup("Livro devolvido com sucesso!")
            break

    janela_devolucao.close()


def salvar_devolucao(titulo, autor, data, genero, data_devolucao):
    # Cria um DataFrame com as informações do livro e data de devolução
    livro_devolvido = pd.DataFrame({
        'Título': [titulo],
        'Autor': [autor],
        'Data': [data],
        'Gênero': [genero],
        'Data Devolução': [data_devolucao]
    })

    # Salva o DataFrame no arquivo Excel
    try:
        usuarios_df = pd.read_excel('usuarios.xlsx')
        usuarios_df = usuarios_df.append(livro_devolvido, ignore_index=True)
    except FileNotFoundError:
        usuarios_df = livro_devolvido

    usuarios_df.to_excel('usuarios.xlsx', index=False)
