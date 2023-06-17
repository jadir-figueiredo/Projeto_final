import PySimpleGUI as sg
import sys
import re
import pandas as pd
import smtplib
from smtplib import SMTPException


def cadastrar_usuario(nome, email):
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
            print("Email enviado com sucesso!")

    except SMTPException as erro:
        # Trata exceções que possam ocorrer durante o envio
        print(f"Erro ao enviar email: {erro}")


def devolver_livro(nome, email):
    # Lê o arquivo "usuarios.xlsx" e verifica se há informações de livro
    # relacionadas ao usuário atual
    arquivo_excel = "usuarios.xlsx"
    planilha_nome = "Sheet1"
    df = pd.read_excel(arquivo_excel, sheet_name=planilha_nome)

    # Filtra as linhas do usuário atual
    linhas_usuario = df.loc[(df['Nome'] == nome) & (df['Email'] == email)]

    if len(linhas_usuario) == 0 or pd.isna(linhas_usuario['Título'].iloc[0]):
        sg.popup("Você não possui nenhum livro para devolver.")
        return

    livros_a_devolver = []

    for index, row in linhas_usuario.iterrows():
        titulo = row['Título']
        autor = row['Autor']
        data = row['Data']
        genero = row['Gênero']

        livros_a_devolver.append(index)

        if len(livros_a_devolver) > 0:
            df.loc[livros_a_devolver,
                   ['Título', 'Autor', 'Data', 'Gênero']] = ''
            df.to_excel(arquivo_excel, sheet_name=planilha_nome, index=False)
        else:
            sg.popup("Nenhum livro selecionado para devolver.",
                     keep_on_top=True, non_blocking=True)

    # Cria a janela para exibir as informações do livro e a opção de "Devolver"
    layout_devolucao = [
        [sg.Text(f"Título: {titulo}")],
        [sg.Text(f"Autor: {autor}")],
        [sg.Text(f"Data: {data}")],
        [sg.Text(f"Gênero: {genero}")],
        [sg.Button("Devolver")],
        [sg.Button("Cancelar")],
    ]

    janela_devolucao = sg.Window("Devolver Livro", layout_devolucao,
                                 resizable=False, finalize=True)

    while True:
        evento, _ = janela_devolucao.read()
        if evento == sg.WINDOW_CLOSED or evento == "Cancelar":
            break

        if evento == "Devolver":
            # Remove as informações do livro do arquivo "usuarios.xlsx"
            filtro = (df['Nome'] == nome) & (df['Email'] == email)
            colunas = ['Título', 'Autor', 'Data', 'Gênero']
            # Atribui valores vazios às células correspondentes
            df.loc[filtro, colunas] = ''

            df.to_excel(arquivo_excel, sheet_name=planilha_nome, index=False)

            # enviar_email de aviso de devolução
            corpo_dev = f"O {titulo}, foi devolvido com sucesso"
            destinatario_dev = email
            enviar_email(destinatario_dev, "Livro devolvido", corpo_dev)

            sg.popup("Livro devolvido com sucesso!",
                     keep_on_top=True, non_blocking=True)

            sys.exit()

    janela_devolucao.close()


def adicionar_livro_escolhido(nome, email, titulo, autor, data, genero):
    # Cria um DataFrame com as informações do livro
    livro = {
        'Nome': nome,
        'Email': email,
        'Título': titulo,
        'Autor': autor,
        'Data': data,
        'Gênero': genero
    }

    # Salva o DataFrame na planilha "usuarios.xlsx"
    arquivo_excel = "usuarios.xlsx"
    planilha_nome = "Sheet1"

    try:
        df = pd.read_excel(arquivo_excel, sheet_name=planilha_nome)
        # Verifica se já existe uma linha com o mesmo nome e email
        linha_existente = df[(df["Nome"] == nome) & (df["Email"] == email)]

        if linha_existente.empty:
            # Adiciona uma nova linha ao DataFrame
            df = pd.concat([df, pd.DataFrame([livro])], ignore_index=True)

        else:
            # Atualiza a linha existente com as informações do livro
            indice = linha_existente.index[0]
            df.loc[indice] = livro

    except FileNotFoundError:

        df = pd.DataFrame([livro],
                          columns=["Nome",
                                   "Email",
                                   "Título",
                                   "Autor",
                                   "Data",
                                   "Gênero"])
        print("Arquivo 'usuarios.xlsx' não encontrado.\
              Verifique se o arquivo existe.")

    df.to_excel(arquivo_excel, sheet_name=planilha_nome, index=False)


def mostrar_informacoes_livro(titulo, autor, data, genero):
    botao_confirma_dev = [
        [sg.Button("Devolver", size=(10, 1))],
        [sg.Button("Cancelar", size=(10, 1))],
    ]

    layout_confirma_dev = [
        [sg.Text(f"Título: {titulo}")],
        [sg.Text(f"Autor: {autor}")],
        [sg.Text(f"Data: {data}")],
        [sg.Text(f"Gênero: {genero}")],
        [botao_confirma_dev]
    ]

    janela_confirma_dev = sg.Window("Informações do Livro",
                                    layout_confirma_dev, finalize=True,
                                    resizable=False, size=(500, 150))

    while True:
        evento, _ = janela_confirma_dev.read()

        if evento == sg.WINDOW_CLOSED or evento == "Cancelar":
            janela_confirma_dev.close()
            return False
        elif evento == "Devolver":
            janela_confirma_dev.close()
            enviar_email(
                destinatario=["-EMAIL-"],
                assunto="Informações do Livro",
                corpo=f"O livro {titulo} foi devolvido com sucesso."
                )
            return True

    janela_confirma_dev.close()
