# Encerra completamente a execução do programa
import sys

# Para importar uma biblioteca para trabalhar-mos fora do ambiente de código
import PySimpleGUI as sg

from datetime import datetime, timedelta

# Importa a tabela de livros do módulo 'livros'
from livros import tabela_livros

# Importa as funções de acesso
from funcoes import (
    cadastrar_usario, validar_email,
    verificar_usuario, verificar_usuario_cadastrado
    )

# Importa a função de enviar email
from funcoes import enviar_email


data_atual = datetime.now()

data_devolucao = data_atual + timedelta(days=7)

data_devolucao_str = data_devolucao.strftime("%d/%m/%Y")

# Define o tema visual da interface gráfica
sg.theme("DarkTeal9")

botoes = sg.Column([
    [sg.Button("Entrar", size=(10, 1)),
     sg.Button("Cadastrar", size=(10, 1)),
     sg.Button("Cancelar", size=(10, 1))]],
     justification='center', )

layout = [
    [sg.Text("Nome Completo:", font=("SegoeUI", 12)),
     sg.Input(key="-NOME-")],
    [sg.Text("E-mail Válido:", font=("SegoeUI", 12)),
     sg.Input(key="-EMAIL-")],
    [botoes]
]

janela_cadastro = sg.Window("Cadastro", layout, resizable=False,
                            size=(500, 150), finalize=True)

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

# Cria uma lista de listas com os dados dos livros da tabela
dados_livros = [[livro["Título"], livro["Autor"], livro["Data"],
                livro["Gênero"]] for livro in tabela_livros]

# Ordena os dados dos livros em ordem alfabética de A a Z pelo título
dados_livros = sorted(dados_livros, key=lambda x: x[0])  # type: ignore

# Define os cabeçalhos das colunas da tabela
header = ["Título", "Autor", "Data", "Gênero"]

# Cria uma tabela com os valores e cabeçalhos especificados
tabela = sg.Table(values=dados_livros, headings=header,
                  justification="left", key="-TABLE-",
                  select_mode=sg.TABLE_SELECT_MODE_BROWSE,
                  enable_events=True, col_widths=[30, 20, 12, 15, 10],
                  num_rows=len(dados_livros), bind_return_key=True,
                  auto_size_columns=False,
                  expand_x=True, expand_y=False)

botoes = sg.Column([
    [sg.Button("Escolher", size=(10, 1)),
     sg.Button("Cancelar", size=(10, 1))]],
     justification='center')

# Define o layout da janela com um texto, botões e a tabela
layout = [
    [sg.Text("Escolha seu livro", font="SegoeUI")],
    [tabela], [botoes]
]

# Cria a janela com o título "Biblioteca Lisboa" e o layout especificado
janela_escolha_livros = sg.Window("Biblioteca Lisboa", layout, resizable=False,
                                  finalize=True)
janela_escolha_livros.Maximize()

while True:
    # Lê os eventos e os valores dos elementos da janela
    evento, valores = janela_escolha_livros.read()
    if evento == sg.WINDOW_CLOSED or evento == "Cancelar":
        break
    if evento == "Escolher":
        # Obtém as linhas selecionadas na tabela
        linhas_selecionadas = tabela.SelectedRows
        if len(linhas_selecionadas) == 0:
            sg.popup("Escolha um livro", non_blocking=True)
        else:
            livro_escolhido = [dados_livros[indice]
                               for indice in linhas_selecionadas]
            titulo = livro_escolhido[0][0]
            autor = livro_escolhido[0][1]
            data = livro_escolhido[0][2]
            genero = livro_escolhido[0][3]
            mensagem = (f"Livro escolhido: "
                        f"{titulo}, {autor}, {data}, {genero}\n"
                        f"Prazo de devolucão são de 07 dias.")
            result = sg.popup_ok_cancel(mensagem)

            if result == "OK":
                email_usuario = values["-EMAIL-"]
                assunto_email = "Livro Escolhido"
                corpo_email = (f"Você escolheu o livro: "
                               f"{titulo}, {autor}, {data}, {genero}\n"
                               f"A data de devolução é até o dia: "
                               f"{data_devolucao_str}")
                enviar_email(email_usuario, assunto_email, corpo_email)
                sg.popup("Você recebeu um email confirmando sua escolha,\n"
                         "juntamente com a data da devolução.\n"
                         "BOA LEITURA!")

janela_escolha_livros.close()
