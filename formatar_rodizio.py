import pandas as pd
from tkinter import Tk
from tkinter import filedialog, messagebox


def selecionar_arquivo():
    """
    Abrir uma janela para que o arquivo Excel seja selecionado
    """
    Tk().withdraw()  # Oculta a janela principal do tkinter
    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
    )
    return caminho


def ler_excel(caminho, aba=None):
    """Função para ler o arquivo Excel com a aba especificada.

    Args:
        caminho: Caminho do arquivo Excel.
        aba: Nome da aba. Se não for fornecida, lê a primeira aba.
    """
    if aba:
        return pd.read_excel(caminho, nome_aba=aba)
    else:
        return pd.read_excel(caminho)


def selecionar_rodizio():
    """
    Permite ao usuário escolher entre as seguintes opções:
    1. Um único rodízio com um único arquivo e aba.
    2. Dois rodízios, podendo ser um único arquivo com duas abas ou dois arquivos separados.
    """
    rodizio_unico = int(input('Será um único rodízio? (1 - Sim / 2 - Não): '))

    if rodizio_unico == 1:
        # Caso seja um único rodízio
        caminho = selecionar_arquivo()
        df = ler_excel(caminho)
        return df
    else:
        # Pergunta ao usuário se vai usar um único arquivo com duas abas ou dois arquivos separados
        rodizio_tipo = int(input("""
        Será um arquivo com duas abas ou dois arquivos separados?
        1 - Um arquivo com duas abas
        2 - Dois arquivos separados
        3 - Um único arquivo com uma única aba para o primeiro rodízio\n"""))

        if rodizio_tipo == 1:
            # Caso seja um arquivo com duas abas
            caminho = selecionar_arquivo()
            aba_1 = input('Digite o nome da aba do primeiro rodízio: ')
            aba_2 = input('Digite o nome da aba do segundo rodízio: ')
            df_1 = ler_excel(caminho, aba_1)
            df_2 = ler_excel(caminho, aba_2)
        elif rodizio_tipo == 2:
            # Caso sejam dois arquivos separados
            caminho_1 = selecionar_arquivo()
            df_1 = ler_excel(caminho_1)

            caminho_2 = selecionar_arquivo()
            df_2 = ler_excel(caminho_2)
        else:
            print('Opção inválida')
            return None

        # Concatena os dois DataFrames (seja de um único arquivo ou arquivos separados)
        df_final = pd.concat([df_1, df_2], ignore_index=True)
        return df_final


def criar_piping(df):
    """
    Função que retorna o dataframe formatado de acordo com o piping.
    """
    qtde_testes = int(input('Digite a qtde de produtos que serão testados: '))
    qtde_piping = int(input("""
    Piping diferente da qtde de testes?
    1 - Sim
    2 - Não\n"""))

    if qtde_piping == 1:
        qtde_testes = int(
            input('Digite a qtde de produtos pra formatação do piping: '))

    colunas_produtos = None
    if qtde_testes == 1:
        colunas_produtos = int(input('Digite o índice do produto: '))
    else:
        colunas_produtos = qtde_testes

    start_script = []

    def formatar_codigos(linha, colunas_produtos):
        codigos_colunas = linha.index[1:colunas_produtos + 1]
        codigos = linha[codigos_colunas]
        return ', '.join(f"'{codigo}'" for codigo in codigos)

    # Iterar por cada do dataframe linha uma vez
    for idx, linha in df.iterrows():
        id_valor = linha['ID']

        # Contemplando só um caso ou múltiplos casos
        # Exemplo: 8 produtos, porém, só os 3 primeiros serão testados.
        if qtde_testes == 1:
            produtos_teste = linha[colunas_produtos]
            codigos_str = f"'{produtos_teste}'"
        else:
            codigos_str = formatar_codigos(linha, colunas_produtos)

        # Adicionar o texto formado por ID
        start_script.append(
            f"if(id == {id_valor}) {{SetTextFormat(CurrQues, {codigos_str})}}")

    return pd.DataFrame(start_script, columns=['Texto formatado'])


def primeira_parte(df, mapeamento):
    """
    Função que retorna o dataframe formatado de acordo com a primeira parte do rodízio
    utilizando o mapeamento para que os códigos sejam trocados de acordo com o respectivo capítulo
    dele no SurveyToGo.
    """
    qtde_testes = int(input('Digite a qtde de produtos que serão testados: '))
    capitulo_seguinte = int(
        input('Digite o número do capítulo que vem após os produtos: '))
    data = []
    for id in df['ID']:
        valores_mapeados = df.loc[df['ID'] == id, df.columns[1:qtde_testes + 1]].applymap(
            mapeamento.get).values[0]
        # o número do capítulo seguinte
        string_formatada = f"var id{
            id}=[{', '.join(map(str, valores_mapeados))}, {capitulo_seguinte}]"
        data.append({'Texto formatado': string_formatada})

    return pd.DataFrame(data, columns=['Texto formatado'])


def segunda_parte(df):
    """
    Função que retorna o dataframe formatado de acordo com a segunda parte do rodízio.
    """
    data = []
    num = int(input('Digite o número da questão com a quant: '))
    for id in df['ID']:
        # Gerar as strings formatadas
        string_formatada = f"if(id == {id}) {{\nExecutionMgr.GotoChapter(id{
            id}[quant])\nSetAnswer(QRef({num}), quant+=1)\n}}"
        data.append({'Formatted Text': string_formatada})

    # Criar o dataframe
    return pd.DataFrame(data, columns=['Formatted Text'])


def rodizio_pipeline(df, mapeamento_produto):
    while True:

        opcao = int(input("""
        Digite a opção:
            1 - Piping/Setagem
            2 - Primeira parte do rodízio
            3 - Segunda parte do rodízio (ExecutionMgr)
            4 - Mudar de arquivo
            5 - Encerrar
        """))

        if opcao == 1:
            start_script = criar_piping(df)
            start_script.to_clipboard(index=False, header=False)
        elif opcao == 2:
            copiar = primeira_parte(
                df, mapeamento_produto)
            copiar.to_clipboard(index=False, header=False)
        elif opcao == 3:
            copiar = segunda_parte(df)
            copiar.to_clipboard(index=False, header=False)
        elif opcao == 4:
            df = selecionar_rodizio()
        elif opcao == 5:
            break
        else:
            print('Essa opção não está disponível')


# Alterar produto_para_capitulo conforme código dos produtos
# e os respectivos capítulos deles no SurveyToGo.
produto_para_capitulo = {
    664: 2,
    825: 5,
    404: 8,
    765: 11,
    862: 14,
    491: 17,
    800: 20,
    353: 23,
}

df = selecionar_rodizio()
# O  DataFrame será o conteúdo do ou dos arquivos Excel.
rodizio_pipeline(df, produto_para_capitulo)
