import pandas as pd

### Vamos definir uma função para que não precisemos ficar copiando o código
def edic_arquivo(df):
    df_editado = pd.read_excel(df)

    ## Renomenado colunas importantes
    df_editado = df_editado.rename({'Unnamed: 0':'Cod', 'Unnamed: 1': 'Desc',
                            'Unnamed: 2':'NF', "Unnamed: 3": 'Nº Nf',
                            'Unnamed: 4': 'Unidade', 'Unnamed: 9':'Fornecedor',
                            'Unnamed: 10':'Fator','Unnamed: 12':'Quantidade',
                            'Unnamed: 17':'Valor Total'}, axis=1)

    ## Agora, vamos excluir as colunas que não são importantes
    df_editado = df_editado.drop(['Unnamed: 5', "Unnamed: 6", "Unnamed: 7",
                          "Unnamed: 8", "Unnamed: 11",
                          "Unnamed: 13", "Unnamed: 14", "Unnamed: 15",
                          "Unnamed: 16"], axis=1)

    for i in df_editado.index:
        condicao = df_editado['Cod'][i]
        ### Verificamos que os valores 'NaN' são considerados 'float'
        # vamos usar essa condição para excluir linhas que obedeçam essa condição

        ### Em relação às linhas com '-', podemos intuir que se a primeira posição
        # do valor for '-', podemos excluir aquela linha, pois a linha não tem um valor importante
        if type(condicao) == float or '-' in condicao[0]:
            df_editado = df_editado.drop([i])

    df_editado.reset_index(inplace=True, drop=True)

    return df_editado

## Criando a função
def extrair_dados(df):
    ### Verificamos que devido a uma quebra de página do documento
    # vamos precisar instituir 2 condições de "break" de loop (linha 38)
    l_dados = []

    ## Loop para retirar informações
    for i in df.index:
        ## Definindo primeira condição
        if df['NF'][i] == 'NF/Série:':
            ## Definindo variáveis fixas
            nf = df['Nº Nf'][i]
            fornecedor = df['Fornecedor'][i]
            data = df['Desc'][i-1][:10]

            ## Criando um novo df para retirar
            # os produtos da nf encontrada
            df_nf = df[i:]

            ## Criando um novo loop
            for k in df_nf.index:
                ## Definindo uma var código
                cod = df_nf['Cod'][k]
                ## Condição para o código ser válido
                if len(cod) == 6 and cod.isdigit():
                    ## Definindo novas vars
                    unid = df_nf['Unidade'][k]
                    fator = df_nf['Fator'][k]
                    val = df_nf['Valor Total'][k]
                    qtd = df_nf['Quantidade'][k]
                    ## tupla de dados que será transferida para a lista de dados
                    t_dados = (cod, unid, fator, fornecedor, data, qtd, val, nf)

                    ## Add os valores
                    l_dados.append(t_dados)

                ## Condição para finalizar o loop
                #  a primeira condição diz respeito a quando o documento é encerrado
                # sem quebra de página.
                ## A segunda condição é atingida após uma quebra de página
                # o que impede que produtos de uma NF diferente sejam registrados
                # em uma NF de outro fornecedor, já registrada anteriormente
                if cod == 'Total de itens:' or cod == 'Chegada:':
                    break

    ### Criando o DF de acordo com a lista de dados
    df_consolidado = pd.DataFrame(l_dados, columns=['Cod', 'Unid',
                                                    'Fator', 'Fornecedor', "Data",
                                                    'Quantidade', 'Valor Total', 'NF'])

    return df_consolidado

### Vamos criar uma função que recebe um
# arquivo com tipos de materiais e descrição completa
# para deixar o DF mais completo possível
def inclusao_tipo(arquivo):
  df = pd.read_excel(arquivo)

  ### Após analisar os dados acima, vamos excluir as colunas a seguir, pois elas possuem muitos
  # valores nulos, logo, não serão usadas na análise
  df = df.drop(['Unnamed: 2', 'Unnamed: 3',
                'Unnamed: 4', 'Unnamed: 5', 'Unnamed: 6',
                'Unnamed: 7', 'Unnamed: 9', 'Unnamed: 10',
                'Unnamed: 11', 'Unnamed: 13',
                'Unnamed: 14', 'Unnamed: 16','Unnamed: 8',
                'Unnamed: 12', 'Unnamed: 15'], axis=1)

  ### Renomeando as colunas para melhor análise
  df = df.rename(columns={"Unnamed: 0":'Cod', 'Unnamed: 1': 'Descricao'})

  ### Aqui, vamos mudar o tipo de todos os dados da coluna 'Codigo' para str
  ## Pois dessa forma, vamos criar uma lista auxiliar para definir o tipo
  # de todos os produtos do df
  df['Cod'] = df['Cod'].astype("str")

  ### Criando a coluna 'Tipo_Material'
  # Vamos criar uma lista auxiliar para salvar os tipos de materiais de cada produto.
  lista_tipo_material = []

  ### Loop para percorrer todas as linhas
  for linha in df.index:
      dado = df["Cod"][linha]
      if len(dado) > 6:
          if '-' == dado[3] and len(dado) < 50:
              tipo = dado
      if len(dado) == 6:
          lista_tipo_material.append(tipo[:3])

  ### Excluindo valores nulos das 'LINHAS'
  df = df.dropna()

  ## Loop para excluir linhas que não serão
  # usadas na análise
  for linha in df.index:
      if len(df['Cod'][linha]) != 6:
          ## Excluindo linha
          df= df.drop(linha)

  ### Adicionando a coluna tipo de material ao df
  df['Tipo Material'] = lista_tipo_material

  return df

## Podemos colocar tudo em uma função que será responsável por chamar as demais
# e editar mais alguns detalhes do df
def consolidar_dados(arquivo_nf, arquivo_tipo):
    df_dados = edic_arquivo(arquivo_nf)
    df_dados_extraidos = extrair_dados(df_dados)
    df_tipo = inclusao_tipo(arquivo_tipo)

    ### Consolidando os df's em apenas um, usando a coluna "Cod" como ref
    df_consolidado = pd.merge(df_dados_extraidos, df_tipo, on='Cod')

    ## Editando os tipos de dados de algumas colunas
    df_consolidado['Valor Total'] = df_consolidado['Valor Total'].astype(float)
    #df_consolidado['Cod'] = df_consolidado['Cod'].astype(int)
    df_consolidado['Mês'] = pd.to_datetime(df_consolidado['Data'], dayfirst=True)
    df_consolidado['Mês'] = pd.to_datetime(df_consolidado['Mês']).dt.month

    ## Reorganizando a ordem das colunas
    df_consolidado = df_consolidado.iloc[:,[0,8,1,2,9,10,4,3,5,6,7]]

    return df_consolidado

