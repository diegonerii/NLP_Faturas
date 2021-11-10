import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import os
import warnings
warnings.filterwarnings('ignore')

root = tk.Tk()
root.withdraw()

mensagem = 'yes'

while mensagem == 'yes':
    caminho_keyvalues = filedialog.askopenfilename()

    df = pd.read_csv(caminho_keyvalues)
    df[df.columns] = df.apply(lambda x: x.str.strip())

    grupo = df[df['key'] == 'Grupo']['value']

    try:
        try:
            lista_keyvalue, lista_table = [], []
            dicionario_keyvalue = ['Conta referente a',
                                   'N° DA INSTALAÇÃO',
                                   'Nome do Cliente',
                                   'CEP:',
                                   'Subgrupo',
                                   'Data da Emissão',
                                   'VENCIMENTO',
                                   'TOTAL A PAGAR (R$)',
                                   'Alíquota',
                                   'ICMS']
            campos_keyvalue = ['Referência',
                               'Instalação',
                               'Cliente',
                               'Estado',
                               'Subgrupo',
                               'Emissão',
                               'Vencimento',
                               'Total Distribuição',
                               'ICMS (%)',
                               'ICMS (R$)']

            # BUSCA OS VALORES DE REFERÊNCIA NA KEY VALUE E FAZ UM ZIP COM A LISTA DE NOVOS NOMES

            for i in range(0, len(dicionario_keyvalue)):
                dicionario_2 = []
                dicionario_2.append(dicionario_keyvalue[i])
                try:
                    df_novo = df[df['key'].isin(dicionario_2)]
                    lista_keyvalue.append(df_novo['value'].iloc[0])
                except:
                    lista_keyvalue.append('0')

            lista_de_tuplas = list(zip(campos_keyvalue, lista_keyvalue))
            novo_dataframe_kv = pd.DataFrame(lista_de_tuplas, columns=['Campo', 'Valor'])

            caminho_table = str(os.path.abspath(caminho_keyvalues))

            # UM FOR PARA BUSCAR CADA TABLE. SE A TABLE CONTIVER UMA COLUNA CHAMADA 'DESCRIÇÃO', ELE VAI RETORNAR

            for i in range(1, 10):
                caminho_table = caminho_table.replace("keyValues", "table-" + str(i))
                caminho_table = caminho_table.replace("table-" + str(i - 1), "table-" + str(i))

                try:
                    df_table = pd.read_csv(caminho_table)
                    df_table = df_table.rename(columns=lambda x: x.strip())
                    #print("TABELA NÚMERO {}".format(i))

                    # if i == 2:
                    if 'DESCRIÇÃO' in df_table.columns:
                        #print('contém a palavra DESCRIÇÃO')
                        df_table_nova = df_table
                except:
                    break

            df_table_nova_consumo = df_table_nova
            grandeza = ['USO SIST. DISTR. (TUSD) ', 'ENERGIA (TE) ']
            kwh = []
            valor_reais = []

            # VAI BUSCAR QUANTIDADE E VALOR DAS LINHAS QUE CONTIVEREM AS GRANDEZAS DESCRITAS NA TABELA

            for i in range(0, len(grandeza)):
                valor = df_table_nova_consumo['QTD kWh'][df_table_nova_consumo['DESCRIÇÃO'] == grandeza[i]].item().rstrip()
                reais = df_table_nova_consumo['VALOR'][df_table_nova_consumo['DESCRIÇÃO'] == grandeza[i]].item().rstrip()
                kwh.append(valor)
                valor_reais.append(reais)

            # DEPOIS SÓ FORMATA OS NÚMEROS PARA FLOAT
            kwh = [x.replace(".", "").replace(",", ".") for x in kwh]
            kwh = [float(x) for x in kwh]

            valor_reais = [x.replace(".", "").replace(",", ".") for x in valor_reais]
            valor_reais = [float(x) for x in valor_reais]

            df_consumo_kwh = pd.DataFrame({"Campo": ["Consumo - TUSD (kWh)", "Consumo - TE (kWh)"], "Valor": kwh})
            df_consumo_reais = pd.DataFrame({"Campo": ["Consumo - TUSD (R$)", "Consumo - TE (R$)"], "Valor": valor_reais})

            df_table_nova_impostos = df_table_nova

            # VAI BUSCAR TODAS AS LINHAS QUE TIVEREM 'COFINS' e 'PIS/PASEP'
            df_table_nova_impostos['teste'] = df_table_nova_impostos['DESCRIÇÃO'].apply(
                lambda x: 1 if len(re.findall(r'COFINS|PIS/PASEP', x.rstrip())) else 0)

            # DEPOIS TRANSFORMA EM FLOAT
            valores_impostos = list(df_table_nova_impostos['VALOR'][df_table_nova_impostos['teste'] == 1])
            valores_impostos = [x.rstrip() for x in valores_impostos]
            valores_impostos = [x.replace(".", "").replace(",", ".") for x in valores_impostos]
            valores_impostos = [float(x) for x in valores_impostos]

            try:
                df_impostos_reais = pd.DataFrame({"Campo": ["PIS/PASEP (R$)", "COFINS (R$)"], "Valor": valores_impostos})
            except:
                df_impostos_reais = pd.DataFrame(
                    {"Campo": ["PIS/PASEP (R$)", "PIS/PASEP - S/ ICMS(R$)", "COFINS (R$)", "COFINS - S/ ICMS(R$)"],
                     "Valor": valores_impostos})

            df_table_nova_impostos_percent = df_table_nova

            # VAI BUSCAR TODAS AS LINHAS QUE TIVEREM O FORMATO (NÚMERO, NÚMERO)
            df_table_nova_impostos_percent['DESCRIÇÃO'] = df_table_nova_impostos_percent['DESCRIÇÃO'].apply(
                lambda x: re.findall(r'([0-9],[0-9]{2})', x.rstrip()) if len(
                    re.findall(r'([0-9],[0-9]{2})', x.rstrip())) else 0)

            valores_impostos = list(
                df_table_nova_impostos_percent['DESCRIÇÃO'][df_table_nova_impostos_percent['DESCRIÇÃO'] != 0])

            # ÀS VEZES VEM MAIS DE UMA LINHA PQ HÁ UMA PARTE DOS IMPOSTOS SEM ICMS, ENTÃO DEVE-SE FAZER O FOR PARA BUSCAR TUDO
            for i in range(0, len(valores_impostos)):
                valores_impostos[i] = valores_impostos[i][0]

            # TRANSFORMA OS RESULTADOS EM FLOAT
            valores_impostos = [x.replace(",", ".") for x in valores_impostos]
            valores_impostos = [float(x) for x in valores_impostos]

            try:
                df_impostos_percent = pd.DataFrame({"Campo": ["PIS/PASEP (%)", "COFINS (%)"], "Valor": valores_impostos})
            except:
                df_impostos_percent = pd.DataFrame(
                    {"Campo": ["PIS/PASEP (%)", "PIS/PASEP - S/ ICMS(%)", "COFINS (%)", "COFINS - S/ ICMS(%)"],
                     "Valor": valores_impostos})

            df_cidade = novo_dataframe_kv

            # BUSCA AS LINHAS COM FORMATO DE CEP (NÚMERO-NÚMERO)
            df_cidade['teste'] = df_cidade['Valor'].apply(
                lambda x: 1 if len(re.findall(r'([0-9]{5}-[0-9]{3})', x.rstrip())) else 0)
            cidade = str(list(df_cidade['Valor'][df_cidade['teste'] == 1])[0])
            cidade = cidade.split(' - ')

            for i in range(0, len(cidade)):
                if len(re.findall(r'.*/.*', cidade[i])) > 0:
                    cidade = re.findall(r'.*/.*', cidade[i])[0]
                    cidade = cidade.split('/')
                    city = cidade[0]
                    estado = cidade[-1]

            df_localidade = pd.DataFrame({"Campo": ["Cidade", "Estado"], "Valor": [city, estado]})

            # TRAZ DE NOVO A TABELA, POIS AS ANTERIORES ESTAVAM TODAS 'SUJAS'
            for i in range(1, 10):
                caminho_table = str(os.path.abspath(caminho_keyvalues))
                caminho_table = caminho_table.replace("keyValues", "table-" + str(i))
                caminho_table = caminho_table.replace("table-" + str(i - 1), "table-" + str(i))

                try:
                    df_table = pd.read_csv(caminho_table)
                    df_table = df_table.rename(columns=lambda x: x.strip())

                    if 'DESCRIÇÃO' in df_table.columns:
                        df_cip = df_table
                except:
                    break

            # PROCURA A LINHA QUE TENHA A PALAVRA 'MUNICIPAL'
            df_cip['teste'] = df_cip['DESCRIÇÃO'].apply(lambda x: 1 if len(re.findall(r'MUNICIPAL', x.rstrip())) else 0)

            cip = str(list(df_cip['VALOR'][df_cip['teste'] == 1])[0])
            cip = cip.rstrip()
            cip = cip.replace(",", ".")
            cip = float(cip)

            df_cip = pd.DataFrame({"Campo": "Iluminação Pública (R$)", "Valor": [cip]});

            novo_dataframe_kv = novo_dataframe_kv.iloc[:, :-1]
            novo_dataframe_kv = novo_dataframe_kv[novo_dataframe_kv['Campo'] != 'Estado']

            df_ambiente = pd.DataFrame({"Campo": "Ambiente de Contratação", "Valor": ["ACR"]})
            df_modalidade = pd.DataFrame({"Campo": "Modalidade Tarifária", "Valor": ["Convencional"]})

            df_gpb = pd.concat([novo_dataframe_kv,
                                df_ambiente,
                                df_modalidade,
                                df_consumo_kwh,
                                df_consumo_reais,
                                df_impostos_reais,
                                df_impostos_percent,
                                df_localidade,
                                df_cip
                                ],
                               ignore_index=True);
            print(df_gpb)
            MsgBox = messagebox.askquestion('Atenção!', 'Você quer exportar o DataFrame?')
            if MsgBox == 'yes':
                export_path = filedialog.askdirectory()
                localizacao_arquivo = export_path + "/DataFrameGrupoB.xlsx"
                check_file = os.path.isfile(localizacao_arquivo)
                if check_file == False:
                    df_gpav.to_excel(localizacao_arquivo)
                else:
                    MsgBox2 = messagebox.askquestion('Atenção!', 'Arquivo já existe. Você quer sobrescrevê-lo?')
                    if MsgBox2 == 'yes':
                        df_gpav.to_excel(localizacao_arquivo)
                    else:
                        pass
            print('Grupo B')
        except:
            print('Grupo A')
            Modalidade = df[df['key'] == 'Modalidade Tarifária']['value']

            lista_keyvalue, lista_table = [], []

            dicionario_keyvalue = ['Conta referente a',
                                   'N° DA INSTALAÇÃO',
                                   'CEP:',
                                   'Subgrupo',
                                   'Modalidade Tarifária',
                                   'VENCIMENTO',
                                   'TOTAL A PAGAR',
                                   'Alíquota',
                                   'ICMS']

            campos_keyvalue = ['Referência',
                               'Instalação',
                               'Estado',
                               'Subgrupo',
                               'Modalidade Tarifária',
                               'Vencimento',
                               'Total Distribuição (R$)',
                               'ICMS (%)',
                               'ICMS (R$)']

            dicionario_table1 = ['DEMANDA PONTA ',
                                 'DEMANDA FORA PONTA INDUTIVA ',
                                 'DEMANDA ÚNICA C/ DESCONTO ',
                                 'CONSUMO ATIVO PONTA TUSD ',
                                 'CONSUMO ATIVO F. PONTA TUSD ']

            campos_dicionario_table1 = ['Demanda Ponta Registrada (kW)',
                                        'Demanda Ponta Faturada (kW)',
                                        'Demanda Ponta (R$)',
                                        'Demanda Fora Ponta Registrada (kW)',
                                        'Demanda Fora Ponta Faturada (kW)',
                                        'Demanda Fora Ponta Faturada (R$)',
                                        'Demanda Única C/ Desconto Registrada (kW)',
                                        'Demanda Única C/ Desconto Faturada (kW)',
                                        'Demanda Única C/ Desconto (R$)',
                                        'Consumo Ativo Ponta TUSD Registrado (kWh)',
                                        'Consumo Ativo Ponta TUSD Faturado (kWh)',
                                        'Consumo Ativo Ponta TUSD (R$)',
                                        'Consumo Ativo F. Ponta TUSD Registrado (kWh)',
                                        'Consumo Ativo F. Ponta TUSD Faturado (kWh)',
                                        'Consumo Ativo F. Ponta TUSD (R$)']

            for i in range(0, len(dicionario_keyvalue)):
                dicionario_2 = []
                dicionario_2.append(dicionario_keyvalue[i])
                try:
                    df_novo = df[df['key'].isin(dicionario_2)]
                    lista_keyvalue.append(df_novo['value'].iloc[0])
                except:
                    lista_keyvalue.append('0')

            lista_keyvalue = [x.replace(".", "").replace(",", ".") for x in lista_keyvalue]
            lista_de_tuplas = list(zip(campos_keyvalue, lista_keyvalue))
            novo_dataframe_kv = pd.DataFrame(lista_de_tuplas, columns=['Campo', 'Valor'])

            caminho_table = str(os.path.abspath(caminho_keyvalues))

            for i in range(1, 14):
                caminho_table = caminho_table.replace("keyValues", "table-" + str(i))
                caminho_table = caminho_table.replace("table-" + str(i - 1), "table-" + str(i))

                try:
                    df_table = pd.read_csv(caminho_table)
                    df_table = df_table.rename(columns=lambda x: x.strip())

                    if 'DESCRIÇÃO' in df_table.columns:
                        for i in range(0, len(dicionario_table1)):
                            dicionario_2 = []
                            dicionario_2.append(dicionario_table1[i])
                            try:
                                df_table_novo = df_table[df_table['DESCRIÇÃO'].isin(dicionario_2)]
                                lista_table.append(df_table_novo['REGISTRADO kW/kWh/kvarh'].iloc[0])
                                lista_table.append(df_table_novo['FATURADO kW/kWh/kvarh'].iloc[0])
                                lista_table.append(df_table_novo['VALOR'].iloc[0])
                                caminho_table_1 = caminho_table
                            except:
                                lista_table.append('0')
                                caminho_table_1 = caminho_table
                except:
                    break

            ##################### MUNICÍPIO ###################
            # FAZ UMA COLUNA DE CHECK PARA VERIFICAR A LINHA QUE TEM A PALAVRA 'MUNICIPAL'
            df_municipal = pd.read_csv(caminho_table_1)
            df_municipal['check'] = df_municipal['DESCRIÇÃO '].apply(
                lambda x: 0 if len(re.findall(r'MUNICIPAL', x)) == 0 else 1)
            df_municipal = df_municipal[df_municipal['check'] == 1]

            # PEGA O VALOR DA COLUNA E TRANSFORMA E FLOAT
            ip = df_municipal.iloc[:, 10].item()
            ip = ip.rstrip()
            ip = float(ip.replace(",", "."))

            # CRIA O DF
            df_municipal = pd.DataFrame({"Campo": "Iluminação Pública (R$)", "Valor": [ip]})


            ##################### EMISSÃO ###################
            df_emissao = df = pd.read_csv(caminho_keyvalues)
            df_emissao['check'] = df_emissao['key'].apply(lambda x: 1 if len(re.findall(r'.miss', str(x))) > 0 else 0)
            df_emissao[df_emissao['check'] == 1]

            emissao = df_emissao['value'][df_emissao['check'] == 1].iloc[0]
            emissao = emissao.rstrip()

            df_emissao = pd.DataFrame({"Campo": "Data de Emissão", "Valor": [emissao]})


            ##################### CIDADE ###################
            # FAZ UMA COLUNA DE CHECK PARA VERIFICAR A LINHA QUE TEM O FORMATO DE CEP #####-### E EXCLUI AS COLUNAS QUE TENHAM 'SACEDO'
            df_cidade = pd.read_csv(caminho_keyvalues)
            df_cidade['check'] = df_cidade['value'].apply(
                lambda x: 0 if len(re.findall(r'[A-Z]/[A-Z]{2} ', str(x))) == 0 else 1)
            df_cidade = df_cidade[
                (df_cidade['check'] == 1) & (df_cidade['key'] != 'Sacedo') & (df_cidade['key'] != 'Sacedo ')]

            # QUEBRA OS VALORES ENCONTRADOS PELO CARACTER '/' PARA CONSEGUIR CIDADE E ESTADO
            cidade = str(df_cidade['value'][df_cidade['check'] == 1]).split('/', 1)
            estado = cidade[1][:2]
            cidade = cidade[0].rsplit(' ')

            # VERIFICAÇÃO CASO O NOME DA CIDADE SEJA COMPOSTO. EX: 'SÃO PAULO'.
            posicao = []

            # NA PRIMEIRA STRING DA LISTA QUE HOUVE NÚMERO, ELE VAI CAPTAR A POSIÇÃO E DEPOIS FAZER O SLICE DA LISTA.
            for i in range(0, len(cidade)):
                if len(re.findall(r'[^0-9]', cidade[i])) > 0:
                    posicao.append(i + 1)
                    break

            cidade = cidade[posicao[0]:]
            cidade = " ".join(cidade)

            # CRIA O DF
            df_cidade = pd.DataFrame({"Campo": ["Cidade", "Estado"], "Valor": [cidade, estado]})


            ##################### CIDADE ###################
            # FAZ UMA COLUNA DE CHECK PARA VERIFICAR A LINHA QUE TEM O FORMATO DE ALÍQUOTA (##,##%)
            df_impostos = pd.read_csv(caminho_table_1)
            df_impostos['check'] = df_impostos['DESCRIÇÃO '].apply(
                lambda x: 0 if len(re.findall(r'([0-9]+,[0-9]+%)', str(x))) == 0 else 1)
            df_impostos = df_impostos[df_impostos['check'] == 1]

            # QUEBRA OS VALORES ENCONTRADOS PELO CARACTER '(' PARA CONSEGUIR SÓ A ALÍQUOTA
            impostos = str(df_impostos['DESCRIÇÃO '][df_impostos['check'] == 1]).rsplit('(')

            # PEGA SÓ A PARTE NUMÉRICA
            aliquotas = []
            for i in range(0, len(impostos)):
                if len(re.findall(r'[0-9]+,[0-9]{2}', impostos[i])) > 0:
                    aliquotas.append(re.findall(r'[0-9]+,[0-9]{2}', impostos[i])[0])

            # TRANSFORMA EM FLOAT
            aliquotas = [float(x.replace(",", ".")) for x in aliquotas]

            # CRIA O DF
            df_impostos = pd.DataFrame(
                {"Campo": ['PIS/PASEP (%)', 'COFINS (%)'], "Valor": [aliquotas[0], aliquotas[2]]})

            novo_dataframe_kv = novo_dataframe_kv.loc[novo_dataframe_kv['Campo'] != 'Estado']

            if Modalidade.item() == 'Verde':
                print('Verde')

                # FAZ UMA COLUNA DE CHECK PARA VERIFICAR A LINHA QUE TEM AS PALAVRAS 'CONSUMO ATIVO'
                df_table_1 = pd.read_csv(caminho_table_1)
                df_table_1['check'] = df_table_1['DESCRIÇÃO '].apply(
                    lambda x: 1 if len(re.findall(r'CONSUMO ATIVO', x)) else 0)

                # PEGA O VALOR DA DESSA LINHA E DA COLUNA 'FATURADO'
                try:
                    registrado = list(df_table_1['FATURADO kW/ kWh/kvarh '][df_table_1['check'] == 1])
                except:
                    registrado = list(df_table_1['FATURADO kW/kWh/kvarh '][df_table_1['check'] == 1])

                # TRANSFORMA EM FLOAT
                registrado = [x.replace('.', '').replace(',', '.').rstrip() for x in registrado]
                registrado = [float(x) for x in registrado]

                # CRIA O DF
                df_consumo_registrado = pd.DataFrame(
                    {"Campo": ["Consumo Ativo Ponta - TUSD(kWh)", "Consumo Ativo Fora Ponta - TUSD (kWh)"],
                     "Valor": registrado})

                # FAZ UMA COLUNA DE CHECK PARA VERIFICAR A LINHA QUE TEM AS PALAVRAS 'CONSUMO ATIVO'
                df_table_1 = pd.read_csv(caminho_table_1)
                df_table_1['check'] = df_table_1['DESCRIÇÃO '].apply(
                    lambda x: 1 if len(re.findall(r'CONSUMO ATIVO.*TE', x)) else 0)

                # PEGA O VALOR DA DESSA LINHA E DA COLUNA 'FATURADO'
                try:
                    registrado = list(df_table_1['FATURADO kW/ kWh/kvarh '][df_table_1['check'] == 1])
                except:
                    registrado = list(df_table_1['FATURADO kW/kWh/kvarh '][df_table_1['check'] == 1])

                # TRANSFORMA EM FLOAT
                registrado = [x.replace('.', '').replace(',', '.').rstrip() for x in registrado]
                registrado = [float(x) for x in registrado]

                # CRIA O DF
                try:
                    df_consumo_registrado_te = pd.DataFrame(
                        {"Campo": ["Consumo Ativo Ponta - TE (kWh)", "Consumo Ativo Fora Ponta - TE (kWh)"],
                         "Valor": registrado})

                    # CRIA A TABELA DE AMBIENTE DE CONTRATAÇÃO
                    df_ambiente = pd.DataFrame({"Campo": "Ambiente de Contratação", "Valor": ["ACR"]})

                except:  # SE FOR LIVRE
                    df_consumo_registrado_te = pd.DataFrame(
                        {"Campo": ["Consumo Ativo Ponta - TE (kWh)", "Consumo Ativo Fora Ponta - TE (kWh)"],
                         "Valor": [0, 0]})

                    df_ambiente = pd.DataFrame({"Campo": "Ambiente de Contratação", "Valor": ["ACL"]})

                # FAZ UMA COLUNA DE CHECK PARA VERIFICAR A LINHA QUE TEM AS PALAVRAS 'CONSUMO ATIVO'
                df_table_1 = pd.read_csv(caminho_table_1)
                df_table_1['check'] = df_table_1['DESCRIÇÃO '].apply(
                    lambda x: 1 if len(re.findall(r'CONSUMO ATIVO', x)) else 0)

                # PEGA O VALOR DA DESSA LINHA E DA COLUNA 'VALOR'
                valor = list(df_table_1['VALOR '][df_table_1['check'] == 1])
                valor = [float(x.rstrip().replace(".", "").replace(",", ".")) for x in valor]

                # CRIA O DF
                df_consumo_valor = pd.DataFrame(
                    {"Campo": ["Consumo Ativo Ponta - TUSD (R$)", "Consumo Ativo Fora Ponta - TUSD (R$)"], "Valor": valor})

                # FAZ UMA COLUNA DE CHECK PARA VERIFICAR A LINHA QUE TEM AS PALAVRAS 'CONSUMO ATIVO'
                df_table_1 = pd.read_csv(caminho_table_1)
                df_table_1['check'] = df_table_1['DESCRIÇÃO '].apply(
                    lambda x: 1 if len(re.findall(r'CONSUMO ATIVO.*TE', x)) else 0)

                # PEGA O VALOR DA DESSA LINHA E DA COLUNA 'VALOR'
                valor = list(df_table_1['VALOR '][df_table_1['check'] == 1])
                valor = [float(x.rstrip().replace(".", "").replace(",", ".")) for x in valor]

                # CRIA O DF
                try:
                    df_consumo_valor_te = pd.DataFrame(
                        {"Campo": ["Consumo Ativo Ponta - TE (R$)", "Consumo Ativo Fora Ponta - TE (R$)"], "Valor": valor})
                except:
                    df_consumo_valor_te = pd.DataFrame(
                        {"Campo": ["Consumo Ativo Ponta - TE (R$)", "Consumo Ativo Fora Ponta - TE (R$)"], "Valor": [0, 0]})

                # FAZ UMA COLUNA DE CHECK PARA VERIFICAR A LINHA QUE TEM AS PALAVRAS 'CONSUMO ATIVO'
                df_table_1 = pd.read_csv(caminho_table_1)
                df_table_1['check'] = df_table_1['DESCRIÇÃO '].apply(lambda x: 1 if len(re.findall(r'DEMANDA', x)) else 0)
                df_table_1 = df_table_1[df_table_1['check'] == 1]

                # FILTRA AS LINHAS QUE NÃO TEM 'DEMANDA CAPACITIVA'
                df_table_1['DESCRIÇÃO '] = df_table_1['DESCRIÇÃO '].apply(
                    lambda x: x if len(re.findall(r'DEMANDA.*CAPACITIVA', x)) == 0 else '')
                df_table_1 = df_table_1[df_table_1['DESCRIÇÃO '] != '']

                # PEGA OS VALORES FATURADOS
                try:
                    registrado = list(
                        df_table_1['FATURADO kW/ kWh/kvarh '][df_table_1['FATURADO kW/ kWh/kvarh '].notnull() == True])
                except:
                    registrado = list(
                        df_table_1['FATURADO kW/kWh/kvarh '][df_table_1['FATURADO kW/kWh/kvarh '].notnull() == True])

                # TRANSFORMA EM FLOAT
                registrado = [float(x.replace(".", "").replace(",", ".")) for x in registrado]

                # CRIA O DF
                try:
                    df_demanda = pd.DataFrame(
                        {"Campo": ["Demanda Única C/ Desconto (kWh)", "Demanda Lei Estadual (kWh)"], "Valor": registrado})
                except:
                    df_demanda = pd.DataFrame({"Campo": ["Demanda Única C/ Desconto (kWh)"], "Valor": registrado})

                #### CRIANDO O DATAFRAME
                df_gpav = pd.concat([novo_dataframe_kv,
                                     df_emissao,
                                     df_ambiente,
                                     df_consumo_registrado,
                                     df_consumo_registrado_te,
                                     df_consumo_valor,
                                     df_consumo_valor_te,
                                     df_demanda,
                                     df_impostos,
                                     df_cidade,
                                     df_municipal
                                     ], ignore_index=True);
                print(df_gpav)
                MsgBox = messagebox.askquestion('Atenção!', 'Você quer exportar o DataFrame?')
                if MsgBox == 'yes':
                    export_path = filedialog.askdirectory()
                    localizacao_arquivo = export_path+"/DataFrameGrupoA-Verde.xlsx"
                    check_file = os.path.isfile(localizacao_arquivo)
                    if check_file == False:
                        df_gpav.to_excel(localizacao_arquivo)
                    else:
                        MsgBox2 = messagebox.askquestion('Atenção!', 'Arquivo já existe. Você quer sobrescrevê-lo?')
                        if MsgBox2 == 'yes':
                            df_gpav.to_excel(localizacao_arquivo)
                        else:
                            pass
            else:
                print('Azul')


                ###### CONSUMO TUSD #####
                # FAZ UMA COLUNA DE CHECK PARA VERIFICAR A LINHA QUE TEM AS PALAVRAS 'CONSUMO ATIVO'
                df_table_1 = pd.read_csv(caminho_table_1)
                df_table_1['check'] = df_table_1['DESCRIÇÃO '].apply(
                    lambda x: 1 if len(re.findall(r'CONSUMO ATIVO.*TUSD', x)) else 0)

                # PEGA O VALOR DA DESSA LINHA E DA COLUNA 'FATURADO'
                try:
                    registrado = list(df_table_1['FATURADO kW/ kWh/kvarh '][df_table_1['check'] == 1])
                except:
                    registrado = list(df_table_1['FATURADO kW/kWh/kvarh '][df_table_1['check'] == 1])

                # TRANSFORMA EM FLOAT
                registrado = [x.replace('.', '').replace(',', '.').rstrip() for x in registrado]
                registrado = [float(x) for x in registrado]

                # CRIA O DF
                df_consumo_registrado = pd.DataFrame(
                    {"Campo": ["Consumo Ativo Ponta - TUSD (kWh)", "Consumo Ativo Fora Ponta - TUSD (kWh)"],
                     "Valor": registrado})



                ####### CONSUMO TE #######
                # FAZ UMA COLUNA DE CHECK PARA VERIFICAR A LINHA QUE TEM AS PALAVRAS 'CONSUMO ATIVO'
                df_table_1 = pd.read_csv(caminho_table_1)
                df_table_1['check'] = df_table_1['DESCRIÇÃO '].apply(
                    lambda x: 1 if len(re.findall(r'CONSUMO ATIVO.*TE', x)) else 0)

                # PEGA O VALOR DA DESSA LINHA E DA COLUNA 'FATURADO'
                try:
                    registrado = list(df_table_1['FATURADO kW/ kWh/kvarh '][df_table_1['check'] == 1])
                except:
                    registrado = list(df_table_1['FATURADO kW/kWh/kvarh '][df_table_1['check'] == 1])

                # TRANSFORMA EM FLOAT
                registrado = [x.replace('.', '').replace(',', '.').rstrip() for x in registrado]
                registrado = [float(x) for x in registrado]

                # CRIA O DF
                try:
                    df_consumo_registrado_te = pd.DataFrame(
                        {"Campo": ["Consumo Ativo Ponta - TE (kWh)", "Consumo Ativo Fora Ponta - TE (kWh)"],
                         "Valor": registrado})

                    # CRIA A TABELA DE AMBIENTE DE CONTRATAÇÃO
                    df_ambiente = pd.DataFrame({"Campo": "Ambiente de Contratação", "Valor": ["ACR"]})

                except:  # SE FOR LIVRE
                    df_consumo_registrado_te = pd.DataFrame(
                        {"Campo": ["Consumo Ativo Ponta - TE (kWh)", "Consumo Ativo Fora Ponta - TE (kWh)"],
                         "Valor": [0, 0]})

                    df_ambiente = pd.DataFrame({"Campo": "Ambiente de Contratação", "Valor": ["ACL"]})

                df_consumo_registrado_te


                ##### CONSUMO ATIVO TUSD REAIS ######
                # FAZ UMA COLUNA DE CHECK PARA VERIFICAR A LINHA QUE TEM AS PALAVRAS 'CONSUMO ATIVO'
                df_table_1 = pd.read_csv(caminho_table_1)
                df_table_1['check'] = df_table_1['DESCRIÇÃO '].apply(
                    lambda x: 1 if len(re.findall(r'CONSUMO ATIVO.*TUSD', x)) else 0)

                # PEGA O VALOR DA DESSA LINHA E DA COLUNA 'VALOR'
                valor = list(df_table_1['VALOR '][df_table_1['check'] == 1])
                valor = [float(x.rstrip().replace(".", "").replace(",", ".")) for x in valor]

                # CRIA O DF
                df_consumo_valor = pd.DataFrame(
                    {"Campo": ["Consumo Ativo Ponta - TUSD (R$)", "Consumo Ativo Fora Ponta - TUSD (R$)"], "Valor": valor})


                #### CONSUMO TE REAIS #####
                # FAZ UMA COLUNA DE CHECK PARA VERIFICAR A LINHA QUE TEM AS PALAVRAS 'CONSUMO ATIVO'
                df_table_1 = pd.read_csv(caminho_table_1)
                df_table_1['check'] = df_table_1['DESCRIÇÃO '].apply(
                    lambda x: 1 if len(re.findall(r'CONSUMO ATIVO.*TE', x)) else 0)

                # PEGA O VALOR DA DESSA LINHA E DA COLUNA 'VALOR'
                valor = list(df_table_1['VALOR '][df_table_1['check'] == 1])
                valor = [float(x.rstrip().replace(".", "").replace(",", ".")) for x in valor]

                # CRIA O DF
                try:
                    df_consumo_valor_te = pd.DataFrame(
                        {"Campo": ["Consumo Ativo Ponta - TE (R$)", "Consumo Ativo Fora Ponta - TE (R$)"], "Valor": valor})
                except:
                    df_consumo_valor_te = pd.DataFrame(
                        {"Campo": ["Consumo Ativo Ponta - TE (R$)", "Consumo Ativo Fora Ponta - TE (R$)"], "Valor": [0, 0]})

                #### DEMANDA KWH #####
                # FAZ UMA COLUNA DE CHECK PARA VERIFICAR A LINHA QUE TEM AS PALAVRAS 'CONSUMO ATIVO'
                df_table_1 = pd.read_csv(caminho_table_1)
                df_table_1['check'] = df_table_1['DESCRIÇÃO '].apply(lambda x: 1 if len(re.findall(r'DEMANDA', x)) else 0)
                df_table_1 = df_table_1[df_table_1['check'] == 1]

                # FILTRA AS LINHAS QUE NÃO TEM 'DEMANDA CAPACITIVA'
                df_table_1['DESCRIÇÃO '] = df_table_1['DESCRIÇÃO '].apply(
                    lambda x: x if len(re.findall(r'DEMANDA.*CAPACITIVA', x)) == 0 else '')
                df_table_1 = df_table_1[df_table_1['DESCRIÇÃO '] != '']

                # PEGA OS VALORES FATURADOS
                try:
                    registrado = list(
                        df_table_1['FATURADO kW/ kWh/kvarh '][df_table_1['FATURADO kW/ kWh/kvarh '].notnull() == True])
                except:
                    registrado = list(
                        df_table_1['FATURADO kW/kWh/kvarh '][df_table_1['FATURADO kW/kWh/kvarh '].notnull() == True])

                # TRANSFORMA EM FLOAT
                registrado = [float(x.replace(".", "").replace(",", ".")) for x in registrado]

                # CRIA O DF
                try:
                    try:
                        df_demanda = pd.DataFrame(
                            {"Campo": ["Demanda Ponta C/ Desconto (kWh)", "Demanda Fora Ponta C/ Desconto (kWh)"],
                             "Valor": registrado})
                    except:
                        df_demanda = pd.DataFrame({"Campo": ["Demanda Ponta C/ Desconto (kWh)",
                                                             "Demanda Fora Ponta C/ Desconto (kWh)",
                                                             "Demanda Lei Estadual Ponta",
                                                             "Demanda Lei Estadual Fora Ponta"], "Valor": registrado})
                except:
                    df_demanda = pd.DataFrame({"Campo": ["Demanda Ponta C/ Desconto (kWh)",
                                                         "Demanda Fora Ponta C/ Desconto (kWh)", "Demanda Lei Estadual"],
                                               "Valor": registrado})

                df_gpaa = pd.concat([novo_dataframe_kv,
                                     df_emissao,
                                     df_ambiente,
                                     df_consumo_registrado,
                                     df_consumo_registrado_te,
                                     df_consumo_valor,
                                     df_consumo_valor_te,
                                     df_demanda,
                                     df_impostos,
                                     df_cidade,
                                     df_municipal
                                     ], ignore_index=True);
                print(df_gpaa)
                MsgBox = messagebox.askquestion('Atenção!', 'Você quer exportar o DataFrame?')
                if MsgBox == 'yes':
                    export_path = filedialog.askdirectory()
                    localizacao_arquivo = export_path + "/DataFrameGrupoA-Azul.xlsx"
                    check_file = os.path.isfile(localizacao_arquivo)
                    if check_file == False:
                        df_gpav.to_excel(localizacao_arquivo)
                    else:
                        MsgBox2 = messagebox.askquestion('Atenção!', 'Arquivo já existe. Você quer sobrescrevê-lo?')
                        if MsgBox2 == 'yes':
                            df_gpav.to_excel(localizacao_arquivo)
                        else:
                            pass
    except:
        pass

    MsgBox = messagebox.askquestion('Atenção!', 'Você quer adicionar mais arquivos?')
    mensagem = MsgBox

