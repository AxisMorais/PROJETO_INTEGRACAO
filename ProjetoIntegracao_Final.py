#IMPORTANDO AS BIBLIOTECAS NECESSÁRIAS
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import pandas as pd
from pandas import DataFrame
from selenium.webdriver.common.by import By
import os
#--------------------------------------------------------------------------------------------

for x in range(0, 2031):

    # VARIAVEL QUE IDENTIFICA O INICIO DA EXECUÇAO PARA FINS DE ANALISE DE PERFORMANCE
    start = time.time()

    #VARIAVEL DIRETORIO ARMAZENA A PLANTILHA PLAN_1
    diretorio = 'C:/Users/pres00310855/Desktop/Integracao/Plan_1.xlsx'

    #ARMAZENADOR TEM A FUNCAO DE ARMAZENAR A LEITURA DO ARQUIVO PLAN_1
    armazenador = pd.read_excel(diretorio)

    #TRANSFORMANDO A PLANILHA EM UM DATA FRAME (PAINEL DE DADOS)
    data_frame = pd.DataFrame(armazenador)

    #CRIACAO DE UMA VARIAVEL PARA CONTAR QUANTAS UNIADES FORAM VINCULADAS
    vinculados = 0

    # O VALOR DA LINHA DO DATA FRAME SERA O MESMO VALOR D QUANRIDADE DE  REPETIÇAO DA ESTRUTURA FOR
    valorLinha = x

    #ACESSANDO O CODIGO DE LOTAÇAO UTILIZANDO O VALOR DA LINHA  NA COLUNA: COLDIGO DO ORGAO INTERNO DO DATA FFRAME
    lotacao = data_frame.at[valorLinha, 'Código Orgão Interno']

    #ARMAZENANDO O CENTRO DE CUSTO
    centroCusto = data_frame.at[valorLinha, 'Nome Orgão Interno']

    # ARMAZENANDO A SENHA REFERENTE AO CENTRO DE CUSTO
    sigla = data_frame.at[valorLinha, 'Sigla Orgão Interno']

    #TRANSFORMANDO O CODIGO DE LOTAÇAO EM UMA STRING
    lotacao = str(lotacao)


    #LOTACAO FORMATADO GERA UMA MASCARA DE FORMATACAO PARA O CODIGO DE  LOTAÇAO INSERINDO OS PONTOS A CADA DOIS DIGITOS
    lotacaoFormatado = lotacao[0] + lotacao[1] + "." + lotacao[2] + lotacao[3] + "." + lotacao[4] + lotacao[5] + "." + \
                       lotacao[6] + lotacao[7] + "." + lotacao[8] + lotacao[9] + "." + lotacao[10] + lotacao[11] + lotacao[
                           12]

    # ABRIR O NAVEGADOR
    navegador = webdriver.Chrome()

    # ACESSAR O SITE DO CIC:
    navegador.get('http://cic.pbh/')

    # ENCONTRAR O ELEMENTO COM A TAG NAME NO HTML E ESCREVENDO NO CAMPO O LOGIN DO SISTEMA:
    navegador.find_element(By.NAME, 'josso_username').send_keys('thiago.conegundes')

    # ENCONTRAR O ELEMENTO COM A TAG NAME NO HTML E ESCREVER A SENHA PARA ACESSAR O SISTEMA:
    navegador.find_element(By.NAME, 'josso_password').send_keys('----')

    # CLICAR NO BOTAO PARA ACESSAR O SISTEMA
    navegador.find_element(By.CLASS_NAME, "botao").click()

    # CLICAR NA OPCAO GERAL DO MENU SUPERIOR
    navegador.find_element('xpath', '//*[@id="geral"]/div[2]/ul/li[2]/a').click()

    # CLICAR NA OPCAO PARA IDENTIFICAR CODIGO DE INTEGRACAO DE UNIDADE:
    navegador.find_element('xpath', '//*[@id="geral"]/div[2]/ul/li[2]/ul/li[4]/a').click()

    # CLICAR NA OPCAO REGISTRO SEM ASSOCIACAO - RETIFICAR O CODIGO
    navegador.find_element('xpath', '//*[@id="fieldset-fieldsetpesquisasoluinteunidnaoiden"]/div[1]/label[2]').click()

    # TEMPO DE ESPERA DE 2 SEGUNDOS
    time.sleep(2)

    # INSERIR O CODIGO DE LOTACAO
    navegador.find_element('xpath', '//*[@id="codigo_centro_custo"]').send_keys(lotacaoFormatado)

    # TEMPO DE ESPERA DE 2 SEGUNDOS
    time.sleep(3)

    # CLICANDO NA OPCAO PESQUISAR
    navegador.find_element('xpath', '//*[@id="pesquisar"]').click()

    # TEMPO DE ESPERA DE 2 SEGUNDOS
    time.sleep(2)

    # EXECUTANDO TENTATIVAS PARA ACESSAR A TELA DE CADASTRO
    try:
        # CLICANDO NO CODIGO DE LOTACAO PARA ENTRAR NA TELA DE CADASTRO
        navegador.find_element('xpath', '//*[@id="divPesquisa"]/table/tbody/tr/td[1]').click()

        time.sleep(2)

        # CLICANDO NA LUPA PARA BUSCAR O CENTRO DE CUSTO
        navegador.find_element('xpath', '//*[@id="detalhe-1-vinculado"]').click()

        time.sleep(2)

        ##################################################################################################
        # TRABALHANDO COM IFRAMES

        # MUDANDO O ACESSO PARA A PARTE INTERNA DO IFRAME
        navegador.switch_to.frame("ifDlgCentroCusto");

        # ACESSANDO O ELEMENTO DO IFRAME E INSRINDO O CENTRO DE CUSTO
        navegador.find_element('xpath', '//*[@id="nome_centro_custo"]').send_keys(centroCusto)

        # ACESSANDO O ELEMENTO DO IFRAME E INSRINDO A SIGLA
        navegador.find_element('xpath', '//*[@id="sigla_centro_custo"]').send_keys(sigla)

        # INSERINDO A SIGLA
        navegador.find_element('xpath', '//*[@id="pesquisar"]').send_keys(sigla)

        # ESPERANDO 3 SEGUNDOS PARA PESQUISAR O CENTRO DE CUSTO
        time.sleep(3)

        # CLICANDO NO BOTAO PESQUISAR
        navegador.find_element('xpath', '// *[ @ id = "pesquisar"]').click()

        time.sleep(3)
        navegador.find_element('xpath', '// *[ @ id = "conteudo"] / table / tbody / tr / td[1]').click()

        time.sleep(3)

        # SAINDO DO IFRAME, E REDIRECIONANDO PARA O AMBIENTE EXTERNO
        navegador.switch_to.default_content()

        # CLICAR NO BOTAO GRAVAR
        navegador.find_element('xpath', '//*[@id="cancelar"]').click()

        # CONTAGEM DOS VINCULADOS
        vinculados = vinculados + 1

        time.sleep(7)

        ###################################################################################################
    except:
        #CALCULO DE NAO VINCULADOS, COMO ELE NÃO ENCONTROU ELE JÁ RECEBE O NÚMERO 1, ELE NÃO PODE COMEÇAR DO ZERO
        naoVinculados = 1
        naoVinculados = naoVinculados +1

    # VARIÁVEL IDENTIFICA O FIM DA EXECUÇÃO
    end = time.time()

# APRESENTANÇÃO DO RELATÓRIO TEMPO DE EXECUÇÃO E QUATIDADE DE VINCULADOS
print('Tempo Total: ')
print(end - start)
print('--------------')
print('Quantidade de vinculados:', vinculados)
print('-------------------------------------')
print('Quantidade de NÃO vinculados:', naoVinculados)

#CRIACAO DE RELATORIO

# ABRE O ARQUIVO RELATORIO FINAL
arquivo = open('C:/Users/pres00310855/Desktop/Integracao/RelatorioFinal.txt', 'w')

#INSERÇÃO DO TITULO TEMPO TOTAL DE EXECUCACAO
txt_execucao = " TEMPO TOTAL DE EXECUÇÃO: "

#CHAMANDO O METODO PARA ESCREVER O QUE FOI ARMAZENADO NA VARIAVEL TXT_EXECUCAO
arquivo.write(txt_execucao)

#CONVERTENDO A VARIAVEL FINAL DO TEMPO - INCIO DO TEMPO EM STRING
tempo = end - start
tempo = str(tempo)

# INSERINDO O CALCULO DO TEMPO NO RELATORIO
arquivo.write(tempo)

#INSERINDO O TITULO QUANTIDADE DE VINCULADO
txt_vinculado = " \n Quantidade total de vinculados: "
arquivo.write(txt_vinculado)

#INSERINDO O DADO DE QUANTIDADE VINCULADO NO RELATORIO
vinculados = str(vinculados)
arquivo.write(vinculados)

#INSERINDO O TITULO QUANTIDADE DE NAO VINCULADOS
txt_naoVinculado = " \n Quantidade total de Não vinculados: "
arquivo.write(txt_naoVinculado)

#INSERINDO O DADO DE QUANTIDADE DE NAO VINCULADOS NO RELATORIO
naoVinculados = str(naoVinculados)
arquivo.write(naoVinculados)

#FECHANDO O ARQUIVO E FINALIZANDO
arquivo.close()



