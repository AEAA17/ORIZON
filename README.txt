Este código utiliza a biblioteca Selenium para automatizar o preenchimento de um formulário de paciente no site Orizon, otimizando a inserção de dados a partir de uma planilha Excel. O código realiza login no site, preenche o formulário com informações dos pacientes e executa ações conforme necessário.

Configuração e Inicialização:

Importação das Bibliotecas: O código importa as bibliotecas necessárias, incluindo selenium, pandas, e webdriver_manager. Essas bibliotecas são usadas para automação do navegador, manipulação de dados e gerenciamento de drivers do Chrome.

Funções Utilitárias: São definidas funções para esperar e clicar em elementos (wait_and_click) e para esperar e enviar teclas (wait_and_send_keys) no navegador. Essas funções ajudam a garantir que os elementos estejam visíveis e clicáveis antes da interação.

Carregamento da Tabela de Dados:

Leitura da Planilha: O código lê uma planilha Excel que contém dados dos pacientes a serem preenchidos no formulário. A planilha deve estar no formato adequado para o código funcionar corretamente.

Configuração do Navegador:
Inicialização do WebDriver: O WebDriver do Chrome é configurado e iniciado. O navegador é maximizado para facilitar a interação com a página.

Acesso e Login no Site:
Acesso ao Site: O navegador acessa a página de login do site Orizon e realiza login usando as credenciais fornecidas.

Autenticação: O código envia as credenciais de login e clica no botão de login para acessar o sistema.

Preenchimento do Formulário
Loop de Preenchimento: O código percorre as linhas da tabela de dados e preenche o formulário com as informações de cada paciente. O preenchimento inclui:

Troca de Frames: Alterna entre os iframes necessários para acessar os campos do formulário.

Seleção de Operadora: Preenche o campo relacionado à operadora de saúde.

Dados do Beneficiário: Preenche informações como número do cartão, tipo de atendimento, e outros detalhes necessários.

Solicitante: Insere informações do solicitante, como conselho e número do CRM.

Solicitação e Informações Adicionais: Preenche os campos de solicitação e informações adicionais, como indicação clínica e código do procedimento.

Upload de Arquivo: Aguarda o upload manual de um arquivo necessário, que deve ser realizado pelo usuário.

Envio do Formulário: Clica no botão de envio do formulário e solicita uma nova senha para a próxima entrada.

Finalização
Fechamento do Navegador: Após o preenchimento de todos os formulários, o navegador é fechado e o usuário é desconectado do sistema.

Código:

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd

def wait_and_click(driver, by, locator, timeout=10):
    WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((by, locator))
    ).click()

def wait_and_send_keys(driver, by, locator, keys, timeout=10):
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, locator))
    ).send_keys(keys)

# # Carregar a tabela
tabela = pd.read_excel(r"Caminho para sua tabela")

# Configuração do WebDriver
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

# Abrir a página desejada
navegador.get(r'https://polimed.com.br/autenticadorOrizon/loginAutenticador')
navegador.maximize_window()

# Login
wait_and_send_keys(navegador, By.XPATH, '//*[@id="login"]', 'SEU LOGIN' + Keys.ENTER)
wait_and_send_keys(navegador, By.XPATH, '//*[@id="senha"]', 'SUA SENHA' + Keys.ENTER)
wait_and_click(navegador, By.XPATH, '//*[@id="formLogin"]/button')

# Solicitar senha
wait_and_click(navegador, By.XPATH, '//*[@id="solicitacao_senha_externa"]')

#Loop
for linha in tabela.index:
        # Trocar para o iframe
    iframe_1 = WebDriverWait(navegador, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'conteudoIframe'))
    )
    navegador.switch_to.frame(iframe_1)
    iframe = WebDriverWait(navegador, 10).until(
        EC.presence_of_element_located((By.ID, 'frameOOL'))
    )
    navegador.switch_to.frame(iframe)
    
    # Operadora(Escolher o tipo de operadora, BRADESCO SAÚDE ou BRADESCO SAÚDE OPERADORA DE PLANOS )
    #wait_and_send_keys(navegador, By.XPATH, '//*[@id="select_Ems"]', 'BRADESCO SAÚDE-005711')
    wait_and_send_keys(navegador, By.XPATH, '//*[@id="select_Ems"]', 'BRADESCO SAÚDE OPERADORA DE PLANOS-421715')

    # Dados do Beneficiário
    wait_and_click(navegador, By.XPATH, '//*[@id="radio_carteira"]')
    wait_and_send_keys(navegador, By.XPATH, '//*[@id="numeroCarteiraBeneficiario"]',str(tabela.loc[linha, "NUMERO DO CARTAO"]))

    wait_and_click(navegador, By.XPATH, '//*[@id="atendimentoRN"]')
    wait_and_send_keys(navegador, By.XPATH, '//*[@id="atendimentoRN"]', 'Não')

    # Solicitante
    wait_and_click(navegador, By.XPATH, '//*[@id="tipoContratadoSolicitante"]')
    wait_and_send_keys(navegador, By.XPATH, '//*[@id="tipoContratadoSolicitante"]', 'Conselho Profissional' + Keys.ENTER)

    wait_and_click(navegador, By.XPATH, '//*[@id="siglaConselhoContratadoSolicitante"]')
    wait_and_send_keys(navegador, By.XPATH, '//*[@id="siglaConselhoContratadoSolicitante"]', 'Conselho Regional de Medicina')
    wait_and_send_keys(navegador, By.XPATH, '//*[@id="numeroConselho"]', str(tabela.loc[linha, "CRM"]))
    wait_and_send_keys(navegador, By.XPATH, '//*[@id="ufConselhoContratadoSolicitante"]', tabela.loc[linha, "ESTADO CBOS"] + Keys.ENTER)
    wait_and_send_keys(navegador, By.XPATH, '//*[@id="nomeProfissionalSolicitante"]', tabela.loc[linha, "NOME DO MÉDICO"] + Keys.ENTER)
    wait_and_click(navegador, By.XPATH, '//*[@id="botaoReplicaDados"]')
    wait_and_send_keys(navegador, By.ID, 'CBOSprofissionalSolicitante', str(tabela.loc[linha, "CBOS"]) + Keys.ENTER)

    # Solicitação
    wait_and_click(navegador, By.XPATH, '//*[@id="radioCaraterSolicitacaoEletiva"]')

    # Contratado
    wait_and_send_keys(navegador, By.XPATH, '//*[@id="matriculaPrestadorContratadoExecutante"]', 'Sua Matricula' + Keys.ENTER)

    # Info.ad
    wait_and_send_keys(navegador, By.ID, 'indicacaoClinica', str(tabela.loc[linha, "INDICAÇÃO CLÍNICA"]) + Keys.ENTER)

    # Procedimento
    wait_and_send_keys(navegador, By.XPATH, '//*[@id="codigoProcedimento"]', 'Codigo do seu procedimento' + Keys.ENTER)
    wait_and_click(navegador, By.XPATH, '//*[@id="AdicionarProcedimento"]')
    wait_and_send_keys(navegador, By.XPATH, '//*[@id="codigoProcedimento"]', 'Codigo do seu procedimento' + Keys.ENTER)
    wait_and_click(navegador, By.ID, 'tdUploadArquivo')

    time.sleep(15)#(Aqui é aberto, uma janela do explorador de arquivos, não é possivel controlar ela, escolha e suba seu arquivo, e deixe o código rodar)
    
    # Executar envio
    navegador.find_element(By.ID, 'enviar').click()

    # Solicitar nova senha
    navegador.refresh()

#Fechar o navegador
navegador.find_element(By.ID, 'sair').click()

