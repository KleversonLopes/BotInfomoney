'''
    Autor: Kleverson C. Lopes
    Data de criação: 10/03/2024
    objetivo:
        - Colhe informações das maiores altas da bolsa de valores através do site infomoney,
        pesquisa o indice e o valor das ações no google finance e gera uma planilha demonstrativa
        enviando um e-mail com a planilha em anexo para uma conta cadastra no arquivo de credenciais
        layout da planilha:
        +------+------------------+-----------------+-----------------+-----------------+
        | ação | indice infomoney | valor infomoney | indice gfinance |  valor gfinance |
        +------+------------------+-----------------+-----------------+-----------------+
'''
# Import for the Web Bot
from botcity.web import WebBot, Browser, By

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

from webdriver_manager.firefox import GeckoDriverManager

import xlsxwriter

URL_INFOMONEY = "https://www.infomoney.com.br"
URL_GOOGLEFINANCE = 'https://www.google.com/finance'
FILE_OUTPUT = 'infomoney.xlsx'
CREDENCIAIS = 'credenciais.json'

def main():
    # Runner passes the server URL, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    ExecutaBotCity()
...

def ExecutaBotCity():
    bot = WebBot()

    # Configure whether or not to run on headless mode
    bot.headless = False

    bot.driver_path = GeckoDriverManager().install()
    bot.browser = Browser.FIREFOX
    bot.browse(URL_INFOMONEY)

    try:
        elemento_tabela = bot.find_element('high',By.ID, 20000, True, True)
        
        aTabela = elemento_tabela.text.split('\n')
        
        tabela = DadosInfomoney(aTabela)

        DadosGoogleFinance(bot, tabela)
            
        GravaPlanilha(tabela)

        EnviaEmail()

    except Exception as E:
        print(f'Erro {E}')
    finally:
        bot.wait(3000)
        bot.stop_browser()
    ...
...

def DadosInfomoney(aTabela):
    tabela = [['AÇÃO','INDICE INFOMONEY','VALOR INFOMONEY', 'INDICE GOOGLE FINANCE', 'VALOR GOOGLE FINANCE']] # Cabeçalho da Tabela
    for linha in aTabela:
        Array = linha.split(" ")
        Array[3] = Array[3].replace(",",".") # Corrige o formato da coluna de valor para transformar em float, substituindo ',' por '.'
        Array[3] = float(Array[3]) # Conversao de string para float
        Array.pop(2) # remove o 'R$' que está na posição 2 do vetor
        tabela.append(Array)
    
    return tabela
...

def DadosGoogleFinance(bot: WebBot, tabela):
    for linha in tabela:
        if linha[0] != 'AÇÃO':  # Ignora a Primeira Linha (Cabecalho)
            Acao = linha[0]
            bot.navigate_to(URL_GOOGLEFINANCE)
            bot.find_element('/html/body/c-wiz[2]/div/div[3]/div[3]/div/div/div/div[1]/input[2]', By.XPATH, 5000).send_keys(Acao)
            bot.enter()
            valor = bot.find_element('//div[@class="YMlKec fxKbKc"]', By.XPATH, 5000)
            transform = valor.text.replace(',', '.').strip() # Troca virgula por ponto e remove espaços
            valor = transform.replace('R$', '') # Remove o R$ que esta na frente
            if valor == '':
                print(f'erro valor em branco - {Acao}')
                valor = '0.00'

            indice = bot.find_element('/html/body/c-wiz[3]/div/div[4]/div/main/div[2]/div[1]/div[1]/c-wiz/div/div[1]/div/div[1]/div/div[2]/div/span[1]/div/div', By.XPATH)
            linha.append(indice.text) # Adiciona o indice da Google Finance
            linha.append(float(valor)) # Adiciona o Valor da Google Finance
...

def GravaPlanilha(tabela):
    workbook = xlsxwriter.Workbook(FILE_OUTPUT) # cria o arquivo de gravação
    worksheet = workbook.add_worksheet('Maiores Altas') #  adiciona uma aba no Excel com o nome Maiores Altas
    col = 0

    for row, data in enumerate(tabela):
        worksheet.write_row(row, col, data)  # Escreve os dados na planilha

    workbook.close()
...


def EnviaEmail():
    from botcity.plugins.email import BotEmailPlugin

    with open(CREDENCIAIS, "r") as file:
        credenciais = json.load(file)

    EMAIL_USER = credenciais['userlogin']
    EMAIL_PASSWORD = credenciais['password']
    EMAIL_DESTINATARIOS = [credenciais['destinatarios']]
    EMAIL_SUBJECT =  'Informe Mensal - Infomoney'
    EMAIL_ANEXOS = [FILE_OUTPUT]

    EMAIL_BODY = '''\
        <html>
            <head>
            </head>
            <body>
                <h1>BotCity</h1>
                <h2>Segue anexo a Planilha com informações das Ações.</h2>
                <h3>Comparativo entre infomoney e google finance</h3>
            </body>
        </html>
    '''

    # Instanciando o plug -in
    email = BotEmailPlugin()

    # Configure SMTP com o servidor Gmail
    email.configure_smtp("imap.gmail.com", 587)

    try:
        # login com a conta de email válida
        email.login(EMAIL_USER, EMAIL_PASSWORD)

        # Enviando a mensagem de e -mail
        email.send_message(EMAIL_SUBJECT, EMAIL_BODY, EMAIL_DESTINATARIOS, attachments=EMAIL_ANEXOS, use_html=True)

        print("*** E-mail enviado! ***")
    except Exception as e:
        print("Não foi possível enviar o e-mail.", str(e))
    finally:
        # Fecha a conexão com os servidores SMTP
        email.disconnect()        
...

if __name__ == '__main__':
    main()
