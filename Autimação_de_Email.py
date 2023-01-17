import pyautogui as pa
import pyperclip  #resolve os erros envolvendo caracteres especiais
import time
import pandas as pd

# PAUSE serve para deixar o código mais controlado com relação ao tempo de execução
pa.PAUSE = 1
pa.FAILSAFE = True

'''
1º Passo: entrar no sistema da empresa
    *Abrindo o navegador na mão*
    *descobrindo o ponto do mouse com o recurso position
'''
pa.hotkey('win', 'r')
pa.write('chrome')
pa.press('enter')

#pg.hotkey('ctrl', 't')  # Com o navegador aberto irá abrir uma nova aba
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing')#copiando o link para ser utilizado
pa.hotkey('ctrl', 'v')  # colando na barra de pesquisa
pa.press('enter')

pa.press('enter')

#2º passo: Navegando no sistema
time.sleep(3)  # se for muito rápido aumentar para 5
pa.click(x=526, y=450,clicks =2)  # clicar na pasta exportação

# 3ºpasso: exportar o relatório
time.sleep(2)  # se for muito rápido aumentar para 5
pa.click(x=589, y=549,clicks =1)  # clica na pasta exportação
pa.click(x=1613, y=274, clicks = 1)  # clica nos 3 pontinhos
pa.click(x=1253, y=933, clicks = 2)  # clica no botão de download
time.sleep(5)


# obs: sempre que tiver passando o endereço de um arquivo no pyton usar o modo 'r string' !

# 4º Passo: calcular os indicadores
tabela = pd.read_excel(r'C:\Users\Família\Downloads\Vendas - Dez.xlsx')  # tabela recebe as informações da base de dados
print(tabela)

# como desejamos calcular o faturamento e a quantidade de itens vendidos, temos:
fat = tabela['Valor Final'].sum()
qte = tabela['Quantidade'].sum()

print(f'O faturamento da empresa: {fat} de reais')
print(f'A quantidade vendida de produtos: {qte}')

# 5º passo: #5º passo: enviar email

# abrir o email
pa.hotkey('ctrl', 't')
pyperclip.copy('https://mail.google.com/mail/u/0/#inbox')
pa.hotkey('ctrl', 'v')
pa.press('enter')
time.sleep(3)

# Navegar Pelo email:
pa.click(x=164, y=297, clicks = 1)
time.sleep(2)

# Preencher o email:
'''
pyperclip.copy('arthur.costa.163@ufrn.edu.br')
pa.hotkey('ctrl', 'v')
pa.press('tab')
pyperclip.copy('Teste Automação')
pa.hotkey('ctrl', 'v')
pa.press('tab')

ou podemos fazer da seguinte maneira:
'''
# Destinatário:
pa.write('arthur.costa.163@ufrn.edu.br')
pa.press('tab')
pa.write('arthur_machado19@yahoo.com.br')
pa.press('tab')
pa.press('tab')  # tab para descer para o preenchimento do assunto

# Assunto:
pyperclip.copy('Teste Automação')
pa.hotkey('ctrl', 'v')
pa.press('tab')

# Texto:
texto = f'''
Prezados, bom dia!

O faturamento do mês até agora foi de: {fat} de reais;
O nº de vendas do mês até agora foi de: {qte} de itens.

Segue em anexo o código fonte do programa utilizado para calcular os indicadores e realizar a automação deste processo

by: Arthur Machado, o estagiário de Engenharia Elétrica.
'''
pyperclip.copy(texto)
pa.hotkey('ctrl', 'v')

# Anexando arquivo do código:
pa.click(x=1310, y=878, clicks=1)  # abrir anexo
pa.click(x=1125, y=287, clicks=1)   # selecionar barra de endereço
pa.press('del')  # pressionando a tecla del para apagar o endereço atual
pyperclip.copy(r'D:\Arthur\ProjetosPyCharm\ProjetoAutomacaoPython')  # copiando endereço atual para pesquisar
pa.hotkey('ctrl', 'v')
pa.press('enter')
pa.click(x=983, y=489, clicks=2)  # selecionando o cod fonte

# enviando o email
pa.hotkey('ctrl', 'enter')

# finalizar operação:
time.sleep(5)
pa.click(x=1883, y=10, clicks=1)
