import openpyxl
import pyperclip
import pyautogui
from time import sleep
# Entrar na planilha

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# Copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
   nome_produto = linha[0].value
   pyperclip.copy(nome_produto)
   pyautogui.click(1170,410,duration=1)
   pyautogui.hotkey('ctrl','v')

   descrição = linha[1].value
   pyperclip.copy(descrição)
   pyautogui.click(1152,480, duration=1)
   pyautogui.hotkey('ctrl','v')

   categoria = linha[2].value
   pyperclip.copy(categoria)
   pyautogui.click(1134,557, duration= 1)
   pyautogui.hotkey('ctrl','v')

   cod_prod = linha[3].value
   pyperclip.copy(cod_prod)
   pyautogui.click(1106,608, duration= 1)
   pyautogui.hotkey('ctrl','v')

   peso = linha[4].value
   pyperclip.copy(peso)
   pyautogui.click(1084,673, duration = 1)
   pyautogui.hotkey('ctrl','v')

   dimensoes = linha[5].value
   pyperclip.copy(dimensoes)
   pyautogui.click(1078,726, duration =1)
   pyautogui.hotkey('ctrl','v')
   pyautogui.click(1005,769, duration =1)
   sleep(3)

   preco = linha[6].value
   pyperclip.copy(preco)
   pyautogui.click(1089,424, duration=1)
   pyautogui.hotkey('ctrl','v')

   quant_estoque = linha[7].value
   pyperclip.copy(quant_estoque)
   pyautogui.click(1066,484, duration=1)
   pyautogui.hotkey('ctrl','v')

   data_val = linha[8].value
   pyperclip.copy(data_val)
   pyautogui.click(1030,537, duration=1)
   pyautogui.hotkey('ctrl','v')

   cor = linha[9].value
   pyperclip.copy(cor)
   pyautogui.click(1024,598, duration=1)
   pyautogui.hotkey('ctrl','v')
 
   tamanho = linha[10].value
   if tamanho == 'Pequeno':
      pyautogui.click(1241,656, duration=1)
      pyautogui.click(1087,674, duration=1)
   
   elif tamanho =='Médio':
      pyautogui.click(1241,656, duration=1)
      pyautogui.click(1049, 695, duration=1)
   else:
      pyautogui.click(1241,656, duration=1)
      pyautogui.click(1142,713, duration=1)

   material = linha[11].value
   pyperclip.copy(material)
   pyautogui.click(1109,711,duration=1)
   pyautogui.hotkey('ctrl','v')
   pyautogui.click(1013,752,duration=1)
   sleep(3)
   

   fabricante = linha[12].value
   pyperclip.copy(fabricante)
   pyautogui.click(1045,437, duration=1)
   pyautogui.hotkey('ctrl','v')

   pais_origem = linha[13].value
   pyperclip.copy(pais_origem)
   pyautogui.click(1040,496, duration=1)
   pyautogui.hotkey('ctrl','v')

   observacoes = linha[14].value
   pyperclip.copy(observacoes)
   pyautogui.click(1027,558, duration=1)
   pyautogui.hotkey('ctrl','v')

   cod_barras = linha[15].value
   pyperclip.copy(cod_barras)
   pyautogui.click(1020,642, duration=1)
   pyautogui.hotkey('ctrl','v')

   loc_armazem = linha[16].value
   pyperclip.copy(loc_armazem)
   pyautogui.click(1016,699,duration=1)
   pyautogui.hotkey('ctrl','v')
   pyautogui.click(1011,741,duration=1)
   pyautogui.click(1583,183,duration=1)
   pyautogui.click(1407,591,duration=1)
   sleep(5)
# Repetir esses passos para outros campos até preencher campos daquela pagina
# Clicar em Proximo 
# Repetir os mesmos passos e ir para a proxima página(pagina 2 )
# Repetir os mesmos passos e finalizar o cadastro daquele produto e clicar em concluir 
# Clicar em ok, para finalizar o processo.
# Clicar no ok mais uma vez na mensagem de confirmação de salvamento no banco de dados.
# Clicar em "adicionar mais um e repetir o processo até finalizar a planilha"

