import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx') #Carregar o arquivo
vendas = workbook['vendas'] #Carregar a Planilha

for linha in vendas.iter_rows(min_row=2):
    
    pyautogui.click(1308,596, duration=1) #Cliente
    pyautogui.write(linha[0].value)
   
    pyautogui.click(1301,619, duration=1) #Produto
    pyautogui.write(linha[1].value)
    
    pyautogui.click(1302,640, duration=1) # Quantidade
    pyautogui.write(str(linha[2].value))
    
    pyautogui.click(1306,661,duration=1) # Categoria
    pyautogui.write(linha[3].value)

    pyautogui.click(1217,700, duration=1) # Salvar
    pyautogui.click(770,446, duration=1) # Ok

    # OBSERVAÇÃO: Caso vá usar o programa, certifique-se de que o mouse
    # esteja na coordenada certa. Para verificação, é possível utilizando
    # a biblioteca "mouseinfo"