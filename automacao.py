import pyautogui
import subprocess
import time
import pandas as pd 


url_formulario = "https://forms.office.com/r/PwS0k9yyqs?origin=lprLink"

subprocess.run(r"C:\Program Files\Google\Chrome\Application\chrome.exe")

time.sleep(1)
pyautogui.write(url_formulario)

time.sleep(1)
pyautogui.press('enter')

tabela_cadastro = pd.read_excel("Lista_de_Contatos.xlsx")  

#loc localizar dentro da linha da planilha a coluna
for linha in tabela_cadastro.index:
    nome = tabela_cadastro.loc[linha, "Nome"]
    empresa = tabela_cadastro.loc[linha, "Empresa"]
    profissao = tabela_cadastro.loc[linha, "Profissao"]

        
    
    time.sleep(1)
    pyautogui.press('tab')#para apertar e preencher a primeira input do formuláriopyautogui.press'tab' #para apertar e preencher a primeira input do formulário
    pyautogui.press('tab')
    pyautogui.press('tab')


    time.sleep(0.5)
    pyautogui.write(nome) #posso colocar aspas duplas para preencher com os valores
    pyautogui.press('tab')

    time.sleep(0.5)
    pyautogui.write(empresa)
    pyautogui.press('tab')

    time.sleep(0.5)
    pyautogui.write(profissao)
    pyautogui.press('tab')

    pyautogui.press('enter')
    time.sleep(1)

    for _ in range(8):
        pyautogui.press('tab')
        time.sleep(1)
        
    pyautogui.press('enter')