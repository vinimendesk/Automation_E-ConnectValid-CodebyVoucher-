# Importando bibliotecas.
import pandas as pd
import pyautogui as pygui
import pyperclip
import time
import re

# Função Ar Connect (Use na página 90%)
def automacao(codigo):
    
    time.sleep(2)

    # Clica no botão filtro.
    pygui.click(x=148, y=194, button="left")
    time.sleep(1)

    # Clica no TextField.
    pygui.click(x=514, y=439, button="left")
    time.sleep(1)

    # Seleciona todo o texto.
    pygui.hotkey("ctrl", "a")
    time.sleep(0.5)  # Espera um pouco para garantir a seleção

    # Escreve o código.
    pygui.write(str(codigo))  # Garante que o código seja uma string
    time.sleep(0.5)  # Espera um pouco após escrever

    # Clica em pesquisar.
    pygui.click(x=871, y=543, button="left")
    time.sleep(4)  # Espera a página carregar

    # Seleciona editar informação.
    pygui.click(x=1247, y=242, button="left")
    time.sleep(5)

    # Clica no campo voucher.
    pygui.click(x=106, y=332, button="left")
    # Seleciona o texto.
    pygui.hotkey("ctrl", "a")
    time.sleep(0.5)
    # Copia o texto.
    pygui.hotkey("ctrl", "c")
    time.sleep(0.5)
    
    # Valor da unidade.
    unidade = pyperclip.paste()

    # Clica em voltar.
    pygui.click(x=1314, y=148, button="left")
    time.sleep(3)
    print(unidade)

    return str(unidade)

#autoArConnect("0006242134")