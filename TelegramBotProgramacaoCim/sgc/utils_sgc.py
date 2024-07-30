import pyautogui
import time
from datetime import datetime


def logar(chave: str, senha: str):
    pyautogui.press(keys="end")
    if len(chave) == 8:
        pyautogui.write(message=chave)
    elif len(chave) < 8:
        pyautogui.write(message=chave)
        pyautogui.press(keys="tab")
    pyautogui.press(keys="end")
    if len(senha) == 8:
        pyautogui.write(message=senha)
    elif len(senha) < 8:
        pyautogui.write(message=senha)
        pyautogui.press(keys="tab")
    pyautogui.press(keys="end")
    pyautogui.write(message="sgc")
    pyautogui.press(keys="tab")
    pyautogui.press(keys="end")
    pyautogui.write(message="p")
    pyautogui.press(keys="enter")


def definir_impressora(impressora: str):
    pyautogui.press(keys="end")
    pyautogui.write(message=impressora)
    pyautogui.press(keys="enter")
    pyautogui.write(message="s")
    pyautogui.press(keys="enter")


def definir_aplicacao(aplicacao: str):
    pyautogui.press(keys="end")
    pyautogui.write(message=aplicacao)
    pyautogui.press(keys="enter")


def tela_opcao(opcao: str):
    pyautogui.press(keys="end")
    pyautogui.write(message=opcao)
    pyautogui.press(keys="enter")
    time.sleep(0.5)


def definir_data_tela_21(data: str, emergencial: bool = True):
    if emergencial:
        data = datetime.now().strftime('%d%m%Y')
    pyautogui.press(keys="down", presses=3)
    pyautogui.press(keys="end")
    pyautogui.write(message=data)
    pyautogui.press(keys="end")
    if emergencial:
        pyautogui.write(message="s")
    else:
        pyautogui.write(message="n")
    pyautogui.press(keys="enter")
    time.sleep(0.5)
