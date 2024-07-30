from sgc import verificarTela, utils_sgc
import pyautogui
import keyboard
import pyperclip

def programar(protocolo:str, equipe:str, index:int, emergencial:bool = True):
    print('a')
    tela = verificarTela.verificar_telas()
    if  tela == 'opcao':
        utils_sgc.tela_opcao(opcao="21")
    elif tela == "data_programacao":
        utils_sgc.definir_data_tela_21(data="11112222", emergencial=emergencial)

    while True:
        print('b')
        if index > 14:
            pyautogui.press('enter')
            index -= 14
            continue
        pyautogui.press('numlock')
        pyautogui.press('down', presses=index)
        pyautogui.press('left', presses=18)
        with pyautogui.hold('shift'):
            pyautogui.press('right', presses=17)
        keyboard.press_and_release(hotkey='ctrl+c')
        if pyperclip.paste() == protocolo:
            pyautogui.press('tab', presses=2)
            pyautogui.write(equipe)
            pyautogui.press('enter')
            pyautogui.press('numlock')
            break
        pyautogui.press('numlock')
        break
    return
