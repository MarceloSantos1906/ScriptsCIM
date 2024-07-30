import pyautogui
import pyperclip
import time
import keyboard

def select_text(presses):
    pyperclip.copy('')
    pyautogui.press('numlock')
    with pyautogui.hold('shift'):
        pyautogui.press('right', presses = presses)
    keyboard.press_and_release(hotkey='ctrl+c')
    pyautogui.press('numlock')
    return pyperclip.paste().strip()

def copy_screen():
    pyperclip.copy('')
    pyautogui.press('numlock')
    keyboard.press_and_release(hotkey='ctrl+a')
    keyboard.press_and_release(hotkey='ctrl+c')
    pyautogui.press('numlock')
    time.sleep(0.05)
    return pyperclip.paste()

def verificar_Carregando():
    imagemCarregando = '.\\imgs\\loading_extra.png'
    imagemCoords = None
    try:
        imagemCoords = pyautogui.locateOnScreen(imagemCarregando, grayscale = True, confidence = 0.7)
    except pyautogui.ImageNotFoundException:
        pass
    if imagemCoords is None:
        return
    else:
        verificar_Carregando()

def verificar_pos_invalida():
    imagemCarregando = '.\\imgs\\posisao_invalida.png'
    imagemCoords = None
    try:
        imagemCoords = pyautogui.locateOnScreen(imagemCarregando, grayscale = True, confidence = 0.7)
    except pyautogui.ImageNotFoundException:
        pass
    if imagemCoords is None:
        return
    else:
        pyautogui.press('esc')
        pyautogui.press('tab')
        verificar_pos_invalida()

def verificar_telas():
    while True:
        pyperclip.copy('')
        clipboard = copy_screen()
        if 'UNSUPORTED FUNCTION' in clipboard or \
            'UNABLE TO ESTABLISH SESSION' in clipboard:
            return 'disconected', clipboard
        elif 'Entre com a sigla da aplicacao:' in clipboard:
            return 'aplicacao', clipboard
        elif 'Companhia de Saneamento do Parana - S A N E P A R' in clipboard:
            return 'login', clipboard
        elif 'INFORME A IMPRESSORA..' in clipboard:
            return 'impressora', clipboard
        elif '32 AS ELETRONIC' in clipboard:
            return 'opcao', clipboard
        elif 'DT PARA VERIFICACAO:' in clipboard:
            return 'data_programacao', clipboard
        elif 'PROGRAMACAO/PRORROGACAO DE SERVICOS' in clipboard:
            return 'tela_21', clipboard
        elif 'PROTOCOLOS EM CAMPO:' in clipboard:
            return 'tela_32', clipboard
        elif 'CARGA AS ELETRONICA' in clipboard:
            return 'tela_32_sel_eq', clipboard
        elif 'SGCPAS30' in clipboard:
            return 'tela_32_lista', clipboard
        elif 'EMISSAO DE A. S.' in clipboard:
            return 'tela_58', clipboard
        elif 'NAO HA SERVICOS' in clipboard:
            return 'nao_ha_servicos', clipboard

def verificar_tela_58():
    loop = 0
    while True:
        pyperclip.copy('')
        clipboard = copy_screen()
        if 'EMISSAO DE A. S.' in clipboard:
            return True
        elif loop == 3:
            pyautogui.press('f3')
        else:
            loop += 1


if __name__ == '__main__':
    time.sleep(2)
    verificar_telas()
