import pyautogui
import time

def reconectar():
    pyautogui.press('alt')
    pyautogui.press('right', presses=4)
    pyautogui.press('down', presses=2)
    pyautogui.press('enter', presses=2)
    time.sleep(0.05)
    pyautogui.press('alt')
    pyautogui.press('right', presses=4)
    pyautogui.press('down')
    pyautogui.press('enter')
    time.sleep(0.5)