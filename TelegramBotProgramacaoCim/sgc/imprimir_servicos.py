import pyautogui
from sgc import verificarTela


def imprimir_servicos(codigo_servico: str, protocolos: str, equipe: str, local: str):
    linha = 0
    print("         COD      PROTOCOLO     EQPE   LOC")
    for protocolo in protocolos:
        if verificarTela.verificar_tela_58():
            linha += 1
            pyautogui.press("down")
            pyautogui.write(protocolo)
            pyautogui.press("enter")
            print(
                f"{str(linha).zfill(3)}..:   \
{codigo_servico[linha-1]}   \
{protocolo}   \
{equipe[[linha-1]]}   \
{local[[linha-1]]}"
            )
