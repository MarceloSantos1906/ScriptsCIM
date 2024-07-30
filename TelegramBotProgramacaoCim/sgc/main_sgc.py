import subprocess
import pyautogui
import keyboard
import pyperclip
import time
from sgc import (
    verificarTela,
    reconectar,
    utils_sgc,
    ativar_janela,
    verificar_janela_existe,
    tela21
)


def verificar_tela_21_emergencial():
    janela = "extra"
    dados = []
    index = []
    ativar_janela.ativarJanela(titulo=janela)
    time.sleep(0.5)
    verificarTela.verificar_pos_invalida()
    verificarTela.verificar_Carregando()

    pyautogui.press("f3")
    while True:
        tela, clipboard = verificarTela.verificar_telas()
        if tela == "opcao":
            verificarTela.verificar_pos_invalida()
            verificarTela.verificar_Carregando()
            pyautogui.press("f3")
            utils_sgc.tela_opcao(opcao="21")
        elif tela == "data_programacao":
            verificarTela.verificar_pos_invalida()
            verificarTela.verificar_Carregando()
            utils_sgc.definir_data_tela_21(data="11112222", emergencial=True)
        elif tela == "tela_21":
            verificarTela.verificar_pos_invalida()
            verificarTela.verificar_Carregando()
            loop = 0
            while True:
                clipboard = verificarTela.copy_screen()
                dado, indexes = tela21.separar_dados_emergencial(clipboard)
                pyautogui.press("enter")
                verificarTela.verificar_pos_invalida()
                verificarTela.verificar_Carregando()
                for i in range(len(dado)):
                    dados.append(dado[i])
                for i in range(len(indexes)):
                    index.append(indexes[i])
                clipboard_anterior = clipboard
                tela, clipboard = verificarTela.verificar_telas()
                verificarTela.verificar_pos_invalida()
                verificarTela.verificar_Carregando()
                time.sleep(0.5)

                if tela == "data_programacao":
                    pyautogui.press("f3")
                    break
                elif clipboard_anterior == clipboard:
                    loop += 1
                elif loop == 4:
                    pyautogui.press("f3")
                    break
            return dados, index
        elif tela == 'nao_ha_servicos':
            pyautogui.press(keys='f3')
            return dados, index


def programar_servico_emergencial(
        codigo_serv: str,
        protocolo: str,
        equipe: str,
):
    ordem = 0
    janela = "extra"
    loop = 0
    protocolo = protocolo.strip()
    ativar_janela.ativarJanela(titulo=janela)
    pyautogui.press("f3")
    clipboard_anterior = ""
    while True:
        tela, clipboard = verificarTela.verificar_telas()
        if tela == "opcao":
            loop = 0
            verificarTela.verificar_pos_invalida()
            verificarTela.verificar_Carregando()
            utils_sgc.tela_opcao(opcao="21")
        elif tela == "data_programacao":
            loop = 0
            verificarTela.verificar_pos_invalida()
            verificarTela.verificar_Carregando()
            utils_sgc.definir_data_tela_21(data="11112222", emergencial=True)
        elif tela == "tela_21":
            loop = 0
            while True:
                tela, clipboard = verificarTela.verificar_telas()
                if tela == "data_programacao":
                    return False
                clipboard = verificarTela.copy_screen()
                time.sleep(0.5)
                cod_sev, protocolos, index = (
                    tela21.buscar_emergencial(codigo_serv, protocolo, clipboard)
                )
                if cod_sev == '' and protocolos == '' and index == '':
                    pyautogui.press('enter')
                elif cod_sev == codigo_serv and protocolos == protocolo:
                    pyautogui.press(keys='down', presses=index)
                    pyautogui.press(keys='tab')
                    pyautogui.write(equipe)
                    pyautogui.press(keys='enter')
                    pyautogui.press(keys='f3')
                    return True
        elif clipboard_anterior == clipboard:
            loop += 1
            time.sleep(1)
        elif loop == 3:
            loop = 0
            pyautogui.press('esc')
            pyautogui.press("f3")
        else:
            clipboard_anterior = clipboard


def enviar_serv_32(
        codigo_serv: str,
        protocolo: str,
        equipe: str,
):
    ordem = 0
    janela = "extra"
    loop = 0
    protocolo = protocolo.strip()
    ativar_janela.ativarJanela(titulo=janela)
    pyautogui.press("f3")
    clipboard_anterior = ""
    while True:
        tela, clipboard = verificarTela.verificar_telas()
        if tela == "opcao":
            loop = 0
            verificarTela.verificar_pos_invalida()
            verificarTela.verificar_Carregando()
            utils_sgc.tela_opcao(opcao="32")
        elif tela == "tela_32":
            loop = 0
            verificarTela.verificar_pos_invalida()
            verificarTela.verificar_Carregando()
            pyautogui.press(keys='x')
            pyautogui.press(keys='enter')
            time.sleep(0.5)
        elif tela == "tela_32_sel_eq":
            loop = 0
            verificarTela.verificar_pos_invalida()
            verificarTela.verificar_Carregando()
            pyautogui.write(equipe)
            pyautogui.press(keys='enter')
            time.sleep(0.5)
        elif tela == "tela_32_lista":
            loop = 0
            while True:
                time.sleep(0.5)
                tela, clipboard = verificarTela.verificar_telas()
                clipboard = verificarTela.copy_screen()
                time.sleep(0.5)
                index = (
                    tela21.buscar_emergencial_32(codigo_serv, protocolo, clipboard)
                )
                if index == 'continua':
                    time.sleep(0.5)
                    pyautogui.press('enter')
                elif type(index) is int:
                    pyautogui.press(keys='down', presses=index)
                    pyautogui.write('01')
                    pyautogui.press(keys='f7')
                    pyautogui.press(keys='s')
                    pyautogui.press(keys='enter')
                    pyautogui.press(keys='f3')
                    pyautogui.press(keys='f1')
                    return True
                elif not index:
                    return False
        elif clipboard_anterior == clipboard:
            loop += 1
            time.sleep(1)
        elif loop == 3:
            loop = 0
            pyautogui.press('esc')
            pyautogui.press("f3")
        else:
            clipboard_anterior = clipboard


def programar_servicos(equipes: list, base: str):
    janela = "extra"
    clipboard_anterior = ""
    codigos_para_programar = [
        "0040",
        "0042",
        "0047",
        "0115",
        "0130",
        "0131",
        "0132",
        "0350",
        "0541",
        "0551",
        "0552",
        "0705",
        "0710",
        "0715",
        "0720",
        "0730",
        "0745",
        "0750",
        "0751",
        "0930",
        "1050",
        "1060",
        "1065",
        "1100",
        "1110",
        "1390",
        "1400",
        "1401",
        "1410",
        "1470",
        "1480",
        "1490",
        "1510",
        "1520",
        "1540",
        "2000",
        "2010",
        "2051",
        "2060",
        "2090",
        "2160",
        "2170",
        "2180",
        "2445",
        "1555",
        "1556",
        "1557",
        "1560",
        "3020",
        "3080",
        "3090",
        "3130",
        "3245",
        "3280",
        "3290",
        "3310",
        "3390",
        "3400",
        "3415",
        "3420",
        "3440",
        "3450",
        "3460",
        "3560",
        "3585",
        "3830",
    ]

    ativar_janela.ativarJanela(titulo=janela)
    time.sleep(0.5)
    pyautogui.press("f3")
    verificarTela.verificar_pos_invalida()
    verificarTela.verificar_Carregando()

    lista_dados = []
    while True:
        tela, clipboard = verificarTela.verificar_telas()
        if tela == "opcao":
            verificarTela.verificar_pos_invalida()
            verificarTela.verificar_Carregando()
            pyautogui.press("f3")
            utils_sgc.tela_opcao(opcao="21")
        elif tela == "data_programacao":
            verificarTela.verificar_pos_invalida()
            verificarTela.verificar_Carregando()
            utils_sgc.definir_data_tela_21(data="11112222", emergencial=False)
        elif tela == "tela_21":
            loop = 0
            while True:
                tela, clipboard = verificarTela.verificar_telas()
                if tela == "data_programacao":
                    lista_dados, excecoes = verificar_21_X(equipes, base, lista_dados)
                    return lista_dados, excecoes
                clipboard = verificarTela.copy_screen()
                time.sleep(0.5)
                cod_sevs, protocolos, data_vencimento, equipe_programada, locais, indexes, excecoes = (
                    tela21.separar_dados(clipboard, base=base)
                )
                if clipboard_anterior == clipboard:
                    loop += 1
                    time.sleep(1)
                if len(indexes) == 0:
                    pyautogui.press("enter")
                    clipboard_anterior = clipboard
                    continue
                for i in range(len(indexes)):
                    codigo_servico = cod_sevs[i]
                    protocolo = protocolos[i]
                    vencimento = data_vencimento[i]
                    equipe = equipe_programada[i]
                    local = locais[i]
                    pyautogui.press("down", presses=indexes[i])
                    pyautogui.press("tab")
                    ##
                    programar_equipe(
                        equipes=equipes,
                        codigo_servico=codigo_servico,
                        equipe_programada=equipe,
                        local=local,
                    )
                    ##
                    pyautogui.press("tab", presses=2)
                    pyautogui.press("up", presses=indexes[i] + 1)
                    lista_dados.append(codigo_servico)
                    lista_dados.append(protocolo)
                    lista_dados.append(vencimento)
                    lista_dados.append(equipe)
                    lista_dados.append(local)

                pyautogui.press("enter")
                clipboard_anterior = clipboard


def verificar_21_X(equipes, base, lista_dados):
    lista_dados_ = []
    for i in range(len(lista_dados)):
        lista_dados_.append(lista_dados[i])
    codigos_para_programar = [
        "0040",
        "0042",
        "0047",
        "0115",
        "0130",
        "0131",
        "0132",
        "0350",
        "0541",
        "0551",
        "0552",
        "0705",
        "0710",
        "0715",
        "0720",
        "0730",
        "0745",
        "0750",
        "0751",
        "0930",
        "1050",
        "1060",
        "1065",
        "1100",
        "1110",
        "1390",
        "1400",
        "1401",
        "1410",
        "1480",
        "1490",
        "1510",
        "1520",
        "1540",
        "2000",
        "2010",
        "2051",
        "2060",
        "2090",
        "2160",
        "2170",
        "2180",
        "2445",
        "1555",
        "1556",
        "1557",
        "1560",
        "2220",
        "3020",
        "3080",
        "3090",
        "3113",
        "3245",
        "3280",
        "3290",
        "3310",
        "3390",
        "3400",
        "3415",
        "3420",
        "3440",
        "3450",
        "3460",
        "3560",
        "3585",
        "3830",
    ]
    pyautogui.press(keys="down", presses=5)
    pyautogui.press(keys="x")
    pyautogui.press(keys="enter")
    time.sleep(1)
    verificarTela.verificar_pos_invalida()
    verificarTela.verificar_Carregando()
    tela, clipboard = verificarTela.verificar_telas()
    clipboard = verificarTela.copy_screen()
    cod_sevs, protocolos, data_vencimento, equipe_programada, locais, indexes, excecoes = (
        tela21.separar_dados(clipboard, base=base)
    )
    if len(indexes) == 0:
        pyautogui.press("enter")
        return lista_dados, excecoes
    for i in range(len(indexes)):
        codigo_servico = cod_sevs[i]
        protocolo = protocolos[i]
        vencimento = data_vencimento[i]
        equipe = equipe_programada[i]
        local = locais[i]
        pyautogui.press("down", presses=indexes[i])
        pyautogui.press("tab")
        ##
        programar_equipe(
            equipes=equipes,
            codigo_servico=codigo_servico,
            equipe_programada=equipe,
            local=local,
        )
        ##
        pyautogui.press("tab", presses=2)
        pyautogui.press("up", presses=indexes[i] + 1)
        lista_dados.append(codigo_servico)
        lista_dados.append(protocolo)
        lista_dados.append(vencimento)
        lista_dados.append(equipe)
        lista_dados.append(local)
    return lista_dados, excecoes


def programar_equipe(
        equipes: str, codigo_servico: str, equipe_programada: str, local: str
):
    """
    Equipes:
    0 - equipe provisória
    1 - hidrometros
    2 - caixa padão subterranea
    3 - caixa padrão de muro
    4 - ligações
    5 - manutenção
    6 - calçadas
    7 - asfalto
    8 - esgoto
    """
    codigos_hidrometros = [
        "0350",
        "0365",
        "0370",
        "0400",
        "0541",
        "0551",
        "0552",
        "0555",
        "0556",
        "0565",
        "0601",
        "0602",
    ]
    codigos_cps = [
        "0042",
        "0047",
        "0115",
        "0130",
        "0131",
        "1401",
    ]
    codigo_cpm = ["0040", "0132"]
    codigos_calcada = ["1555", "1556", "1557"]
    codigo_asfalto = "1560"
    codigos_esgoto = [
        "3020",
        "3030",
        "3040",
        "3080",
        "3090",
        "3113",
        "3130",
        "3140",
        "3242",
        "3243",
        "3245",
        "3280",
        "3290",
        "3300",
        "3310",
        "3390",
        "3400",
        "3415",
        "3420",
        "3440",
        "3450",
        "3460",
        "3545",
        "3560",
        "3562",
        "3570",
        "3580",
        "3585",
        "3610",
        "3620",
        "3670",
        "3700",
        "3830",
    ]
    codigos_ligacao_autorizacao = ["0750", "0751"]
    locais_n_precisa_autorizacao = ["055", "088", "104", "413", "674"]
    codigos_manutencao = [
        "0705",
        "0710",
        "0715",
        "0720",
        "0730",
        "0745",
        "0930",
        "1050",
        "1060",
        "1065",
        "1100",
        "1110",
        "1310",
        "1390",
        "1400",
        "1410",
        "1470",
        "1480",
        "1490",
        "1510",
        "1520",
        "1540",
        "1565",
        "2000",
        "2010",
        "2051",
        "2060",
        "2090",
        "2160",
        "2170",
        "2180",
        "2220",
        "2445",
    ]
    codioReaterro  = '2180'
    if codigo_servico in codigos_hidrometros:
        pyperclip.copy(equipes[1])
        keyboard.press_and_release(hotkey="ctrl+v")
    elif codigo_servico in codigos_cps:
        pyperclip.copy(equipes[2])
        keyboard.press_and_release(hotkey="ctrl+v")
    elif codigo_servico in codigo_cpm:
        pyperclip.copy(equipes[3])
        keyboard.press_and_release(hotkey="ctrl+v")
    elif (
            codigo_servico in codigos_ligacao_autorizacao
            and equipe_programada == equipes[0]
            and local not in locais_n_precisa_autorizacao
    ):
        pyperclip.copy(equipes[4])
        keyboard.press_and_release(hotkey="ctrl+v")
    elif (
            codigo_servico in codigos_ligacao_autorizacao
            and local in locais_n_precisa_autorizacao
    ):
        pyperclip.copy(equipes[4])
        keyboard.press_and_release(hotkey="ctrl+v")
    elif codigo_servico in codigos_manutencao:
        pyperclip.copy(equipes[5])
        keyboard.press_and_release(hotkey="ctrl+v")
    elif codigo_servico in codigos_calcada:
        pyperclip.copy(equipes[6])
        keyboard.press_and_release(hotkey="ctrl+v")
    elif codigo_servico == codigo_asfalto:
        pyperclip.copy(equipes[7])
        keyboard.press_and_release(hotkey="ctrl+v")
    elif codigo_servico in codigos_esgoto:
        pyperclip.copy(equipes[8])
        keyboard.press_and_release(hotkey="ctrl+v")
    elif codigo_servico in codioReaterro:
        pyperclip.copy(equipes[9])
        keyboard.press_and_release(hotkey="ctrl+v")


def imprimir_servico(protocolo: str, impressora: str):
    janela = "extra"
    protocolo = protocolo.strip()
    ativar_janela.ativarJanela(titulo=janela)
    pyautogui.press("f3")
    while True:
        verificarTela.verificar_pos_invalida()
        verificarTela.verificar_Carregando()

        tela, clipboard = verificarTela.verificar_telas()
        if tela == "opcao":
            verificarTela.verificar_pos_invalida()
            verificarTela.verificar_Carregando()
            utils_sgc.tela_opcao(opcao="58")
        elif tela == "tela_58":
            pyautogui.press("down")
            pyautogui.write(protocolo)
            if impressora != "pdpgo125":
                pyautogui.press("down", presses=3)
                pyautogui.write(impressora)
            pyautogui.press("enter")
            pyautogui.press("f3")
            return
        else:
            pyautogui.press("f3")


def sgc(chave: str, senha: str, impressora: str):
    janela = "extra"
    caminho = "extra.edp"
    if not (verificar_janela_existe.main(classname="SDIMainFrame")):
        subprocess.Popen(
            caminho,
            shell=True,
        )
        while True:
            if not (verificar_janela_existe.main(classname="SDIMainFrame")):
                time.sleep(1)
            else:
                break
    ativar_janela.ativarJanela(titulo=janela)
    while True:
        verificarTela.verificar_pos_invalida()
        verificarTela.verificar_Carregando()

        tela, clipboard = verificarTela.verificar_telas()
        if tela == "disconected":
            reconectar.reconectar()
        elif tela == "aplicacao":
            verificarTela.verificar_pos_invalida()
            verificarTela.verificar_Carregando()
            utils_sgc.definir_aplicacao(aplicacao="sgc")
        elif tela == "login":
            verificarTela.verificar_pos_invalida()
            verificarTela.verificar_Carregando()
            utils_sgc.logar(chave=chave, senha=senha)
        elif tela == "impressora":
            verificarTela.verificar_pos_invalida()
            verificarTela.verificar_Carregando()
            utils_sgc.definir_impressora(impressora=impressora)
        elif tela == "opcao":
            return
        elif tela == "data_programacao":
            return
        elif tela == "tela_21":
            return


if __name__ == "__main__":
    exit(1)
