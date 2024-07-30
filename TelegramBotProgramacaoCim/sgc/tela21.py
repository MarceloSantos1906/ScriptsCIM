import pyautogui
from sgc import verificarTela
import datetime


def separar_dados_emergencial(clipboard):
    codigos_emergenciais = [
        "0010",
        "0050",
        "0070",
        "0090",
        "0110",
        "0151",
        "0245",
        "0365",
        "0370",
        "1271",
        "1280",
        "1330",
        "1350",
        "1360",
        "1380",
        "1430",
        "1500",
        "1740",
        "1750",
        "1760",
        "1770",
        "1771",
        "1780",
        "1800",
        "1830",
        "1840",
        "1850",
        "1870",
        "1970",
        "2050",
        "2350",
        "2455",
        "2465",
        "3410",
        "3455",
        "3465",
        "3466",
        "3470",
        "3600",
        "3690",
        "3695",
        "3701",
    ]

    linhas = clipboard.splitlines()
    for linha in range(len(linhas)):
        while True:
            if len(linhas[linha]) < 80:
                linhas[linha] += " "
            else:
                break
    informacoes = []
    indexes = []
    for linha in range(8, 22):
        if linhas[linha][2:6] in codigos_emergenciais:
            verificarTela.verificar_Carregando()
            indexes.append(linha - 8)
            codigo_servico = linhas[linha][2:6]
            protocolo = linhas[linha][12:29]
            equipe = linhas[linha][39:43]
            local = linhas[linha][49:52]
            informacoes.append(codigo_servico)
            informacoes.append(protocolo)
            informacoes.append(equipe)
            informacoes.append(local)
            pyautogui.press("left", presses=2)
            pyautogui.press("down", presses=linha - 8)
            pyautogui.press("f5")
            clipboard = verificarTela.copy_screen()
            pyautogui.press("enter")
            linhas_inf = clipboard.splitlines()
            endereco = linhas_inf[11][17:53]
            motivo = linhas_inf[18][17:53]
            informacoes.append(endereco)
            informacoes.append(motivo)
    return informacoes, indexes

def buscar_emergencial(cod_serv, protocolo, clipboard):
    linhas = clipboard.splitlines()
    for linha in range(len(linhas)):
        while True:
            if len(linhas[linha]) < 80:
                linhas[linha] += " "
            else:
                break
    for linha in range(8, 22):
        if cod_serv in linhas[linha][2:6] and protocolo in linhas[linha][12:29]:
            codigo_servico = linhas[linha][2:6]
            protocolo = linhas[linha][12:29]
            local = linhas[linha][49:52]
            indexes = linha - 8
            return codigo_servico, protocolo, indexes
    return '', '', ''

def buscar_emergencial_32(cod_serv, protocolo, clipboard):
    linhas = clipboard.splitlines()
    for linha in range(len(linhas)):
        while True:
            if len(linhas[linha]) < 80:
                linhas[linha] += " "
            else:
                break
    for linha in range(6, 22):
        if cod_serv in linhas[linha][5:9] and protocolo in linhas[linha][10:27]:
            print(linhas[linha])
            index = linha - 6
            return index
        elif 'CONTINUA..' in linhas[linha][67:77]:
            return 'continua'
    return False


def separar_dados(clipboard, base):
    cidades = []
    locais_n_precisa_autorizacao = []
    if base == "uva":
        cidades = [
            "088",
            "104",
            "196",
            "211",
            "283",
            "400",
            "644",
            "645",
            "674",
            "682",
            "668",
        ]
        locais_n_precisa_autorizacao = ["055", "088", "104", "413", "674"]
    elif base == "bituruna":
        cidades = ["055", "413"]

    agora = datetime.datetime.now()
    if agora.isoweekday() == 1:
        ontem = agora - datetime.timedelta(days=3)
    else:
        ontem = agora - datetime.timedelta(days=1)
    ontem_cinco_hrs = ontem.strftime("%Y%m%d") + "1700"
    ontem_cinco_hrs = datetime.datetime.strptime(ontem_cinco_hrs, "%Y%m%d%H%M")
    codigos_para_programar = [
        "0040",
        "0042",
        "0047",
        "0115",
        "0130",
        "0131",
        "0132",
        "0185",
        "0186",
        "0205",
        "0280",
        "0300",
        "0350",
        "0400",
        "0541",
        "0551",
        "0552",
        "0555",
        "0556",
        "0565",
        "0601",
        "0602",
        "0700",
        "0705",
        "0710",
        "0715",
        "0720",
        "0725",
        "0730",
        "0740",
        "0745",
        "0750",
        "0751",
        "0752",
        "0813",
        "0870",
        "0880",
        "0930",
        "0990",
        "1050",
        "1060",
        "1065",
        "1070",
        "1075",
        "1090",
        "1100",
        "1105",
        "1110",
        "1260",
        "1270",
        "1310",
        "1390",
        "1400",
        "1401",
        "1410",
        "1415",
        "1470",
        "1480",
        "1490",
        "1493",
        "1496",
        "1510",
        "1520",
        "1540",
        "1555",
        "1556",
        "1557",
        "1560",
        "1565",
        "1745",
        "2000",
        "2010",
        "2040",
        "2051",
        "2060",
        "2070",
        "2090",
        "2110",
        "2150",
        "2160",
        "2170",
        "2180",
        "2185",
        "2186",
        "2220",
        "2230",
        "2445",
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

    linhas = clipboard.splitlines()
    for linha in range(len(linhas)):
        while True:
            if len(linhas[linha]) < 80:
                linhas[linha] += " "
            else:
                break
    cod_sevs = []
    protocolos = []
    data_vencimento = []
    equipes = []
    locais = []
    indexes = []
    excecoes = []
    print(f"protocolos entre as datas: {ontem_cinco_hrs} e {agora}")
    for linha in range(8, 22):
        try:
            codigo_servico = linhas[linha][2:6]
            protocolo = linhas[linha][12:29]
            data_protocolo = linhas[linha][12:24]
            vencimento = linhas[linha][30:38]
            equipe = linhas[linha][39:43]
            local = linhas[linha][49:52]
            if (
                    codigo_servico == "    "
                    or (equipe != "    " and equipe != "1999")
                    or local not in cidades
            ):
                continue
            try:
                data_protocolo = datetime.datetime.strptime(data_protocolo, "%Y%m%d%H%M")
            except Exception as e:
                excecoes.append(e)
                excecoes.append(f"|{linhas[linha]}|")
                continue

            if (
                    codigo_servico in codigos_ligacao_autorizacao
                    and equipe == "1999"
                    and local not in locais_n_precisa_autorizacao
            ):
                print(
                    f"{codigo_servico}  {protocolo}  {vencimento}  {equipe}  {local}"
                )
                indexes.append(linha - 8)
                cod_sevs.append(codigo_servico)
                protocolos.append(protocolo)
                data_vencimento.append(vencimento)
                equipes.append(equipe)
                locais.append(local)
            elif (codigo_servico in codigos_ligacao_autorizacao
                and equipe == "    "
                and local in locais_n_precisa_autorizacao):
                print(
                    f"{codigo_servico}  {protocolo}  {vencimento}  {equipe}  {local}"
                )
                indexes.append(linha - 8)
                cod_sevs.append(codigo_servico)
                protocolos.append(protocolo)
                data_vencimento.append(vencimento)
                equipes.append(equipe)
                locais.append(local)
            elif (codigo_servico in codigos_para_programar and (
                    ontem_cinco_hrs < data_protocolo < agora
            ) and codigo_servico not in codigos_ligacao_autorizacao):
                print(
                    f"{codigo_servico}  {protocolo}  {vencimento}  {equipe}  {local}"
                )
                indexes.append(linha - 8)
                cod_sevs.append(codigo_servico)
                protocolos.append(protocolo)
                data_vencimento.append(vencimento)
                equipes.append(equipe)
                locais.append(local)
        except Exception as e:
            print(f'Exception: {e}')
            continue
    return cod_sevs, protocolos, data_vencimento, equipes, locais, indexes, excecoes
