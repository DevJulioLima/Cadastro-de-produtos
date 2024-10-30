import openpyxl
import pyperclip
import pyautogui
from time import sleep

# entrar na planilha
worbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
pagina_de_produtos = worbook['Produtos']

# copiar informações de um campo e colar no seu campo correspondente
for linha in pagina_de_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1225,367,duration=1)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1222,450,duration=1)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1212,585,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1216,676,duration=1)
    pyautogui.hotkey('ctrl','v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1262,757,duration=1)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1203,847,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1230,907,duration=1)
    sleep(3)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1206,390,duration=1)
    pyautogui.hotkey('ctrl','v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1214,478,duration=1)
    pyautogui.hotkey('ctrl','v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1204,568,duration=1)
    pyautogui.hotkey('ctrl','v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1201,656,duration=1)
    pyautogui.hotkey('ctrl','v')

    tamanho = linha[10].value
    if tamanho == 'Pequeno':
        pyautogui.click(1213,771,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(1218,801,duration=1)
    else:
        pyautogui.click(1219,828,duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1205,813,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1229,890, duration=1)
    sleep(3)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1206,428,duration=1)
    pyautogui.hotkey('ctrl','v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1217,414,duration=1)
    pyautogui.hotkey('ctrl','v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1205,605,duration=1)
    pyautogui.hotkey('ctrl','v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1212,736,duration=1)
    pyautogui.hotkey('ctrl','v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1206,822,duration=1)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1223,880, duration=1)

    pyautogui.click(1720,342, duration=1)
    sleep(3)

    pyautogui.click(1516,633,duration=1)
    sleep(3)