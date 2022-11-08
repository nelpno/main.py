import math
from datetime import datetime, timedelta
import time
import pyautogui
import pyautogui as pa
import pandas as pd
from PIL import ImageGrab, ImageEnhance, ImageOps
import pytesseract

pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files/Tesseract-OCR/tesseract.exe'


def verificar_cor_eixo(r, g, b):
    achou_cor = False

    while not achou_cor:
        x, y = pyautogui.position()
        print(x, y)
        cor = pyautogui.pixel(x, y)
        print(cor)
        time.sleep(2)
        if pa.pixelMatchesColor(x, y, (r, g, b)):
            achou_cor = True


def verificar_cor_pixel(x, y, r, g, b):
    achou_cor = False

    while not achou_cor:
        if pa.pixelMatchesColor(x, y, (r, g, b)):
            achou_cor = True

    return True


def tempo(segundos):
    pyautogui.countdown(segundos)


def press_junto(x, y):
    pyautogui.hotkey(x, y)


def apertar(letra):
    pyautogui.press(letra)


def escreve(texto):
    pyautogui.write(texto)


def imagem_clicar(arquivo):
    achar_arquivo = achar_img_tela(arquivo)
    pyautogui.click(achar_arquivo)


def data_ontem():
    data_atual = datetime.now() - timedelta(1)
    data_em_texto = data_atual.strftime('%d%m%Y')
    return data_em_texto


def achar_img_tela(arquivo):
    achado = pyautogui.locateOnScreen(arquivo, confidence=0.98)
    return achado


def clicar(x, y):
    pyautogui.click(x, y)

def transforme_img_str(arquivo):
    img = pytesseract.image_to_string(arquivo, lang="eng", config=r'--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789,')
    return img

# Abrir o programa no início, programa fechado.
if achar_img_tela('icone_prosol_aberto.png') is None:
    imagem_clicar('icone_prosol.png')
    while achar_img_tela('acessar_prosol.png') is None:
        time.sleep(.25)
    imagem_clicar('acessar_prosol.png')
    pa.moveTo(80, 327)
    # abrir tela de pedidos para notas
    while achar_img_tela('menu_prosol_esquerda.png') is None:
        print(1)
        time.sleep(.25)
    imagem_clicar('movimentos_menu_esquerda.png')
    while achar_img_tela('pedido_vendas.png') is None:
        print(1)
        time.sleep(.25)
    imagem_clicar('pedido_vendas.png')
else:
    imagem_clicar('icone_prosol_aberto.png')

# abrir nova nota
while achar_img_tela('novo_pedido.png') is None:
    print(2)
    time.sleep(.25)

imagem_clicar('novo_pedido.png')

# inserir produtos e preço
i = 0

# lista de clientes puxada pelo Excel
i = 0
# p = 1 caso queira mais de uma nota por vez
df = pd.read_excel('teste.xlsx')
codigo_lista = df['Código'].tolist()
quantidade_lista = df['Quantidade'].tolist()
valor_lista = df['Valor'].tolist()
cliente_lista = df['Cliente'].tolist()
boleto_lista = df['Boleto'].tolist()

while i < len(codigo_lista) and math.isnan(codigo_lista[i]) is not True:
    escreve('{:.0f}'.format(codigo_lista[i]))
    time.sleep(0.2)
    apertar('enter')
    apertar('enter')
    escreve('{:.0f}'.format(quantidade_lista[i]))
    time.sleep(0.2)
    apertar('tab')
    escreve('{}'.format(valor_lista[i]).replace('.', ','))
    apertar('enter')
    time.sleep(0.2)
    i += 1

# achar cliente
imagem_clicar('fechar_itens_nota.png')
while achar_img_tela('lupa_cliente.png') is None:
    print(2)
    time.sleep(.25)

imagem_clicar('lupa_cliente.png')

while achar_img_tela('que_contem.png') is None:
    print(2)
    time.sleep(.25)

imagem_clicar('que_contem.png')
imagem_clicar('nome_fantasia.png')
imagem_clicar('pesquisar_por.png')
time.sleep(0.25)

escreve('{}'.format(cliente_lista[0]))
apertar('enter')
apertar('enter')
time.sleep(0.25)

# fechar pedido e emitir nota
while achar_img_tela('fecha_pedido.png') is None:
    print(2)
    time.sleep(.25)
imagem_clicar('fecha_pedido.png')

while achar_img_tela('pedido_atualizado.png') is None:
    print(2)
    time.sleep(.25)

apertar('enter')

while achar_img_tela('nota_feita.png') is not None:
    print(2)
    time.sleep(.25)

imagem_clicar('gerar_nota.png')

while achar_img_tela('logo_nfe.png') is None:
    print(2)
    time.sleep(.25)

imagem_clicar('logo_nfe.png')

# verificar se tem boleto
if math.isnan(boleto_lista[0]):
    # imprimi
    while achar_img_tela('gerar_nota_eletronica.png') is None:
        print(2)
        time.sleep(.25)
    imagem_clicar('gerar_nota_eletronica.png')
    pyautogui.doubleClick(imagem_clicar('gerar_nota_eletronica.png'))
    if achar_img_tela('abertura_caixa.png') is not None:
        apertar('tab')
        escreve('caixa')
        apertar('enter')
        apertar('enter')
        time.sleep(0.5)
        if achar_img_tela('pulou_caixa.png') is None:
            escreve('caixa')
            apertar('tab')
            imagem_clicar('selecionar_conta.png')
            apertar('down')
            imagem_clicar('abrir_caixa.png')
            while achar_img_tela('deu_certo_caixa.png') is None:
                print(2)
                time.sleep(.25)
            apertar('enter')
            apertar('enter')
else:
    # emitir nota + boleto e imprimir
    while achar_img_tela('pode_fazer_nota.png') is None:
        print(5)
        time.sleep(.25)
    imagem_clicar('plano_pagamento.png')

    while achar_img_tela('imagem_plano_pagamento.png') is None:
        print(5)
        time.sleep(.25)
    imagem_clicar('limpar_pagamento.png')
    imagem_clicar('scroll_boleto.png')
    pyautogui.click(x=1260, y=540)
    imagem_clicar('prazo_boleto.png')

    while achar_img_tela('botao_gerar_prazo_boleto.png') is None:
        print(2)
        time.sleep(.25)

    if achar_img_tela('clicar_parcelas.png') is None:
        imagem_clicar('intervalo_prazo.png')
        apertar('left')
        apertar('left')
        escreve('{:.0f}'.format(boleto_lista[0]))
        apertar('f5')
        apertar('f12')
        while achar_img_tela('salvar_prazo_pagamento.png') is None:
            print(2)
            time.sleep(.25)
        imagem_clicar('salvar_prazo_pagamento.png')
    imagem_clicar('gerar_nota_eletronica.png')
    imagem_clicar('gerar_nota_eletronica.png')
    tempo(2)
    if achar_img_tela('abertura_caixa.png') is not None:
        apertar('tab')
        escreve('caixa')
        apertar('enter')
        apertar('enter')
        time.sleep(0.5)
        if achar_img_tela('pulou_caixa.png') is None:
            escreve('caixa')
            apertar('tab')
            imagem_clicar('selecionar_conta.png')
            apertar('down')
            apertar('enter')
            while achar_img_tela('abrir_caixa.png') is None:
                print(2)
                time.sleep(.5)
            imagem_clicar('abrir_caixa.png')
            while achar_img_tela('deu_certo_caixa.png') is None:
                print(2)
                time.sleep(.25)
            apertar('enter')
            apertar('enter')

while achar_img_tela('nota_finalizada.png') is None:
    print(2)
    time.sleep(.25)

apertar('enter')

while achar_img_tela('nota_apareceu.png') is None:
    print('esperando aqui nota apareceu')
    time.sleep(.25)

# Print da nota e pega de Dados
tempo(1)
img = ImageGrab.grab((899, 381, 1653, 820))

valor_img = img.crop((694, 416, 736, 432))
valor_img_conv = valor_img.convert('1')
data_img = img.crop((21, 321, 90, 336))
numero_nfe_img = img.crop((324, 75, 368, 89))
nome_img = img.crop((11, 210, 92, 230))

nome = pytesseract.image_to_string(nome_img, lang="eng", config=r'--oem 3 --psm 7')
valor = transforme_img_str(valor_img_conv)
data = transforme_img_str(data_img)
numero_nfe = transforme_img_str(numero_nfe_img)

if math.isnan(boleto_lista[0]) is not True:
    if achar_img_tela('logo_caixa_aberto.png') is None:
        imagem_clicar('logo_caixa.png')
        imagem_clicar('logo_caixa.png')
        tempo(1)
        clicar(1136, 633)  # clique usuario
        escreve("master")
        apertar("tab")
        escreve("master")
        apertar("tab")
        press_junto("alt", "a")
        clicar(2487, 19)  # reduzir janela
        press_junto("alt", "c")
        apertar("s")
        apertar("enter")
    else:
        imagem_clicar('logo_caixa_aberto.png')

    while achar_img_tela('abriu_cobranca.png') is None:
        time.sleep(.25)
    press_junto('alt', 'f')
    apertar('alt')
    apertar('c')
    apertar('t')
    apertar('enter')
    apertar('alt')
    apertar('u')
    apertar('o')
    apertar('tab')
    escreve(data)
    pyautogui.doubleClick('duplo_clique_nfe.png')
    apertar('backspace')
    escreve('NFE {}'.format(numero_nfe))
    imagem_clicar('campo_valor_boleto.png')
    escreve('{}'.format(valor))
    imagem_clicar('clicar_sacado.png')
    press_junto('alt', 'n')
    escreve(nome)
    press_junto('alt', 'c')

# pyautogui.doubleClick(imagem_clicar('icone_impressora.png'))
#
# while achar_img_tela('impressora_ok.png') is None:
#     print(2)
#     time.sleep(.25)
#
# apertar('enter')
# verificar_cor_eixo(255,255,255)
