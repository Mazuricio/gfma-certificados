from PIL import Image, ImageDraw, ImageFont
import openpyxl


def criar(nome, cpf, titulo, data, ciclo, horas, ordem, saveas):
    ## configs ## 
    modelo = "modelo.png"
    with Image.open(modelo) as img:
        img.load()

    d1 = ImageDraw.Draw(img)
    cor = (0,0,0)
    ## Texto
    l1 = "pela participação como ouvinte no seminário"
    l2 = f'"{titulo}"'
    l3 = f'ministrado no dia {data}, que integra o evento'
    l4 = f'{ciclo},' 
    l5 = f'com carga horária de {horas}h, na Universidade Estadual do Centro-Oeste, UNICENTRO.'

    ## escrevendo nome
    font_size = 48
    fonte_name = ImageFont.truetype('fontes/Montserrat-SemiBold.ttf', font_size)
    x = (img.size[0]/2-((len(nome)/3.57)*font_size))
    d1.text((x, 235), nome, font=fonte_name, fill =cor)
    # escrevendo CPF #
    fsize_cpf = 22
    x_cpf = (img.size[0]/2-((len(cpf)/3.7)*fsize_cpf))
    f_cpf = ImageFont.truetype('fontes/Montserrat-SemiBold.ttf', fsize_cpf)
    d1.text((x_cpf, 298), cpf, font=f_cpf, fill=cor)

    #Escrevendo texto#
    fts = 25
    font = ImageFont.truetype('fontes/Montserrat-SemiBold.ttf', fts)
    #linha 1
    xl1 = (img.size[0]/2-((len(l1)/3.7)*fts))
    d1.text((xl1, 375), l1, font=font, fill= cor)
    #l2
    xl2 = (img.size[0]/2-((len(l2)/3.7)*fts))
    font2 = ImageFont.truetype('fontes/Montserrat-Bold.ttf', fts)
    d1.text((xl2, 400), l2, font=font2, fill= cor)
    #l3
    xl3 = (img.size[0]/2-((len(l3)/3.7)*fts))
    d1.text((xl3, 425), l3, font=font, fill= cor)
    #l4
    xl4 = (img.size[0]/2-((len(l4)/3.7)*fts))
    d1.text((xl4, 450), l4, font=font2, fill= cor)
    #l5
    xl5 = (img.size[0]/2-((len(l5)/3.7)*fts))
    d1.text((xl5, 475), l5, font=font, fill= cor)
    file_name = saveas + "/" + str(ordem) + "-" + nome.split()[0] + ".png"
    img.save(file_name)


def ler(excel):
    dados = openpyxl.load_workbook(filename=excel)
    lista = []
    for d in dados['Planilha1'].iter_rows(values_only = True):
        lista.append(d)
    lista.remove(lista[0])
    return lista

def gerar(lista, titulo, data, ciclo, horas, saveas):
    conta = 0
    for i in lista:
        criar(i[0], i[1], titulo, data, ciclo, horas, conta, saveas)
        conta += 1

