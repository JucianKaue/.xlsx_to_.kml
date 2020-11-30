from time import sleep
import openpyxl
import utm


class Arvore():
    def __init__(self, numero, nome_comum, nome_cientifico, x, y):
        self.number = numero
        self.nome_comum = nome_comum
        self.nome_cientifico = nome_cientifico
        self.coordenadas = (str(x), str(y))

    def __repr__(self):
        return f"{self.number} \t\t\t\t {self.nome_comum} \t\t\t\t {self.nome_cientifico} \t\t\t\t {self.coordenadas}"


erro = True
while erro:
    erro = False
    nome_tabela = str(input('Escreva o nome do planilha (O padrão é "arvores.slsx"):\n'))

    if nome_tabela == '':
        nome_tabela = "arvores.xlsx"
        print('\033[32mO nome foi definido para "arvores.xlsx"\033[0m')
        print("=======================================")

    # Abrir tabela do excel
    try:
        wb = openpyxl.load_workbook(filename=nome_tabela, read_only=True)
    except:
        print('\033[31mO arquivo precisa estar na mesma pasta que este programa\n'
              'Se estiver na mesma pasta, verifique o nome digitado\033[0m\n'
              '=================================================')
        erro = True

table = wb["Plan1"]
# Pegar os elementos da tabela e agrupar
lista_arvores = []
for line in table:
    arvore = Arvore(
        numero=line[1].value,
        nome_comum=line[2].value,
        nome_cientifico=line[3].value,
        x=line[4].value,
        y=line[5].value
    )

    # Verificar se é coordenada
    if arvore.coordenadas[0] != None and arvore.coordenadas[1] != None:
        if arvore.coordenadas[0].isnumeric() and arvore.coordenadas[1].isnumeric():

            # Transformar coordenadas de UTM para decimal
            coordenadas_decimal = utm.to_latlon(easting=int(arvore.coordenadas[0]), northing=int(arvore.coordenadas[1]), zone_number=22, zone_letter='J')
            x = coordenadas_decimal[0]
            y = coordenadas_decimal[1]

            # Gerar o código para cada arvore
            arvore_str = f"\n<Placemark>" \
                         f"\n<name>{arvore.nome_comum}</name>" \
                         f"\n<description> " \
                         f"Número: {arvore.number}\n" \
                         f"Nome Científico: {arvore.nome_cientifico}" \
                         f"\n</description>" \
                         f"\n<Point>" \
                         f"\n<coordinates>{y}, {x}</coordinates>" \
                         f"\n</Point>" \
                         f"\n</Placemark>"
            lista_arvores.append(arvore_str)

# Gerar o código KML final
output = ['<Folder xmlns:gx="http://www.google.com/kml/ext/2.2" xmlns:atom="http://www.w3.org/2005/Atom" xmlns="http://www.opengis.net/kml/2.2">',]
for a in lista_arvores:
    output += a
output += "</Folder>"

erro_nome_saida = True
while erro_nome_saida:
    erro_nome_saida = False
    nome_saida = str(input('Qual será o nome do arquivo KML?\n'
                           'Para usar o nome padrão aperte "enter"\n'))
    if nome_saida == '':
        nome_saida = "Mapa_árvores"
        print(f'O nome foi definido como "{nome_saida}.kml"')

    # Gerar o arquivo KML
    try:
        with open(nome_saida.rstrip(".py") + ".kml", "w") as outputfile:
            outputfile.write(''.join(output))
        print("\n\033[33m====== \033[32mO programa foi finalizado com sucesso \033[33m======\033[0m")
    except:
        erro_nome_saida = True
        print("\033[31mNão é possível utilizar o nome digitado, escreva outro\033[0m")
        print("=========================================================")
    sleep(2)
