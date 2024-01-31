




import openpyxl
from PIL import Image, ImageDraw, ImageFont

workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')

sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participante = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horario = linha[5].value
    data_emissao = linha[6].value

    fonte_nome = ImageFont.truetype('./tahomabd.ttf')
    fonte_geral = ImageFont.truetype('./tahoma.ttf')

    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)

    desenhar.text((200, 600), nome_participante, fill='black', font=fonte_nome)

    image.save(f'./{indice}{nome_participante}certificado.png')