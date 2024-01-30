




import openpyxl

workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')

sheet_alunos = workbook_alunos['Sheet1']

for linha in sheet_alunos.iter_rows(min_row=2):
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participante = linha[2].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    carga_horario = linha[5].value
    data_emissao = linha[6].value
