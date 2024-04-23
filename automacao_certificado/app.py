# ler dados da planilha
# tranasferir para imagen do certificado



import openpyxl
from PIL import Image, ImageDraw, ImageFont 

# abrindo a planilha 
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

# para cada linha na planilha ler a partir da linha 2
for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    # cada celula que contém informações que precisamos
    nome_curso = linha[0].value # nome do curso
    nome_participante = linha[1].value # nome do participante
    tipo_participacao = linha[2].value # tipo de participação
    data_inicio = linha[3].value # data de inicio
    data_final = linha[4].value # data final 
    carga_horaria = linha[5].value # carga horaria do curso 
    data_emissao = linha[6].value # data de emissão do certificado

# tranaferindo os dados da planilha para o certificado
# definindo fonte
    fonte_nome = ImageFont.truetype('./tahomabd.ttf',90)
    fonte_geral = ImageFont.truetype('./tahoma.ttf',80)
    fonte_data = ImageFont.truetype('./tahoma.ttf',50)

    image = Image.open('./certificado_padrao.jpg')  
    desenhar = ImageDraw.Draw(image)

    desenhar.text((1000,829), nome_participante,fill='black',font=fonte_nome)
    desenhar.text((1060,959), nome_curso,fill='black',font=fonte_geral)
    desenhar.text((1425,1070), tipo_participacao,fill='black',font=fonte_geral)
    desenhar.text((755,1780), data_inicio,fill='black',font=fonte_data)
    desenhar.text((755,1935), data_final,fill='black',font=fonte_data)
    desenhar.text((1479,1190), str(carga_horaria),fill='black',font=fonte_geral)
    desenhar.text((2235,1935), data_emissao,fill='black',font=fonte_data)

# salvando certificados
    image.save(f'./{indice} {nome_participante} certificado.png')