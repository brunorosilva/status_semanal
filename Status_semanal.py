import xlsxwriter
import pandas as pd
import numpy as np
import excel2img

df = pd.read_excel('D:\\Users\\50047924\\Documents\\Status Semanal v2.xlsx', 'Base', encoding = 'latin-1')

workbook = xlsxwriter.Workbook('Teste.xlsx')
worksheet = workbook.add_worksheet()
worksheet.hide_gridlines(2)
# Some data we want to write to the worksheet.
df = df[df['Dia inicial'] == max(df['Dia inicial'])]

colunas = df.columns

projetos = df['Projeto'].unique()

titulo = workbook.add_format()
titulo.set_font_size(26)
titulo.set_align('center')
titulo.set_align('vcenter')
titulo.set_bottom()
titulo.set_right()
titulo.set_left()
titulo.set_top()

formato_coluna = workbook.add_format()
formato_coluna.set_align('center')
formato_coluna.set_align('vcenter')
formato_coluna.set_bold()
formato_coluna.set_bottom()
formato_coluna.set_right()
formato_coluna.set_left()
formato_coluna.set_top()

formato_celula = workbook.add_format()
formato_celula.set_align('center')
formato_celula.set_align('vcenter')
formato_celula.set_bottom()
formato_celula.set_right()
formato_celula.set_left()
formato_celula.set_top()

formato_acao = workbook.add_format()
formato_acao.set_bottom()
formato_acao.set_right()
formato_acao.set_left()
formato_acao.set_top()


row = 1
worksheet.set_row(0, 4)
col = 0
count = 0
for projeto in projetos:
    i = 0
    df_aux = df[df['Projeto'] == projeto].reset_index()
    acoes = df_aux['Ações'].unique()
    worksheet.merge_range(row, col, row, col + 5, projeto, titulo)
    for coluna in colunas[3:]:
        worksheet.set_column(0, 0, 20)
        worksheet.set_column(1, 1, 35)
        worksheet.set_column(2, 6, 16)
        worksheet.write(row+1, col, coluna, formato_coluna)
        col += 1
        if col == len(colunas)-3:
            col = 0
    
    worksheet.set_row(row + 1, None, None, {'level':1})

    worksheet.merge_range(row+2, col, row + 1 + len(acoes), col, df_aux['Semana de atuação'].values[0])

    for acao in acoes:
        formato_check = workbook.add_format()
        formato_check.set_align('center')
        formato_check.set_align('vcenter')
        formato_check.set_bottom()
        formato_check.set_right()
        formato_check.set_left()
        formato_check.set_top()
        event = u""+str(df_aux['Status'].values[i])
        comp = u'✅'
        if event == comp:
            formato_check.set_font_color('#008000')
            print('bateu')
        else:
            formato_check.set_font_color('#FF0000')
        worksheet.write(row + 2, col, df_aux['Semana de atuação'].values[i], formato_celula)
        worksheet.write(row + 2, col + 1, acao, formato_acao)
        worksheet.write(row + 2, col + 2, df_aux['Responsável'].values[i], formato_celula)
        worksheet.write(row + 2, col + 3, str(df_aux['Previsão'].values[i]) if (df_aux['Previsão'].values[i] == np.nan) else "", formato_celula)
        worksheet.write(row + 2, col + 4, df_aux['Status'].values[i], formato_check)
        worksheet.write(row + 2, col + 5, str(df_aux['Comentário'].values[i]) if (df_aux['Comentário'].values[i] == np.nan) else "", formato_celula)
        worksheet.set_row(row + 2, None, None, {'level':1})
        #worksheet.set_row(row + 2, None, None, {'level':2})
        row += 1
        i += 1
        if i == len(df_aux):
            i = 0
        
    #print('inicio', row - count)
    #print('final', row + len(acoes) - 1 - count)
    pulo = 2
    row += pulo
    count += 1


workbook.close()


excel2img.export_img("Teste.xlsx", "test.png", "Sheet1", None)

