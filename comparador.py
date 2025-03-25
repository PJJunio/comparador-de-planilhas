from openpyxl import *
from openpyxl.styles import *;

def comparar (plan1, plan2, plansaida ="comparacao.xlsx"):
    wb1 = load_workbook(plan1)
    wb2 = load_workbook(plan2)

    wb_saida = Workbook()
    ws_saida = wb_saida.active
    ws_saida.title = "comparacao"

    fill_amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    ws1 = wb1.active
    ws2 = wb2.active

    max_row = max(ws1.max_row, ws2.max_row)

    for row in range (1, max_row + 1):
        valor1 = ws1.cell(row=row, column=1).value
        valor2 = ws2.cell(row=row, column=1).value

        ws_saida.cell(row=row, column=1, value=valor1)
        ws_saida.cell(row=row, column=2, value=valor2)

        if valor1 == valor2:
            ws_saida.cell(row=row, column=1).fill = fill_amarelo
            ws_saida.cell(row=row, column=2).fill = fill_amarelo

    wb_saida.save(plansaida)

    print ("Feito :)")

comparar("plan1.xlsx", "plan2.xlsx", "resultado.xlsx")