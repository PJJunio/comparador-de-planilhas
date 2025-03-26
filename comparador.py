import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from openpyxl import *
from openpyxl.styles import *;
from openpyxl.utils import get_column_letter

def comparar():

    root = tk.Tk()
    root.withdraw()

    file_types = ([("Arquivos Excel (.xlsx)", "*.xlsx"), ("Todos os arquivos", "*.*")])

    messagebox.showinfo("Seleção", "Selecione a PRIMEIRA planilha")
    plan1 = filedialog.askopenfilename(
        title="Seleciona uma planilha",
        filetypes=file_types
    )
    if not plan1:
        return
    
    messagebox.showinfo("Seleção", "Selecione a SEGUNDA planilha")
    plan2 = filedialog.askopenfilename(
        title="Seleciona uma planilha",
        filetypes=file_types
    )
    if not plan2:
        return
    
    messagebox.showinfo("Seleção", "Selecione onde salvar o resultado")
    plansaida = filedialog.asksaveasfilename(
        title="Salvar como",
        defaultextension=".xlsx",
        filetypes=file_types,
        initialfile="resultado.xlsx"
    )
    if not plansaida:
        return
    
    try:
        wb1 = load_workbook(plan1)
        wb2 = load_workbook(plan2)

        wb_saida = Workbook()
        ws_saida = wb_saida.active
        ws_saida.title = "comparacao"

        fill_amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        borda = Border(left=Side(style='thin'), right=Side(style='thin'), 
                        top=Side(style='thin'), bottom=Side(style='thin'))

        ws1 = wb1.active
        ws2 = wb2.active

        max_row = max(ws1.max_row, ws2.max_row)

        cols_comparar = {1, 5}

        for row in range (1, max_row + 1):
            valores1 = {c: str(ws1.cell(row=row, column=c).value or "") for c in range (1, 6)}
            valores2 = {c: str(ws2.cell(row=row, column=c).value or "") for c in range (1, 6)}

            for c in range (1, 6):
                ws_saida.cell(row=row, column=c, value=valores1[c])
                ws_saida.cell(row=row, column=c+6, value=valores2[c])

            if all (valores1[c] == valores2[c] for c in cols_comparar):
                for col in list(range(1, 6)) + list(range(7, 12)):
                    celula = ws_saida.cell(row=row, column=col)
                    celula.fill = fill_amarelo

            for col in range(1, 12):
                celula = ws_saida.cell(row=row, column=col)
                if celula.value is not None:
                    celula.border = borda

        for col in ws_saida.columns:
            max_length = 0
            column = get_column_letter(col[0].column)
            
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass

            adjusted_width = (max_length + 2) * 1.2
            ws_saida.column_dimensions[column].width = adjusted_width

        wb_saida.save(plansaida)
        messagebox.showinfo("Sucesso! resultado salvo em ", plansaida)

    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao processar:\n{str(e)}")
    

if __name__ == "__main__":
    comparar()