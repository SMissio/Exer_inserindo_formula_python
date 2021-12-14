from openpyxl import load_workbook
import os

caminho_arquivo = "C:\\Users\\SylviaMissio\\Desktop\\RPA1\\openpyxl\\Formulas.xlsx"

plan_aberta = load_workbook(filename=caminho_arquivo)
sheet_seleciona = plan_aberta['Aluno']
sheet_seleciona['A6'] = "=SUM(A2:A5)"
sheet_seleciona['B6'] = "=SUM(B2:B5)"
sheet_seleciona['D2'] = "=A2+B2"
sheet_seleciona['D3'] = "=A3-B3"
sheet_seleciona['D4'] = "=A4*B4"
sheet_seleciona['D5'] = "=A5/B5"

sheet_seleciona['B12'] = "=MID(A12,1,3)"
sheet_seleciona['C12'] = "=MID(A12,5,3)"
sheet_seleciona['D12'] = "=MID(A12,9,3)"
sheet_seleciona['E12'] = "=MID(A12,13,2)"

plan_aberta.save(filename=caminho_arquivo)

os.startfile(caminho_arquivo)