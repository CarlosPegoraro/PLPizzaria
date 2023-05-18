import pyautogui as py
import time
import pandas as pd
import pyperclip as pc
import plotly.express as px

py.press("win")
py.write("word")
py.press("enter")

time.sleep(3)

py.press("enter")
time.sleep(3)

py.write("Lista de precos atualizada")
py.press("enter")


table = pd.read_excel(r"D:\5_Programação\PLPizzaria\auto\assets\database.xlsx")

print(table)

table["CUSTO"] = pd.to_numeric(table["CUSTO"], errors="coerce")

lucre = 1.9

impost = 1.16

for line in  table.index:
    
    name = table.loc[line, "NAME"]
    value = table.loc[line, "CUSTO"]
    #multiplique  o valor por x porcentagem
    valueL = value * lucre
    # pegue o valor do produto e aplique o imposto
    price = valueL * impost
    # INSIRA NO CAMPO PRICE O VALOR DO PRODUTO FINAL
    table.loc[line, "PRICE"] = price
    # PEGUE O NOME DO PRODUTO E O SEU VALOR FINAL E MONTE UMA LISTA COM ELES
    product = f"{name} - R$ {price:,.2f}"
    py.write(product)
    py.press("enter")
    # INSIRA A LISTA EM UM DOCUMENTO WORD

table.to_excel(f"D:/5_Programação/PLPizzaria/auto/assets/updateDatabase.xlsx", index=False)
