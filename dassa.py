import pandas as pd
import flet as ft

# Cria um DataFrame do pandas
df = pd.DataFrame({
    'First name': ['John', 'Jack', 'Alice'],
    'Last name': ['Smith', 'Brown', 'Wong'],
    'Age': [43, 19, 25]
})

def main(page: ft.Page):
    # Converte o DataFrame do pandas para uma tabela Flet
    page.add(
        ft.DataTable(
            columns=[ft.DataColumn(ft.Text(name)) for name in df.columns],
            rows=[
                ft.DataRow(
                    cells=[ft.DataCell(ft.Text(str(value))) for value in row]
                ) for row in df.values
            ],
        ),
    )

ft.app(target=main)
