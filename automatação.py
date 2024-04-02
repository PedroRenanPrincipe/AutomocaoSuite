import flet as ft 
import pandas as pd 
import os
from suite import automation, stop_automation
import time
from dassa import texto


def main(page: ft.Page):
    

    def on_dialog_result(e: ft.FilePickerResultEvent):
        e.path
        print("Selected file or directory:", e.path)
        print(file_picker.result.path)
    
    file_picker = ft.FilePicker(on_result=on_dialog_result)

    page.overlay.append(file_picker)
    page.update()

    def saibamais(e):
        page.clean()

        page.add(ft.Container(
            content= ft.Column([
            ft.Row([ft.Icon(ft.icons.ROCKET_LAUNCH, size=50),
        ft.Text("LAUDINHO", size= 60 , font_family= "Kanit")
        ],
        ft.MainAxisAlignment.CENTER),]), padding=40))

        page.add(ft.Row([ft.Text("Requisitos para rodar o Laudinho:", size=25)], ft.MainAxisAlignment.CENTER)   ,
            ft.Column([ft.Row([ft.Text(texto)], ft.MainAxisAlignment.CENTER),
                        ft.Row([ft.ElevatedButton("Voltar", on_click=voltar)], ft.MainAxisAlignment.CENTER)]))

        page.update()


    def voltar(e):
        page.clean()
        page.add(main_page) # Adicionar a página principal de volta
        page.update()
        
    page.title = "Laudinho"
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.fonts = {
        "Kanit": "https://raw.githubusercontent.com/google/fonts/master/ofl/kanit/Kanit-Bold.ttf"
    }

    

    def continuar(e):
        page.clean()

        def executar(e):
            arquivos = os.listdir(file_picker.result.path)
            tinicial = time.time()
            result_view.controls.append(ft.Column([ft.Text(""),
                                                   ft.Row([
                ft.ElevatedButton("Interromper", on_click=stop_automation)], ft.MainAxisAlignment.CENTER)
            ], ft.MainAxisAlignment.CENTER))
            page.update()
            automation(tipodedocumento.value, file_picker.result.path, pasta_zipada.value, file_name.value)
            tfinal = time.time()
            total_time = tfinal - tinicial
            result_view.controls.append(ft.Row(
                [ft.Text(f"O tempo total foi: {total_time:.1f}"),
                 ft.Text(),
                 ft.Text(),
                 ft.Text(f"A média de tempo por processo: {(total_time/len(arquivos)):.1f}"),
                 ft.Text(),
                 ft.Text(),
                 ft.Text(len(arquivos))], ft.MainAxisAlignment.CENTER
            ))
            page.update()

        

        def aparecer(l):
            result_view.controls.clear()


            print(tipodedocumento.value)
            if tipodedocumento.value == ".zip":
                result_view.controls.append(ft.Column([
                            ft.Row([pasta_zipada], ft.MainAxisAlignment.CENTER),
                    ft.Text(""),
                    ft.Row([ft.ElevatedButton("voltar" ,on_click=voltar), ft.ElevatedButton("Executar", on_click=executar)], ft.MainAxisAlignment.CENTER)])
                )
                
                

            elif tipodedocumento.value == ".docx":
                result_view.controls.append(ft.Row([ft.ElevatedButton("voltar" ,on_click=voltar), ft.ElevatedButton("Executar", on_click=executar)], ft.MainAxisAlignment.CENTER)
                )
            page.update()
                
                    
                

        tipodedocumento= ft.Dropdown(
            label= "Tipo de documento",
            hint_text= "Escolha qual é o tipo de documento da pasta",
            options = [
                ft.dropdown.Option(".docx"),
                ft.dropdown.Option(".zip")
            ],
            on_change=aparecer
        )

        pasta_zipada = ft.Dropdown(
                label= "Pasta está zipada?",
                options= [
                    ft.dropdown.Option("SIM"),
                    ft.dropdown.Option("NÃO")]

        )

        file_name = ft.TextField(label="nome da planilha de extração?")

        result_view = ft.Column()

        page.add(ft.Container(
            content= ft.Column([
            ft.Row([ft.Icon(ft.icons.ROCKET_LAUNCH, size=50),
        ft.Text("LAUDINHO", size= 60 , font_family= "Kanit")
        ],
        ft.MainAxisAlignment.CENTER),]), padding=40))
        
        page.add(ft.Container(
            content= 
            ft.Column([
        ft.Row([ft.ElevatedButton("Selecionar a pasta",icon= ft.icons.FOLDER, on_click= lambda _: file_picker.get_directory_path(dialog_title=
                        "Pasta dos arquivos"))], ft.MainAxisAlignment.CENTER),
        ft.Text(""),
        ft.Row([file_name], ft.MainAxisAlignment.CENTER),
        ft.Text(""),
        ft.Row([tipodedocumento], ft.MainAxisAlignment.CENTER),
        ft.Text(""),
        ft.Row([
        result_view], ft.MainAxisAlignment.CENTER)]), padding=30))
        
        page.update()

        
            

        
               
        

    # Armazenar o estado da página principal
    main_page = ft.Container(
        content=ft.Column([
        ft.Row([ft.Icon(ft.icons.ROCKET_LAUNCH, size=50),
        ft.Text("LAUDINHO", size= 60 , font_family= "Kanit")
        ],
        ft.MainAxisAlignment.CENTER), 

        ft.Text(""),
        ft.Text(""),

        ft.Row([ft.Text("Requisitos para rodar o Laudinho", size=15)], ft.MainAxisAlignment.CENTER),
                           ft.Row([ft.ElevatedButton("Saiba mais", on_click=saibamais),
                                   ft.ElevatedButton("Continuar", on_click=continuar)], ft.MainAxisAlignment.CENTER)
    ]),
    padding=200 # Adicionando padding ao Container
    )


    page.add(main_page)

ft.app(main)