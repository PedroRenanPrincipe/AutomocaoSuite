# %%
import pyautogui as pya
import time
import pandas as pd
import os
import warnings
from docx2python import docx2python
import shutil
import pyperclip
warnings.filterwarnings("ignore")

def automation(resposta, pasta, zipadooudocx, nomedoarquivo):
    global stopflag
    stopflag = False
# %%
    def pesquisar(e):
        try:
            pya.click(pya.locateCenterOnScreen("1pesquisa.png", confidence=0.7))
        except:
            pya.click(1208,161)
        pyperclip.copy(e)
        time.sleep(1)
        pya.rightClick()
        time.sleep(1)
        pya.click(pya.locateCenterOnScreen("2paste.png", confidence=0.8))
        



    # %%
    def funcionando():
        while True:
            try:
                iamgem = pya.locateCenterOnScreen("1nãorespondendo.png", confidence=0.8)
                time.sleep(1)
            except:
                break

    # %%
    def noaguardode(e,v):
        tempo = 0
        while True:
            try:
                imaaagem = pya.locateCenterOnScreen(e,confidence=0.8 )
                time.sleep(1)
                break
            except:
                time.sleep(1)
                tempo+=1
                if tempo==40:
                    print(v)
                elif tempo >130:
                    break
                pass

    #noaguardode(r"C:\Users\jur.pedrorenan\Documents\automações\Suits\Imagens\buscarprocesso.png")



    # %%
    resposta = resposta
    caminho_da_pasta = pasta
    resposta1 = zipadooudocx


    # %%

    if resposta==".docx":
        # Defina o caminho da pasta que você deseja ler
        arquivos = os.listdir(caminho_da_pasta)

        name = [i.split(".")[0] for i in arquivos]

        name = {"Nomes": name}

        df = pd.DataFrame(name)
        df["status"]="Pendente"
        df["Observações"]=""

        df.insert(1, "Processo", "")
        for arquivo in arquivos:
            caminhodoarquivo = os.path.join(caminho_da_pasta,arquivo)
            with docx2python(caminhodoarquivo) as document:
                texto = document.text
                indice = texto.find(".8.")
                processo = texto[int(indice)-15 : int(indice)+10]
            for i,v in enumerate(df["Nomes"]):
                if arquivo.split(".")[0] == v:
                    df["Processo"].iloc[i] = processo
    elif resposta==".zip":
        if resposta1=="SIM":
            arquivos = os.listdir(caminho_da_pasta)
            name = [i.split(".ra")[0] for i in arquivos]
            name = {"Processo": name}
            df = pd.DataFrame(name)
            df["status"] = "Pendente"
            df["Observações"] = ""
        else:
            name = os.listdir(caminho_da_pasta)
            name = {"Processo": name}
            df = pd.DataFrame(name)
            df["status"]="Pendente"
            df["Observações"] = ""
    print(df)


    # %%
    def compactar_pasta_para_rar(caminho_pasta, nome_arquivo_rar, output_pasta):
    # Compacta a pasta para um arquivo zip
        shutil.make_archive(nome_arquivo_rar, 'zip', caminho_pasta)
    
    # Renomeia o arquivo zip para .rar
        os.rename(nome_arquivo_rar + '.zip', nome_arquivo_rar + '.rar')

        shutil.move(nome_arquivo_rar + '.rar', output_pasta)
        

        
    try:
        if resposta1=="NÃO":
            pya.alert("Vai começar a zipar as pastas")
            arquivoss = os.listdir(caminho_da_pasta)
            for i in arquivoss:
                compactar_pasta_para_rar(os.path.join(caminho_da_pasta,i), i, caminho_da_pasta)
                
            pya.alert("pastas zipadas")
            
    except:
        pass
    

    # %%
    def automação(arquivo, texto,  imagemrar, filesubmited):
        pya.alert("""Requisitos para começar o código.
                - Uma pasta no suite chamada 'pasta de download com os mesmo arquivos que voce quer baixar 
                - A vizualização do explorador de arquivos do suite tem que está no 'details'
                """)
        pya.PAUSE = 2
        time.sleep(1)
        for i,v in enumerate(df[f"{arquivo}"]):
            print(i,v)
            if stopflag:
                break
            try:
                noaguardode("3buscarpesquisa.png", "aguardo do botão pesquisa")
                pya.click(pya.locateCenterOnScreen("3buscarpesquisa.png", confidence=0.8))
                try:
                    imagem = pya.locateCenterOnScreen("3errorsemafor", confidence=0.8)
                    pya.click(pya.locateCenterOnScreen("2ok.png", confidence=0.7))
                    try:
                        imagem = pya.locateCenterOnScreen("3errosamafor2.png")
                        pya.click(pya.locateCenterOnScreen("2ok.png", confidence=0.7))
                    except:
                        pass
                    try:
                        pya.click(pya.locateCenterOnScreen("3errorindata.png", confidence=0.8))
                    except:
                        pass
                except:
                    pass
                        
                
                noaguardode("2_busca.png", "aguardo do botão buscar")
                busca = pya.locateCenterOnScreen("2_busca.png", confidence=0.8)
                pya.click(busca.x+50, busca.y)
                noaguardode("2processo.png", "aguardo do numero do processo")
                pya.click(pya.locateCenterOnScreen("2processo.png", confidence=0.8))
                noaguardode("2preencher.png", "no aguardo de preencher tabela")
                pya.click(pya.locateCenterOnScreen("2preencher.png", confidence=0.8))
                pya.doubleClick()
                pya.write(f"{df[f'{texto}'].iloc[i]}")
                pya.click(pya.locateCenterOnScreen("2ok.png", confidence=0.8))
                try:
                    pya.locateCenterOnScreen("2suitendigita.png", confidence=0.8)
                    pya.click(pya.locateCenterOnScreen("2okndigita.png", confidence=0.8))
                    pya.click(926, 142)
                    time.sleep(2)
                    pya.click(1342, 7)
                    time.sleep(2)
                    pya.click(662, 445)
                    noaguardode("1pastadedowloads.png", "no aguardo da pasta de downloadss")
                    pya.click(pya.locateCenterOnScreen("2suite.png", confidence=0.8))
                    noaguardode("3buscarpesquisa.png", "aguardo de botão de pesquisa")
                    pya.click(pya.locateCenterOnScreen("3buscarpesquisa.png", confidence=0.8))
                    noaguardode("2_busca.png", "aguardo do botão buscar")
                    busca = pya.locateCenterOnScreen("2_busca.png", confidence=0.8)
                    pya.click(busca.x+50, busca.y)
                    noaguardode("2processo.png", "aguardo do numero do processo")
                    pya.click(pya.locateCenterOnScreen("2processo.png", confidence=0.8))
                    noaguardode("2preencher.png", "no aguardo de preencher tabela")
                    pya.click(pya.locateCenterOnScreen("2preencher.png", confidence=0.8))
                    pya.doubleClick()
                    pya.write(f"{v}")
                    pya.click(pya.locateCenterOnScreen("2ok.png", confidence=0.8))
                    pass
                except: 
                    pass
                time.sleep(3)
                pya.doubleClick(501,250)
                

                funcionando()
                noaguardode("1documentos.png", "no aguardo de atos e documentos")
                #atos e documentos
                pya.click(pya.locateCenterOnScreen("1documentos.png", confidence=0.8))
                time.sleep(6)

                pya.click(1364, 738)
                noaguardode("1pastadedowloads.png", "aguardando a pasta de download")
                
                pya.doubleClick(pya.locateCenterOnScreen("1pastadedowloads.png", confidence=0.8))
                
                

                time.sleep(2)
                
                try:
                    imagem = pya.locateCenterOnScreen("1acessodeleitura.png", confidence=0.8)
                    pya.click(imagem)
                except: 
                    pass

                try:
                    pya.click(pya.locateCenterOnScreen("1telacheia.png"))
                except:
                    pass
                pesquisar(f"{df[f'{arquivo}'].iloc[i]}")
                noaguardode(f"{imagemrar}", "no aguardo do documento")
                time.sleep(2)
                pya.moveTo(pya.locateCenterOnScreen(f"{imagemrar}", confidence=0.8))
                pya.mouseDown()
                pya.moveTo(610,404)
                time.sleep(1.5)
                pya.moveTo(271,745)
                time.sleep(5)
                pya.moveTo(680,440)
                pya.mouseUp()

                noaguardode("1colardocumento.png", "no aguardo da tela de colocar o processo")
                funcionando()
                pya.click(792,568)
                time.sleep(1)
                try:
                    arquivojaexiste = pya.locateCenterOnScreen("1arquivojaexistente.png", confidence=0.8)
                    pya.click(865, 448)
                    pya.click(877, 569)
                    time.sleep(3)
                    pya.click(208,85)
                    time.sleep(2)

                    #fechar o arquivo
                    pya.click(220,751)
                    pya.click(219, 695)
                    try:
                        pya.click(pya.locateCenterOnScreen("1fechar_arquivo.png"))
                    except:
                        pya.click(1340,6)
                    
                    df["status"].iloc[i] = "Anexado"
                    df["Observações"].iloc[i] = "O arquivo ja existe"
                    continue
                
                except:
                    pya.click(676,443)
                    pass



                time.sleep(5)
                pya.click()
                noaguardode(f"{filesubmited}", "aguardando o arquivo ser anexado")
                try: 
                    imagem = pya.locateCenterOnScreen("3errorindata.png", confidence=0.8)
                    pya.click(imagem)
                except:
                    pass
                #fechar o processo
                funcionando()
                pya.click(pya.locateCenterOnScreen("1fechar_processo.png", confidence=0.8))
                funcionando()
                time.sleep(1.5)

                #fechar o arquivo
                pya.click(220,751)
                noaguardode(f"{imagemrar}", "no aguardo de abrir o explorador de arquivos")
                try:
                    pya.click(pya.locateCenterOnScreen("1fechar_arquivo.png", confidence=0.8))
                except:
                    pya.click(1340,6)
                
                df["status"].iloc[i] = "Anexado"
                

            #tentando voltar para o suite
            except:
                time.sleep(3.5)
                try: 
                    imagem = pya.locateCenterOnScreen("3errorindata.png", confidence=0.8)
                    pya.click(imagem)
                except:
                    pass

                
                try:
                    imagem = pya.locateCenterOnScreen("3errorsemafor", confidence=0.8)
                    pya.click(pya.locateCenterOnScreen("2ok.png", confidence=0.7))
                    try:
                        imagem = pya.locateCenterOnScreen("3errosamafor2.png")
                        pya.click(pya.locateCenterOnScreen("2ok.png", confidence=0.7))
                    except:
                        pass
                    try:
                        pya.click(pya.locateCenterOnScreen("3errorindata.png", confidence=0.8))
                    except:
                        pass
                except:
                    pass

                try:
                    imagem = pya.locateCenterOnScreen("2teladebuscar.png", confidence=0.8)
                    df["Observações"].iloc[i] = "Não conseguiu buscar o processo"
                    pya.click(pya.locateCenterOnScreen("2xteladebuscar.png", confidence=0.8))
                    continue
                except:
                    pass

                try:
                    imagem = pya.locateCenterOnScreen("1pastadedowloads.png", confidence=0.8)
                    df["Observações"].iloc[i] = "ficou na area de trabalho"
                    pya.click(pya.locateCenterOnScreen("1suite.png", confidence=0.8))
                    continue
                except:
                    pass
                
                try:
                    imagem = pya.locateCenterOnScreen("2viewex.png", confidence=0.8)
                    pya.click(imagem)
                    time.sleep(2)
                    try:
                        imagem2 = pya.locateCenterOnScreen("2details.png", confidence=0.8)
                        pya.click(imagem2)
                        pya.click(pya.locateCenterOnScreen("2xndigita.png", confidence=0.8))
                        noaguardode("1pastadedowloads.png", "na espera da pasta de download")
                        pya.click(pya.locateCenterOnScreen("1suite.png", confidence=0.8))
                        noaguardode("1fechar_processo.png", "no aguardo de fechar processo")
                        pya.click(pya.locateCenterOnScreen("1fechar_processo.png", confidence=0.8))
                        df["Observações"].iloc[i] = "não tava ativado o modo details"
                        continue
                    except:
                        pya.click(pya.locateCenterOnScreen("2xndigita.png", confidence=0.8))
                        noaguardode("1pastadedowloads.png", "na espera da pasta de download")
                        pya.click(pya.locateCenterOnScreen("1suite.png", confidence=0.8))
                        noaguardode("1fechar_processo.png", "no aguardo de fechar processo")
                        pya.click(pya.locateCenterOnScreen("1fechar_processo.png", confidence=0.8))
                        df["Observações"].iloc[i] = "não achou o processo"
                        continue
                except:
                    pass
                try:
                    imagem =  pya.locateCenterOnScreen("1fechar_processo.png", confidence=0.8)
                    #fechar o processo
                    funcionando()
                    pya.click(pya.locateCenterOnScreen("1fechar_processo.png", confidence=0.8))
                    funcionando()
                    time.sleep(1.5)
                    continue
                except:
                    pass

                try:
                    imagem = pya.locateCenterOnScreen("3errorsemafor", confidence=0.8)
                    pya.click(pya.locateCenterOnScreen("2ok.png", confidence=0.7))
                    try:
                        imagem = pya.locateCenterOnScreen("3errosamafor2.png")
                        pya.click(pya.locateCenterOnScreen("2ok.png", confidence=0.7))
                    except:
                        pass
                    try:
                        pya.click(pya.locateCenterOnScreen("3errorindata.png", confidence=0.8))
                    except:
                        pass
                except:
                    pass




        pya.alert("concluido")


    # %%
    if resposta==".docx":
        automação("Nomes", "Processo", "1word.png", "2arquivoanexado.png")
    elif resposta==".zip":
        automação("Processo", "Processo", "2rar.png", "2raranexado.png")

    total_processo = df["Processo"].count()

    print(df)
    # %%
    df.to_excel(os.path.join(caminho_da_pasta, f"{nomedoarquivo}.xlsx"), index=False)

def stop_automation(e):
    global stopflag
    stopflag=True

