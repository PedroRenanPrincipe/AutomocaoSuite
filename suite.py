# %%
import pyautogui as pya
import time
import pandas as pd
import os
import warnings
warnings.filterwarnings("ignore")

# %%


# Defina o caminho da pasta que você deseja ler
caminho_da_pasta = r"C:\Users\jur.pedrorenan\Desktop\suite"

arquivos = os.listdir(caminho_da_pasta)

name = [i.split(".")[0] for i in arquivos]

name = {"Nomes": name}


# %%
df = pd.DataFrame(name)
df["status"]="Pendente"
df["Observações"]=""

# %%
#pya.doubleClick(184,335)
#pya.sleep(40)


# %%
def pesquisar(e):
    try:
        pya.click(pya.locateCenterOnScreen(r"C:\Users\jur.pedrorenan\Documents\automações\Suits\Imagens\pesquisa.png"))
    except:
        pya.click(1208,161)
    pya.write(f"{e}")
    time.sleep(10)


# %%
def funcionando():
    while True:
        try:
            iamgem = pya.locateCenterOnScreen("1nãorespondendo.png")
            time.sleep(1)
        except:
            break
funcionando()

# %%
def noaguardode(e):
    while True:
        try:
            imaaagem = pya.locateCenterOnScreen(e,confidence=0.8 )
            time.sleep(1)
            break
        except:
            time.sleep(1)
            pass

#noaguardode(r"C:\Users\jur.pedrorenan\Documents\automações\Suits\Imagens\buscarprocesso.png")

# %%
def qtspendente():
    try:
        abc =list(pya.locateAllOnScreen("1pendente.png", confidence=0.8))
        cont = 0
        for i in abc:
            cont+=1
        return cont
    except:
        return 0


# %%

pya.alert("o código vai começar")
pya.PAUSE = 7
time.sleep(1)
for i,processo in enumerate(df["Nomes"]):
    try:
        pya.doubleClick(698,65)
        
        noaguardode("1buscarprocesso.png")

        #ativar o modo para pesquisar pela partefdgg
        pya.click(498, 360)
        #escrever a parte e apeartar ok
        #imagem = pya.locateCenterOnScreen(r"C:\Users\jur.pedrorenan\Documents\automações\Suits\Imagens\contraparte.png")
        pya.click(600, 411)
    
        pya.write(processo)
        pya.click(644,498)
        time.sleep(6) 

        #problema de digitar e não aparecer
        try:
            atençao = pya.locateCenterOnScreen("1atenção.png")
            pya.click(pya.locateCenterOnScreen("1oko.png"))
            pya.click(874,253)
            df["Observações"].iloc[i] = "Error no suite de não digitalizar"
            break
        except:
            pass
        
        
        try:
            telamultiprocessos = pya.locateCenterOnScreen("1telamultiprocessos.png", confidence=0.8)
            cont = qtspendente()
            if cont!=1:
                try: 
                    pya.click(pya.locateCenterOnScreen("1anular.png"))
                except:
                    pya.click(953, 598)
                pya.click(874, 253)

                df["Observações"].iloc[i] = "Mais de um processo no nome do autor"

                continue
            else:
                pya.doubleClick(pya.locateCenterOnScreen("1pendente.png", confidence=0.8))
                pya.click(875,593)
                
        except:
            pass

        funcionando()
        noaguardode("1documentos.png")
        #atos e documentos
        pya.click(706,140)
        time.sleep(6)

        pya.click(1364, 738)
        time.sleep(2)
        try:
            pya.doubleClick(pya.locateCenterOnScreen("1pastadedownloads.png"))
        except:
            pya.doubleClick(198, 453)
        

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
        pesquisar(f"{processo}")

        time.sleep(2)
        pya.moveTo(pya.locateCenterOnScreen("1word.png"))
        pya.mouseDown()
        pya.moveTo(610,404)
        time.sleep(1.5)
        pya.moveTo(271,745)
        time.sleep(5)
        pya.moveTo(680,440)
        pya.mouseUp()

        noaguardode("1colardocumento.png")
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
            
            df["Observações"].iloc[i] = "O arquivo ja existe"
            continue
        
        except:
            pya.click(676,443)
            pass



        time.sleep(3)

        #fechar o processo
        pya.click(208,85)
        time.sleep(2)

        #fechar o arquivo
        pya.click(220,751)
        pya.click(219, 695)
        try:
            pya.click(pya.locateCenterOnScreen("1fechar_arquivo.png"))
        except:
            pya.click(1340,6)
        
        df["status"].iloc[i] = "Concluido"


    except:
        df["Observações"].iloc[i] = "error desconhecido"





    # pya.click(1554, 577)
    # pya.click(1736,197)
    # pya.sleep(4)
    # pya.click(1724,311)
    # pya.click(1726,351)

# %%
print(df)


