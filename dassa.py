import pyautogui as pya
import time

def noaguardode(e):
    cont=0
    while cont<4:
        try:
            imaaagem = pya.locateCenterOnScreen(e, grayscale=True)
            print("VAMOO krlho!!!")
            time.sleep(1)
            break
        except:
            time.sleep(1)
            cont+=1
            print("esperando")
            pass

pya.locateOnScreen(r"C:\Users\jur.pedrorenan\Documents\automações\Suits\Imagens\1_buscar.png")

#noaguardode(r"C:\Users\jur.pedrorenan\Documents\automações\Suits\Imagens\buscarprocesso.png")