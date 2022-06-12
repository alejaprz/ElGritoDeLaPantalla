from selenium import webdriver
import os
import pandas as pandas 
import pyautogui
import pyperclip
import time
from selenium.webdriver.chrome.options import Options as ChromeOptions
import pathlib


print("INICIANDO EJECUCION")
ruta=str(pathlib.Path().resolve())
print("RUTA: "+ ruta)
options = webdriver.ChromeOptions()        
options.add_argument("--disable-gpu")      
options.add_argument("--headless")
driver = webdriver.Chrome(executable_path=ruta + "/chromedriver.exe")
fileName = ruta+'/Input/Conjunto de datos_Denuncia y protesta feminista_2021.xlsx'
denunciayprotestadataframe = pandas.read_excel(fileName,sheet_name="Conjunto de datos",skiprows = 1)
denunciayprotestadataframe = denunciayprotestadataframe.astype(str)
#denunciayprotestadataframe.to_excel("prueba.xlsx")
contador=1

for index, row in denunciayprotestadataframe.iterrows():
      
    UrlImagen=row['Query Image']
    print('Imagen a descargar: '+UrlImagen)
    driver.get(UrlImagen)
    time.sleep(1)   
    driver.maximize_window()
    Ruta=ruta+ r'\Output\00'+str(contador)+'.jpg'
    time.sleep(1)   
    pyautogui.hotkey('ctrl','s')
    time.sleep(1)   
    pyperclip.copy(str(Ruta))
    time.sleep(1)   
    pyautogui.press('left')
    time.sleep(1)   
    pyautogui.hotkey('ctrl', 'a', interval = 0.15)
    pyautogui.hotkey('del', interval = 0.15)
    pyautogui.hotkey('ctrl', 'v', interval = 0.15)  
    time.sleep(1)   
    pyautogui.press('enter')
    contador=contador+1  
    time.sleep(7)   

driver.close()
print("EJECUCION FINALIZADA")