# -*- coding: utf-8 -*-

import os
import pandas as pandas
import requests
import json
print("INICIANDO EJECUCION")
#----------------------------------CONFIG----------------------------------
URLAZURE='incluir url'
KEYAZURE='incluir key'
FILENAME='denunciayprotesta.xlsx'
SHEETNAME='Sheet1'
FILENAMEOUTPUT='denunciayprotestaprocesado.xlsx'
#----------------------------------CONFIG----------------------------------

denunciayprotestadataframe = pandas.read_excel(FILENAME,sheet_name=SHEETNAME)
denunciayprotestadataframe = denunciayprotestadataframe.astype(str)

for index, row in denunciayprotestadataframe.iterrows():
  print("--------------------Iniciando Consumo de API de Imagen: "+str(index+1)+"--------------------")
  UrlImagen=row['query imagen']
  headers = {"Content-Type": "application/json","Ocp-Apim-Subscription-Key": KEYAZURE}
  bodystring='{"url":"'+UrlImagen+'"}'   
  body = json.loads(bodystring)  
  params = {"visualFeatures": "Description,Color,Objects"}          
        
  #print(bodystring)
  #print(type(body))
  #print(headers)
  #print(body)
  #print(params)
  
  try:
    response = requests.post(URLAZURE,json=body,headers=headers,params=params)
  except exception:

    try:
      response = requests.post(URLAZURE,json=body,headers=headers,params=params)
    except exception:
      print("Error al consumir el webservice: "+str(exception))
      denunciayprotestadataframe.at[index,"Color dominante en primer plano"]="Error de Sistema al consumir la api: "+str(exception)
      denunciayprotestadataframe.at[index,"Color dominante en el fondo"]="Error de Sistema al consumir la api: "+str(exception)
      denunciayprotestadataframe.at[index,"Lista de colores dominantes"]="Error de Sistema al consumir la api: "+str(exception)
      denunciayprotestadataframe.at[index,"Descripción tags"]="Error de Sistema al consumir la api: "+str(exception)
      denunciayprotestadataframe.at[index,"Descripción texto"]="Error de Sistema al consumir la api: "+str(exception)
      denunciayprotestadataframe.at[index,"Objetos"]="Error de Sistema al consumir la api: "+str(exception)
      denunciayprotestadataframe.at[index,"Json"]="Error de Sistema al consumir la api: "+str(exception)
      continue

  responsestatus=response.status_code
  print("Codigo de respuesta del API: "+str(responsestatus))
  
  if responsestatus==200:

    try:
      json_response = response.json()
      json_responsestring=json.dumps(response.json())

    except exception:
      json_responsestring = "Error al obtener json: "+exception

    try:
      colorprimerplano = json_response['color']['dominantColorForeground']
    except exception:
      colorprimerplano = "Error al obtener colorprimerplano: "+exception

    try:
      colorfondo = json_response['color']['dominantColorBackground']
    except exception:
      colorfondo = "Error al obtener colorfondo: "+exception

    try:
      coloresdominantes = json_response['color']['dominantColors']
      coloresdominantesstring=","
      coloresdominantesstring=coloresdominantesstring.join(coloresdominantes)
    except exception:
      coloresdominantesstring="Error al obtener coloresdominantes: "+exception
  
      
    try:
      descripciontags= json_response['description']['tags']
      descripciontagsstring=","
      descripciontagsstring=descripciontagsstring.join(descripciontags)
    except exception:
      descripciontagsstring="Error al obtener descripciontags: "+exception

    try:
      descripciontexto=((json_response['description']['captions'])[0])['text']
    except IndexError:
      descripciontagsstring="Error al obtener descripciontexto posiblemente azure no arrojo resultados de descripcion: "+str(IndexError)
    except exception:
      descripciontagsstring="Error al obtener descripciontexto: "+exception
    try:
      
      objetos=json.dumps(json_response['objects'])
    except exception:
      objetos="Error al obtener objetos: "+exception 

    print(json_response)
    print(colorprimerplano)
    print(colorfondo)
    print(coloresdominantes)
    print(descripciontags)
    print(descripciontexto)
    print(objetos)
    denunciayprotestadataframe.at[index,"Color dominante en primer plano"]=colorprimerplano
    denunciayprotestadataframe.at[index,"Color dominante en el fondo"]=colorfondo
    denunciayprotestadataframe.at[index,"Lista de colores dominantes"]=coloresdominantesstring
    denunciayprotestadataframe.at[index,"Descripción tags"]=descripciontagsstring
    denunciayprotestadataframe.at[index,"Descripción texto"]=descripciontexto
    denunciayprotestadataframe.at[index,"Objetos"]=objetos
    denunciayprotestadataframe.at[index,"Json"]=json_responsestring

  else:

    json_response = response.json()
    json_responsestring=json.dumps(response.json())
    denunciayprotestadataframe.at[index,"Color dominante en primer plano"]="Error al consumir la api: "+json_responsestring
    denunciayprotestadataframe.at[index,"Color dominante en el fondo"]="Error al consumir la api: "+json_responsestring
    denunciayprotestadataframe.at[index,"Lista de colores dominantes"]="Error al consumir la api: "+json_responsestring
    denunciayprotestadataframe.at[index,"Descripción tags"]="Error al consumir la api: "+json_responsestring
    denunciayprotestadataframe.at[index,"Descripción texto"]="Error al consumir la api: "+json_responsestring
    denunciayprotestadataframe.at[index,"Objetos"]="Error al consumir la api: "+json_responsestring
    denunciayprotestadataframe.at[index,"Json"]="Error al consumir la api: "+json_responsestring

  print("Consumo de API de Imagen: "+str(index+1)+" Finalizada")
    
#print(index)
#print(row['url estatica imagen'])

print("Creando Archivo Procesado")
denunciayprotestadataframe.to_excel(FILENAMEOUTPUT,index=False) 
print("Creacion Archivo Procesado Finalizado")
print("EJECUCION FINALIZADA")