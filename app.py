"""
Created on Tue Aug  6 13:48:22 2024

@author: lportilla
"""
import pandas as pd
import numpy as np
import os
import holidays
from datetime import datetime, timedelta
import base64
import io

import dash
import dash_core_components as dcc
from dash.dash_table import DataTable
import dash_html_components as html
import dash_bootstrap_components as dbc
import plotly.graph_objs as go
from dash.dependencies import Input, Output, State
from dash import no_update


def cargar_datos_excel( fecha):
    fecha_3 = fecha.date()
    dia_semana = fecha_3.strftime('%A')
 
    if dia_semana in ['Tuesday', 'Wednesday', 'Thursday', 'Friday']:
        hojas = ['martes', 'miercoles', 'jueves', 'viernes']
    if dia_semana == 'Monday':
        hojas = ['lunes']
    if dia_semana == 'Saturday':
        hojas = ['sabado']
    if dia_semana == 'Sunday':
        hojas = ['domingo']
    if es_feriado(fecha):
        hojas = ['domingo']
        
    dataframes = []
    fecha_referencia = fecha.date()    

    if dia_semana in ['Tuesday', 'Wednesday', 'Thursday', 'Friday']:
        fecha_inicio = fecha_referencia - timedelta(weeks=2)
    else:
        fecha_inicio = fecha_referencia - timedelta(weeks=8)
    
    directorio = os.getcwd()
    #ruta_archivo = os.path.join(directorio, hojas[0] + '.csv')
    for hoja in hojas:        
        df = pd.read_csv(os.path.join(directorio, hoja + '.csv'))      
        #df['fecha_hora'] = pd.to_datetime(df['fecha_hora'])
        df['fecha_hora'] = pd.to_datetime(df['fecha_hora'])
        df = df.query('@fecha_inicio <= fecha_hora <= @fecha_referencia')  
        df = df.melt(id_vars=['fecha_hora'], var_name='usuario', value_name='demanda')     
        dataframes.append(df)      
        df= pd.DataFrame()

    datos_completos = pd.concat(dataframes, ignore_index=True)  
    return datos_completos

def es_feriado(fecha):
    peru_holidays = holidays.Peru()
    return fecha in peru_holidays

def fechas_feriado_cercano(fecha):
    fecha = fecha.date()
    fecha_actual = fecha
    dia_semana = fecha.strftime('%A')
    
    while True: 
        fecha_actual -= timedelta(days=7)  
        if es_feriado(fecha_actual) and fecha_actual.strftime('%A') == dia_semana:
                return fecha

        if fecha_actual < fecha - timedelta(days=300):
            print("No se encontro feriado semejante")
            return []
            break
        
def fechas_domingos(fecha):
    domingos_encontrados = []
    fecha = fecha.date()
    fecha_actual = fecha
    while len(domingos_encontrados) < 4:
        fecha_actual -= timedelta(days=1)
        if fecha_actual.strftime('%A') == 'Sunday':
            domingos_encontrados.append(fecha_actual)
            
        if fecha_actual < fecha -  timedelta(days=60):
            break
    return domingos_encontrados

def fechas_dia_no_tipico(fecha):   
    dia_semana = fecha.strftime('%A')  
    fecha_extraida = []
    sem_encontrada = 0
    i = 1
    if dia_semana == 'Saturday':
        while sem_encontrada < 4 and i <= 10:
            if not es_feriado(fecha):
                fecha_dia = fecha - timedelta(weeks=i)
                fecha_extraida.append(fecha_dia)
                sem_encontrada = sem_encontrada + 1
            i = i + 1
        return fecha_extraida
            
    elif dia_semana == 'Sunday':
        while sem_encontrada < 4 and i <= 10:
            if not es_feriado(fecha):
                fecha_dia = fecha - timedelta(weeks=i)
                fecha_extraida.append(fecha_dia)
                sem_encontrada = sem_encontrada + 1
            i = i + 1
        return fecha_extraida     
                
    elif dia_semana == 'Monday':
        while sem_encontrada < 4 and i <= 10:
            if not es_feriado(fecha):
                fecha_dia = fecha - timedelta(weeks=i)
                fecha_extraida.append(fecha_dia)
                sem_encontrada = sem_encontrada + 1
            i = i + 1
        return fecha_extraida
    else:
        return print("No es Sabado, Domingo o Lunes")  

def fechas_dia_tipico_sem_actual(fecha):
    fecha = fecha.date()
    dia_semana = fecha.strftime('%A')
    dias_a_comparar = []
    
    if dia_semana == 'Friday':
        dias_a_comparar = ['Thursday', 'Wednesday', 'Tuesday']
    elif dia_semana == 'Thursday':
        dias_a_comparar = ['Wednesday', 'Tuesday']
    elif dia_semana == 'Wednesday':
        dias_a_comparar = ['Tuesday']
    
    fechas_a_comparar = []
    dias_semana = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
    
    primer_dia_semana = fecha - timedelta(days=fecha.weekday()) - timedelta(days=1) 
    ultimo_dia_semana = primer_dia_semana + timedelta(days=5)
    
    for dia in dias_a_comparar:       
        dias_a_restar = fecha.weekday() - dias_semana.index(dia)
        fecha_comparar = fecha - timedelta(days=dias_a_restar)       
       
        if primer_dia_semana <= fecha_comparar <= ultimo_dia_semana and not es_feriado(fecha_comparar):
            fechas_a_comparar.append(fecha_comparar)
    
    return fechas_a_comparar

def fechas_dia_tipico_semanal(fecha):
    fecha = fecha.date()
    dia_semana = fecha.strftime('%A')

    fechas = {
        'Tuesday': [],
        'Wednesday': [],
        'Thursday': [],
        'Friday': []
    }
    
    lunes_actual = fecha - timedelta(days=fecha.weekday())

    for i in range(1, 3): 
        lunes_semana_anterior = lunes_actual - timedelta(weeks=i)
        martes_anterior = lunes_semana_anterior + timedelta(days=1)  # Martes
        miercoles_anterior = lunes_semana_anterior + timedelta(days=2)  # Miércoles
        jueves_anterior = lunes_semana_anterior + timedelta(days=3)  # Jueves
        viernes_anterior = lunes_semana_anterior + timedelta(days=4)  # Viernes

        if es_feriado(martes_anterior) or (i == range(1,3)[-1]):
            fechas['Tuesday'].append(None)
        else:
            fechas['Tuesday'].append(martes_anterior)
        
        if es_feriado(miercoles_anterior) or (i == range(1,3)[-1]):
            fechas['Wednesday'].append(None)
        else:
            fechas['Wednesday'].append(miercoles_anterior)
        
        if es_feriado(jueves_anterior) or ((dia_semana == "Friday" or dia_semana ==  "Thursday") and i == range(1,3)[-1]):
            fechas['Thursday'].append(None)
        else:
            fechas['Thursday'].append(jueves_anterior)
        
        if es_feriado(viernes_anterior) or ((dia_semana == "Friday" or dia_semana ==  "Thursday") and i == range(1,3)[-1]):
            fechas['Friday'].append(None)
        else:
            fechas['Friday'].append(viernes_anterior)
    lista_tuplas = list(fechas.items())
    lista_tuplas_invertida = lista_tuplas[::-1]
    fechas_2 = dict(lista_tuplas_invertida)
        
    return fechas_2


def fechas_dem_total(tabla_datos_totales_final):   
    a =list(tabla_datos_totales_final.columns.tolist())
    del a[0]
    del a[0]
    return a

def dem_total(tabla_datos_totales_final):  
    tabla_dem_total = pd.DataFrame()
    
    for col in tabla_datos_totales_final.columns:
            if col not in ['usuario', 'hora']:
               tabla_dem = tabla_datos_totales_final.pivot_table(index=["hora"], columns="usuario", values=col, aggfunc="sum") 
               tabla_dem = tabla_dem.reset_index()
               tabla_dem = pd.concat([tabla_dem, pd.DataFrame({'fecha': [col for i in range(48)]})], axis=1)
               tabla_dem_total = pd.concat([tabla_dem_total, tabla_dem], ignore_index=False)
               tabla_dem = pd.DataFrame()
    tabla_dem_total = tabla_dem_total.set_index(['hora', 'fecha'])
    for col in tabla_dem_total.columns:
        tabla_dem_total[col] = pd.to_numeric(tabla_dem_total[col], errors='coerce') 
    tabla_dem_total['total'] = tabla_dem_total.sum(axis=1)
    tabla = tabla_dem_total['total']
    tabla = tabla.reset_index()
    tabla = tabla.pivot_table(index=["hora"], columns="fecha", values='total', aggfunc="sum")
    tabla = tabla.reset_index()
    
    return tabla
 
def mover_a_final(usuario_2, lista_bajada):
    
    repetidos = []

    for item in lista_bajada:
        if item in usuario_2 and item not in repetidos:
            repetidos.append(item)
    
    for item in repetidos:
        while item in usuario_2:
            usuario_2.remove(item)

    usuario_2 = usuario_2 + repetidos
    
    return usuario_2

def aproximacion_polinimica_curva_analisis(fecha, hora_1, usuarios, df_curva_analisis_dem_a):
    usuario_2, hora_2, desv_2, ajuste, valor_real, valor_maximo, valor_cambio = [], [], [], [], [], [], []
    df_curva_analisis = pd.DataFrame()
    df_dic = pd.DataFrame()
    
    
    for usuario in usuarios['usuarios']:
        
        
        curva_analisis = pd.DataFrame()  
        curvas_analisis = pd.DataFrame()     
        desviacion = pd.DataFrame()                
        
        curva_analisis = df_curva_analisis_dem_a.loc[(df_curva_analisis_dem_a['usuario'] == usuario) & (df_curva_analisis_dem_a['fecha_hora'].dt.date == fecha.date()) , 'demanda'].values                         
        curvas_analisis = pd.concat([curvas_analisis, pd.DataFrame(curva_analisis, columns=['usuario_1']), pd.DataFrame({'hora': [ i for i in range(48)]})], axis=1)
       
        curvas_analisis['usuario_1'] = pd.to_numeric(curvas_analisis['usuario_1'], errors='coerce') 
        
        x = curvas_analisis['hora'].values
        y = curvas_analisis['usuario_1'].values
            
        coeficientes = np.polyfit(x, y, 150)
        polinomio = np.poly1d(coeficientes)
        
        poli = polinomio(x)
        curvas_analisis['ajuste'] = poli
        
        
        curvas_analisis['desviacion'] = 100 * (curvas_analisis['ajuste'] - curvas_analisis['usuario_1']) / abs(curvas_analisis['ajuste'])
        curvas_analisis['diferencia'] = curvas_analisis['usuario_1'].diff()
        
        desviacion['atipico'] = pd.DataFrame(curvas_analisis['desviacion'])
        desviacion['ajuste'] = curvas_analisis['ajuste']
        desviacion['usuario_1'] = curvas_analisis['usuario_1']
        desviacion['diferencia'] = curvas_analisis['diferencia'] 
        
        curvas_analisis = pd.concat([curvas_analisis ,pd.DataFrame({'usuario': [usuario for i in range(48)]}), pd.DataFrame({'hora_1': hora_1})], axis=1)
        
        df_curva_analisis = pd.concat([df_curva_analisis, curvas_analisis], ignore_index= True)

        for indice, valor in desviacion.iterrows():                    

           a = timedelta(minutes=0)
           if indice == 47: a = timedelta(minutes=1)            
           hora = (indice + 1) * timedelta(minutes=30) - a
           maximo = np.max(abs(curvas_analisis['usuario_1']))
           
           
           if abs(valor[0]) >= 8 and abs(valor[2]) > 40 and abs(valor[3]) > 10:  
                       usuario_2.append(usuario) 
                       hora_2.append(hora)
                       desv_2.append(valor[0])
                       ajuste.append(valor[1])
                       valor_real.append(valor[2])
                       valor_maximo.append(maximo)
                       valor_cambio.append(valor[3]) 
                   
           elif abs(valor[0]) >= 20 and abs(valor[2]) > 16 and abs(valor[2]) < 40 and abs(valor[3]) > 10:                   
                   usuario_2.append(usuario) 
                   hora_2.append(hora)
                   desv_2.append(valor[0])
                   ajuste.append(valor[1])
                   valor_real.append(valor[2])
                   valor_maximo.append(maximo)
                   valor_cambio.append(valor[3]) 
                   
           elif abs(valor[0]) >= 60 and abs(valor[2]) <= 16 and abs(valor[2]) >= 5 and abs(valor[3]) >= 10:
                   usuario_2.append(usuario)
                   hora_2.append(hora)
                   desv_2.append(valor[0])
                   ajuste.append(valor[1])
                   valor_real.append(valor[2])
                   valor_maximo.append(maximo)
                   valor_cambio.append(valor[3]) 
               
           elif abs(valor[0]) >= 20 and abs(valor[2]) <= 16 and abs(valor[2]) >= 0 and maximo >= 30 and abs(valor[3]) >= 10:
                   usuario_2.append(usuario)
                   hora_2.append(hora)
                   desv_2.append(valor[0])
                   ajuste.append(valor[1])
                   valor_real.append(valor[2])
                   valor_maximo.append(maximo)
                   valor_cambio.append(valor[3]) 


            
        datos_dict = {
                'usuario': usuario_2,
                'hora': hora_2,
                'desviacion': desv_2,
                'valor_real': valor_real,
                'ajuste': ajuste,
                'max': valor_maximo,
                'diferencia': valor_cambio
                }
        
        filtro = pd.DataFrame(datos_dict)
        df_dic = pd.concat([df_dic, filtro], ignore_index=True) 
        
    df_sorted = pd.DataFrame()
    df_dic = df_dic.drop_duplicates(subset='desviacion')
    
    df_sorted = df_dic.sort_values(by='max', ascending=False)  
    df_sorted_2 = df_sorted.drop_duplicates(subset='usuario')
    
    usuario_2 = list(df_sorted_2['usuario'])
    
    lista_bajada = list(['Cajamarq', 'AceArequ', 'Chimbot1', 'Independ', 'Botiflac', 
                         'Cachimay', 'Ica', 'KimAyllu', 'ParamExi', 'Chincha',
                         'Valledel', 'Yura', 'PushBack','Lixiviac'])
    
    usuario_2 = mover_a_final(usuario_2, lista_bajada)
   
    return usuario_2, df_dic , df_curva_analisis 

def contar_nulos(data, fecha):
    
    nulos_antes = data.isnull().sum().sum()
    data.fillna(method='ffill', inplace=True)     
    nulos_despues = data.isnull().sum().sum()
    valores_completados = nulos_antes - nulos_despues
    if valores_completados > 3:
        print(f"{valores_completados} valores completados en {fecha.date()}")
       
def curvas_hist(fecha, df, usuario):
    
    dia_semana = fecha.strftime('%A')
    curvas_historicas = pd.DataFrame()
    
    if dia_semana in ['Tuesday', 'Wednesday', 'Thursday', 'Friday']:
        fechas_extraidas = fechas_dia_tipico_sem_actual(fecha)
    if dia_semana in ['Monday', 'Saturday', 'Sunday']:
        fechas_extraidas = fechas_dia_no_tipico(fecha)
    if es_feriado(fecha):
        fechas_extraidas = fechas_domingos(fecha)     
    
    for fecha_1 in fechas_extraidas:
        fecha_1 = pd.to_datetime(fecha_1)
                  
        curva_historica = df.loc[(df['usuario'] == usuario) & (df['fecha_hora'].dt.date == fecha_1.date()), 'demanda'].values
        curva_historica = pd.DataFrame(curva_historica, columns=[fecha_1.strftime('%Y-%m-%d')])
        curvas_historicas = pd.concat([curvas_historicas, curva_historica], axis=1)
                    
        contar_nulos(curvas_historicas, fecha_1)                 
    return curvas_historicas

def curvas_hist_dia_tipico(fecha, df, usuario, curvas_historicas):
    
    for fila in pd.DataFrame(fechas_dia_tipico_semanal(fecha)).itertuples(index=False):
                                                  
        if fila.Tuesday != None and not es_feriado(pd.to_datetime(fila.Tuesday).date()):
                                curva_historica_martes = df.loc[(df['usuario'] == usuario) & (df['fecha_hora'].dt.date == pd.to_datetime(fila.Tuesday).date()), 'demanda'].values
                                martes_1 = pd.DataFrame( curva_historica_martes, columns=[f'Semana_{pd.to_datetime(fila.Tuesday).isocalendar()[1]}_{pd.to_datetime(fila.Tuesday).date()}'])
        else:
                      martes_1 = pd.DataFrame()   
                       
        if fila.Wednesday != None and not es_feriado(pd.to_datetime(fila.Wednesday).date()):
                                curva_historica_miercoles = df.loc[(df['usuario'] == usuario) & (df['fecha_hora'].dt.date == pd.to_datetime(fila.Wednesday).date()), 'demanda'].values
                                miercoles_1 = pd.DataFrame( curva_historica_miercoles, columns=[f'Semana_{pd.to_datetime(fila.Wednesday).isocalendar()[1]}_{pd.to_datetime(fila.Wednesday).date()}'])
        else:
                      miercoles_1 = pd.DataFrame()   
                      
                      
        if fila.Thursday != None and not es_feriado(pd.to_datetime(fila.Thursday).date()):
                                curva_historica_jueves = df.loc[(df['usuario'] == usuario) & (df['fecha_hora'].dt.date == pd.to_datetime(fila.Thursday).date()), 'demanda'].values
                                jueves_1 = pd.DataFrame( curva_historica_jueves, columns=[f'Semana_{pd.to_datetime(fila.Thursday).isocalendar()[1]}_{pd.to_datetime(fila.Thursday).date()}'])
        else:
                      jueves_1 = pd.DataFrame()   
                
        if fila.Friday != None and not es_feriado(pd.to_datetime(fila.Friday).date()):
                                curva_historica_viernes = df.loc[(df['usuario'] == usuario) & (df['fecha_hora'].dt.date == pd.to_datetime(fila.Friday).date()), 'demanda'].values
                                viernes_1 = pd.DataFrame( curva_historica_viernes, columns=[f'Semana_{pd.to_datetime(fila.Friday).isocalendar()[1]}_{pd.to_datetime(fila.Friday).date()}'])
        else:
                     viernes_1 = pd.DataFrame()    
                          
        curvas_historicas = pd.concat([curvas_historicas,viernes_1, jueves_1, miercoles_1, martes_1], axis=1)       
    return curvas_historicas

def calculo_desviaciones(fecha, curva_analisis, curvas_historicas):

    dia_semana = fecha.strftime('%A')
    desviaciones = pd.DataFrame()
    
    if dia_semana in ['Tuesday', 'Wednesday', 'Thursday', 'Friday']:
        for col, fecha_1 in zip(curvas_historicas.columns, list(curvas_historicas.columns)):                     
            desviacion = curva_analisis - curvas_historicas[col].values                  
            desviaciones = pd.concat([desviaciones, pd.DataFrame(desviacion, columns=[fecha_1])], axis=1)
            desviacion = None
            
    if dia_semana in ['Monday', 'Saturday', 'Sunday']:
        for col, fecha_1 in zip(curvas_historicas.columns, fechas_dia_no_tipico(fecha)):              
                desviacion = curva_analisis - curvas_historicas[col].values               
                desviaciones = pd.concat([desviaciones, pd.DataFrame(desviacion, columns=[fecha_1.strftime('%Y-%m-%d')])], axis=1)
                desviacion = None
        
    if es_feriado(fecha):
        for col, fecha_1 in zip(curvas_historicas.columns, (fechas_domingos(fecha))):
                desviacion = curva_analisis - curvas_historicas[col].values 
                desviaciones = pd.concat([desviaciones, pd.DataFrame(desviacion, columns=[fecha_1.strftime('%Y-%m-%d')])], axis=1)
                desviacion = None
    return desviaciones
        
def comparacion_desviaciones(fecha, usuario, usuario_3, curvas_analisis, curvas_historicas, desviaciones, cont_1, hora_1):       
      
     tabla_datos_totales = pd.concat([pd.DataFrame({'usuario': [usuario for i in range(48)]}), pd.DataFrame({'hora': hora_1}), curvas_analisis, curvas_historicas], axis=1)

     df_dic = pd.DataFrame()
     tabla_datos_desviaciones = pd.DataFrame()

     df_dic = pd.concat([pd.DataFrame({'fecha': [fecha for i in range(48)]}), pd.DataFrame({'hora': hora_1}), pd.DataFrame({'usuario': [usuario for i in range(48)]}), desviaciones], axis=1)
     
     #tabla_datos_desviaciones  = pd.concat([tabla_datos_desviaciones ,df_dic], axis=1)
     
     return tabla_datos_totales, tabla_datos_desviaciones    


def ejecucion_demanda_parte_1(fecha, df, usuario, curva_analisis, curvas_analisis, cont_1, hora_1, usuario_3, tabla_datos_desviaciones_final, tabla_datos_totales_final):
    dia_semana_analisis = fecha.strftime('%A')
    curvas_historicas = curvas_hist(fecha, df, usuario)
    
    if dia_semana_analisis in ['Tuesday', 'Wednesday', 'Thursday', 'Friday'] and not es_feriado(fecha):
        curvas_historicas =curvas_hist_dia_tipico(fecha, df, usuario, curvas_historicas)
    #desviaciones = calculo_desviaciones(fecha, curva_analisis, curvas_historicas) 
    desviaciones = pd.DataFrame()
    tabla_datos_totales, tabla_datos_desviaciones  = comparacion_desviaciones(fecha, usuario, usuario_3, curvas_analisis, curvas_historicas, desviaciones, cont_1, hora_1)
    
    tabla_datos_totales_final = pd.concat([tabla_datos_totales_final, tabla_datos_totales], ignore_index=True) 
    tabla_datos_desviaciones_final =  pd.concat([tabla_datos_desviaciones_final, tabla_datos_desviaciones], ignore_index=True)
    
    
    return tabla_datos_desviaciones_final, tabla_datos_totales_final, cont_1


    
def ejecucion_demanda_parte_2(fecha, df_curva_historicas_dem_a, df_curva_analisis_dem_a, usuarios, usuario_3, num):
    
    cont_1 = 0
    hora_1 = []
    
    #hora_1 = [ timedelta(minutes=0) + (i + 1) * timedelta(minutes=30) - timedelta(minutes=1) if i == 47 else timedelta(minutes=0) + (i + 1) * timedelta(minutes=30) for i in range(48)]
    for index in range(48):
        a = timedelta(minutes=0)       
        if index == 47: a = timedelta(minutes=1) 
        hora = fecha + (index + 1) * timedelta(minutes=30) - a                  
        hora_1.append(hora.time().strftime('%H:%M'))
    """        
    ventana = tk.Tk()
    ventana.title("Ejecutando")

    marco = tk.Frame(ventana)
    marco.pack(padx=20, pady=20)

    etiqueta = tk.Label(marco, text="Progreso:")
    etiqueta.pack()

    barra_progreso = ttk.Progressbar(marco, orient='horizontal', length=300, mode='determinate')
    barra_progreso.pack(pady=10)
    """
    usuario_atipico, data_atipico, df_curva_analisis = aproximacion_polinimica_curva_analisis(fecha, hora_1, usuarios, df_curva_analisis_dem_a)
    
    
    for usuario in usuarios['usuarios']:
            if cont_1 == 0:
               tabla_datos_desviaciones_final = pd.DataFrame()
               tabla_datos_totales_final = pd.DataFrame()
               
            
            curvas_analisis = pd.DataFrame()                        
            curva_analisis = df_curva_analisis_dem_a.loc[(df_curva_analisis_dem_a['usuario'] == usuario) & (df_curva_analisis_dem_a['fecha_hora'].dt.date == fecha.date()) , 'demanda'].values                         
            curvas_analisis = pd.concat([curvas_analisis, pd.DataFrame(curva_analisis, columns=[fecha.strftime('%Y-%m-%d')])], axis=1)
                                
                            
            dia_semana_analisis = fecha.strftime('%A')
                            
            if es_feriado(fecha):
               tabla_datos_desviaciones_final, tabla_datos_totales_final, cont_1 = ejecucion_demanda_parte_1(fecha, df_curva_historicas_dem_a, usuario, curva_analisis, curvas_analisis, cont_1, hora_1, usuario_3, tabla_datos_desviaciones_final, tabla_datos_totales_final)
               
               cont_1 = cont_1 + 1
               """
               progreso = (cont_1 / num) * 100
               barra_progreso['value'] = progreso
               ventana.update() 
               """

               print(cont_1) 
               
            if dia_semana_analisis in ['Monday', 'Saturday', 'Sunday'] and not es_feriado(fecha):
               tabla_datos_desviaciones_final, tabla_datos_totales_final, cont_1 = ejecucion_demanda_parte_1(fecha, df_curva_historicas_dem_a, usuario, curva_analisis, curvas_analisis, cont_1, hora_1, usuario_3, tabla_datos_desviaciones_final, tabla_datos_totales_final)
               
               cont_1 = cont_1 + 1 
               """
               progreso = (cont_1 / num) * 100
               barra_progreso['value'] = progreso
               ventana.update() 
               """

               print(cont_1)
               
            if dia_semana_analisis in ['Tuesday', 'Wednesday', 'Thursday', 'Friday'] and not es_feriado(fecha):                
               tabla_datos_desviaciones_final, tabla_datos_totales_final, cont_1 = ejecucion_demanda_parte_1(fecha, df_curva_historicas_dem_a, usuario, curva_analisis, curvas_analisis, cont_1, hora_1, usuario_3, tabla_datos_desviaciones_final, tabla_datos_totales_final)
               
               cont_1 = cont_1 + 1
               """
               progreso = (cont_1 / num) * 100
               barra_progreso['value'] = progreso
               ventana.update() 
               """
               
               print(cont_1)
            
    
    #ventana.destroy()
    datos_frecuencia_desviaciones = pd.DataFrame()
    #datos_frecuencia_desviaciones = pd.concat([datos_frecuencia_desviaciones, tabla_datos_desviaciones_final['usuario'], tabla_datos_desviaciones_final['hora'], tabla_datos_desviaciones_final['num']], axis=1)
    #datos_frecuencia_desviaciones = pd.concat([datos_frecuencia_desviaciones, tabla_datos_desviaciones_final['usuario'], tabla_datos_desviaciones_final['hora']], axis=1)
    
    #ventana.mainloop()
    return tabla_datos_desviaciones_final, tabla_datos_totales_final, datos_frecuencia_desviaciones, usuario_atipico, data_atipico, df_curva_analisis



#=====================================================   DASH   =================================================


def tab_tablas_dem_a(tab_tabla, tabla_formato_encabezado, tabla_formato_celdas, df_grafico_histograma, df_tabla_desv, df_tabla_total_datos):
    tabla_1_value = df_tabla_desv.to_dict("records")
    tabla_2_value = df_tabla_total_datos.to_dict("records")
    

    #tab grafico de tabla de datos y desviaciones
    if tab_tabla == 'tabla_1_value_dem_activa':
            columna_excluir = ['usuario', 'fecha', 'num']
            columna_excluir = [col for col in df_tabla_desv.columns if col not in columna_excluir]
            tabla_desviaciones = html.Div(
                   [html.Div(
                           [ html.Div( [html.H6("DESVIACIONES (MW)", className="graph__title")] ),
                             DataTable(id="tabla_1_value_dem_activa",
                                       data = tabla_1_value, 
                                       columns=[{'name': col, 'id': col} for col in columna_excluir],
                                       page_size=16,
                                       style_header = tabla_formato_encabezado,
                                       style_data = tabla_formato_celdas,
                                       fixed_rows={'headers': True},
                                       page_action='none',
                                       style_table={'overflowX': 'auto'},
                                       style_cell={ 'textAlign': 'center',  'padding': '10px', 'minWidth': '0px', 'maxWidth': '0px', 'width': '0px'  },
                                       style_data_conditional=[ { 'if': {'column_id': df_tabla_desv.columns[1]}, 'position': 'sticky', 'left': 0, 'backgroundColor': 'white', 'zIndex': 1 }],
                                       ),
                             ],
                             className="one-third column wind__speed__container app__content_1") ],
                             className="app__content"),
            

            return tabla_desviaciones
        
    elif tab_tabla == 'tabla_2_value_dem_activa':
            columna_excluir = ['usuario']
            columna_excluir = [col for col in df_tabla_total_datos.columns if col not in columna_excluir]
            tabla_datos = html.Div(
                  [html.Div(
                          [ html.Div( [html.H6("DATOS (MW)", className="graph__title")] ),
                            DataTable(id="tabla_2_value_dem_activa",
                                      data = tabla_2_value,
                                      columns=[{'name': col, 'id': col} for col in columna_excluir],                                      
                                      page_size=16,
                                      style_header = tabla_formato_encabezado,
                                      style_data = tabla_formato_celdas,
                                      fixed_rows={'headers': True},
                                      page_action='none',
                                      style_table={'overflowX': 'auto'},
                                      style_cell={'textAlign': 'center', 'padding': '10px', 'minWidth': '0px', 'maxWidth': '0px',  'width': '0px'},
                                      style_data_conditional=[{ 'if': {'column_id': df_tabla_total_datos.columns[1]},'position': 'sticky', 'left': 0, 'backgroundColor': 'white', 'zIndex': 1 }],
                                      ),
                            ],    
                            className="one-third column wind__speed__container app__content_1")],
                            className="app__content"
                            ),

            return tabla_datos
    

def tab_graficos_dem_a(tab_linea, app_color,tabla_formato_encabezado, tabla_formato_celdas,  df_filtrado_dem, df_tabla_total_datos, nombre_seleccionado_dem_a):

    
     if tab_linea == 'grafico-dem-activa-por-usuario':
         figura_dem_a_usuario = go.Figure()
         for idx, columna in enumerate(df_tabla_total_datos.columns):
           if idx > 2: 

              line_style = dict(width=0.8)  # Estilo predeterminado para otras gráficas
              figura_dem_a_usuario.add_trace(go.Scatter( x=df_tabla_total_datos['hora'], y=pd.to_numeric(df_tabla_total_datos[columna], errors='coerce'),
              mode='lines+markers', name=columna, line=line_style ))
              
         for idx, columna in enumerate(df_tabla_total_datos.columns):
           if idx == 2: 
              line_style = dict(color='black', width=1.5)  # Estilo negro punteado y grosor 2
              figura_dem_a_usuario.add_trace(go.Scatter( x=df_tabla_total_datos['hora'], y=pd.to_numeric(df_tabla_total_datos[columna], errors='coerce'),
              mode='lines+markers', name=columna, line=line_style,  legendrank=0 ))
         
      
         figura_dem_a_usuario.update_layout(
          xaxis_title='Hora', yaxis_title='MW', legend_title='Dias', paper_bgcolor='#082255',
          font=dict(color='white', size=15), xaxis=dict(gridwidth=0.01, tickangle=90, zeroline=False), yaxis=dict(gridwidth=0.01,zeroline=False), height=500, margin=dict(t=20))
      

         
         grafico_linea = html.Div(
                         [ html.Div(
                           [html.Div([html.H6(nombre_seleccionado_dem_a, className="graph__title")]),
                            dcc.Graph(
                                    id="grafico-dem-activa-por-usuario",
                                    figure=figura_dem_a_usuario)],
                           className="one-third column wind__speed__container app__content_1" ), ],
                           className="app__content"
                           ),
       
 
         #grafico_tabla = tab_tablas_dem_a(tab_tabla, tabla_formato_encabezado, tabla_formato_celdas, df_grafico_histograma, df_tabla_desv, df_tabla_total_datos)
         tabla_2_value = df_tabla_total_datos.to_dict("records")
         columna_excluir = ['usuario']
         columna_excluir = [col for col in df_tabla_total_datos.columns if col not in columna_excluir]
         tabla_datos = html.Div(
               [html.Div(
                       [ html.Div( [html.H6("DATOS (MW)", className="graph__title")] ),
                         DataTable(id="tabla_2_value_dem_activa",
                                   data = tabla_2_value,
                                   columns=[{'name': col, 'id': col} for col in columna_excluir],                                      
                                   page_size=16,
                                   style_header = tabla_formato_encabezado,
                                   style_data = tabla_formato_celdas,
                                   fixed_rows={'headers': True},
                                   page_action='none',
                                   style_table={'overflowX': 'auto'},
                                   style_cell={'textAlign': 'center', 'padding': '10px', 'minWidth': '0px', 'maxWidth': '0px',  'width': '0px'},
                                   style_data_conditional=[{ 'if': {'column_id': df_tabla_total_datos.columns[1]},'position': 'sticky', 'left': 0, 'backgroundColor': 'white', 'zIndex': 1 }],
                                   ),
                         ],    
                         className="one-third column wind__speed__container app__content_1")],
                         className="app__content"
                         ),
         

         return  grafico_linea, tabla_datos
        
     elif tab_linea == 'grafico-dem-activa-total':
            figura_dem_a_total = go.Figure()
            for idx, col in enumerate(df_filtrado_dem.columns):
                if idx > 1:
                   line_style = dict(width=0.8)
                   figura_dem_a_total.add_trace(go.Scatter( x=df_filtrado_dem['hora'], y=pd.to_numeric(df_filtrado_dem[col], errors='coerce'),
                   mode='lines+markers', name=col, line=line_style))
                   
            for idx, col in enumerate(df_filtrado_dem.columns):
                if idx == 1:
                   line_style = dict(color='black', width=1.5)  # Estilo negro punteado y grosor 2
                   figura_dem_a_total.add_trace(go.Scatter( x=df_filtrado_dem['hora'], y=pd.to_numeric(df_filtrado_dem[col], errors='coerce'),
                   mode='lines+markers', name=col, line=line_style,  legendrank=0))

            figura_dem_a_total.update_layout(
                xaxis_title='Hora', yaxis_title='MW', legend_title='Dias', paper_bgcolor='#082255',
                font=dict(color='white', size=15), xaxis=dict(gridwidth=0.01, tickangle=90, zeroline=False), yaxis=dict(gridwidth=0.01,zeroline=False), height=500, margin=dict(t=20))


            grafico_linea = html.Div(
                   [ html.Div(
                           [html.Div([html.H6("DEMANDA TOTAL (MW)", className="graph__title")]),
                           dcc.Graph(
                                   id="grafico-dem-activa-total",
                                   figure=figura_dem_a_total)],
                           className="one-third column wind__speed__container app__content_1" ), ],
                           className="app__content",),

            
            #grafico_tabla = tab_tablas_dem_a(tab_tabla, tabla_formato_encabezado, tabla_formato_celdas, df_grafico_histograma, df_tabla_desv, df_tabla_total_datos)
            tabla_2_value = df_tabla_total_datos.to_dict("records")
            columna_excluir = ['usuario']
            columna_excluir = [col for col in df_tabla_total_datos.columns if col not in columna_excluir]
            tabla_datos = html.Div(
                  [html.Div(
                          [ html.Div( [html.H6("DATOS (MW)", className="graph__title")] ),
                            DataTable(id="tabla_2_value_dem_activa",
                                      data = tabla_2_value,
                                      columns=[{'name': col, 'id': col} for col in columna_excluir],                                      
                                      page_size=16,
                                      style_header = tabla_formato_encabezado,
                                      style_data = tabla_formato_celdas,
                                      fixed_rows={'headers': True},
                                      page_action='none',
                                      style_table={'overflowX': 'auto'},
                                      style_cell={'textAlign': 'center', 'padding': '10px', 'minWidth': '0px', 'maxWidth': '0px',  'width': '0px'},
                                      style_data_conditional=[{ 'if': {'column_id': df_tabla_total_datos.columns[1]},'position': 'sticky', 'left': 0, 'backgroundColor': 'white', 'zIndex': 1 }],
                                      ),
                            ],    
                            className="one-third column wind__speed__container app__content_1")],
                            className="app__content"
                            ),
            
            return  grafico_linea, tabla_datos
        

def Aplicativo(app, url, ejecucion_demanda_parte_2, dem_total):   

    app.title = "Alerta de Desviaciones de Demanda"
    app_color = {"graph_bg": "#082255", "graph_line": "#007ACE"}  
    tabla_formato_encabezado = {'text-align': 'center', 'border': '0.5px solid black', 'padding': '2px', 'backgroundColor': 'midnightblue',  'color': 'white', 'font-size': '15px'}
    tabla_formato_celdas = {'text-align': 'center', 'border': '0.5px solid black', 'padding': '2px', 'color': 'black', 'backgroundColor': 'gainsboro', 'font-size': '15px'}
    tab_boton_formato_seleccion = { 'borderTop': '1px solid #00ffff', 'borderBottom': '1px solid #00ffff', 'borderLeft': '1px solid #00ffff', 'borderRight': '1px solid #00ffff', 'text-align': 'center', 'backgroundColor': '#082255', 'color': '#00ffff', 'padding': '10px', 'justifyContent': 'center', 'display': 'flex', 'justifyContent': 'center', 'alignItems': 'center', 'height': '40px',}
    tab_boton_formato_sin_seleccion = { 'text-align': 'center', 'backgroundColor': '#082255', 'color': 'white', 'padding': '10px', 'justifyContent': 'center', 'borderRadius': '0px', 'marginBottom': '3px', 'display': 'flex', 'justifyContent': 'center', 'alignItems': 'center', 'height': '40px' }

    app.layout = html.Div([
        html.Div([
               html.Div(id='page-content'),]),
        
        html.Br(),
        html.Div(
               [html.Div(
                       [html.H4("APLICATIVO DE DETECCIÓN DE PUNTOS ATÍPICOS EN LA DEMANDA", className="app__header__title", style={'marginLeft': '10px'}),],
                       className="app__header__desc",),
                   html.Div(
                       [html.A(html.Img( src=app.get_asset_url("dash-new-logo.png"), className="app__menu__img", style={'marginTop': '25px', 'marginRight': '25px'}),
                       href="https://www.coes.org.pe/Portal/",),],
                       className="app__header__logo",),],
                       className="app__header",),
       #app__header__title
       dbc.Row([
       html.Div([
       html.H6("SUBIR INFORMACIÓN A BD:", className="app__header", style={'marginLeft': '10px','color':'white'}),
       ],
       className="app__header__desc"),]),
       
       dbc.Row([
       dbc.Col([
       dcc.Loading( id="loading-1", type="default",
       children = html.Div([ 
            dcc.Upload(
                id='upload-data',
                children=html.Div([
                    'Arrastar o ',
                    html.A('Seleccionar Archivos', style ={'color': '#00ffff', 'fontSize': '18px'})
                ]),
                style={
                    'width': '50%',
                    'height': '60px',
                    'lineHeight': '60px',
                    'borderWidth': '1px',
                    'borderRadius': '5px',
                    'textAlign': 'center',
                    'margin': '10px',
                    'border': '2px dashed #00ffff',
                    'color': 'white',
                    'display': 'flex',
                    'justifyContent': 'center',
                    'alignItems': 'center',
                    'marginLeft': '10px',
                    'fontSize': '16px'},
                multiple=True
            ),

            html.Div(id="loading-output-1")
            
        ]),),
       ]),
       dbc.Col([
           html.Div(id="mensaje-1")
           ]),
       ]),

  
       html.Br(),
       dbc.Row([
       html.Div([
       html.H6("EXTRAER INFORMACION DE CARPETA:", className="app__header", style={'marginLeft': '10px','color':'white'}),
       ],
       className="app__header__desc"),
       ]),
       
       dbc.Row([
       dbc.Col([
       dcc.Loading( id="loading-2", type="default",
       children = html.Div([ 
            dcc.Upload(
                id='upload-data-2',
                children=html.Div([
                    'Arrastar o ',
                    html.A('Seleccionar Archivos', style = {'fontSize': '16px', 'color': '#00ffff', 'fontWeight': 'bold'})
                ]),
                style={
                    'width': '50%',
                    'height': '60px',
                    'lineHeight': '60px',
                    'borderWidth': '1px',
                    'borderRadius': '5px',
                    'textAlign': 'center',
                    'margin': '10px',
                    'border': '2px dashed #00ffff',
                    'color': 'white',
                    'display': 'flex',
                    'justifyContent': 'center',
                    'alignItems': 'center',
                    'marginLeft': '10px',
                    'fontSize': '16px'
                    
                },
                multiple=True
            ),
            dcc.Store(id='informacion-1'),
            html.Div(id="loading-output-2")
            
        ]),),
       ]),
       dbc.Col([
           html.Div(id="mensaje-2")
           ]),
       ]),
       
       html.Br(),
       html.Br(),
       
       html.Div(
           [                    
            html.Div(           
                   [html.Div([
                           dcc.Dropdown( id='lista_desplegable_dem_activa',
                           
                           options= [], value= None,
                           style={ 'width': '300px', 'height': '40px', 'fontSize': '16px', 'marginTop': '20px',
                                       'borderRadius': '5px'},
                           clearable=False,
                           optionHeight=40,
                           
                           ),
                           html.Div([
                           html.Button('Retroceder', id='retroceder-boton', n_clicks=0, style={'marginTop': '20px','marginRight': '15px', 'color':'#00ffff', 'borderTop': '1px solid #00ffff', 'borderBottom': '1px solid #00ffff', 'borderLeft': '1px solid #00ffff', 'borderRight': '1px solid #00ffff'}),
                           html.Button('Continuar', id='continuar-boton', n_clicks=0, style={'marginTop': '20px', 'color':'#00ffff', 'borderTop': '1px solid #00ffff', 'borderBottom': '1px solid #00ffff', 'borderLeft': '1px solid #00ffff', 'borderRight': '1px solid #00ffff'}),
                           html.Div(id='output-container', style={'marginTop': '20px'}),],), 
                           ])]),]),
            
            html.Br(),
            dbc.Row([
            html.Div( 

                    html.Div([
                           dcc.Tabs(id="tabs_grafico_dem_activa", value='grafico-dem-activa-por-usuario',className='dash-tabs', 
                           children=[
                               dcc.Tab(label='Demanda Por Usuario (MW)', value='grafico-dem-activa-por-usuario',style = tab_boton_formato_sin_seleccion , selected_style=tab_boton_formato_seleccion),
                               dcc.Tab(label='Demanda Total (MW)', value='grafico-dem-activa-total',style = tab_boton_formato_sin_seleccion , selected_style=tab_boton_formato_seleccion),
                               ]),
                           html.Div(id='tabs_content_grafico_dem_activa')]),),]),
            
            dbc.Row([
            html.Br(),
            html.Div(id = 'tabla_2_value'),
            html.Br(),
            ],),
                          
       ],className="app__container",)
    
                          
       
    
    @app.callback(Output("loading-output-1", "children"),
              Output("mensaje-1", "children"),
              Input('upload-data', 'contents'),
              State('upload-data', 'filename'),)
    

    
    def cargar_1(list_of_contents, nombre_archivo):
        
        if list_of_contents is not None:
            df_total = pd.DataFrame()
            for contenido_1, archivo in zip(list_of_contents, nombre_archivo):
                if isinstance(contenido_1, str):
                    content_type, content_string = contenido_1.split(',')
                    decoded = base64.b64decode(content_string)
                    
                    if archivo.endswith('.raw'):
                        with io.StringIO(decoded.decode('utf-8')) as raw:
                            contenido = raw.readlines()
                           
                        #cont = cont + 1
                        ind_line = 0
                        ind_inicio = 0
                        ind_fin = 0
                            
                        for line in range(len(contenido)):
                                
                            if contenido[line] == ' 0 / END OF BUS DATA, BEGIN LOAD DATA\n':
                                ind_inicio = ind_line + 1 
                            if contenido[line] == ' 0 / END OF LOAD DATA, BEGIN FIXED BUS SHUNT DATA\n':
                                ind_fin = ind_line - 1
                                break
                            ind_line = ind_line + 1
                            
                        fecha = datetime(int(archivo[5:9]), int(archivo[9:11]), int(archivo[11:13]), int(archivo[13:15]), int(archivo[15:17]))
                           
                        lineas_a_leer = contenido[ind_inicio:ind_fin + 1] 
                        
                        nombre_demanda = [linea[87-1:95].strip() if linea[87-1:95].strip() else None for linea in lineas_a_leer]
                        activa_demanda = [linea[23-1:30].strip() if linea[23-1:30].strip() else None for linea in lineas_a_leer]
                        reactiva_demanda = [linea[32-1:39].strip() if linea[32-1:39].strip() else None for linea in lineas_a_leer]
                        
                        fecha_demanda = [fecha.strftime('%Y-%m-%d %H:%M') if fecha else None for linea in lineas_a_leer]       
                        

                        datos_dict = {           
                                    'fecha_hora': fecha_demanda,
                                    'Nombre': nombre_demanda,
                                    'Activa': activa_demanda,
                                    'Reactiva': reactiva_demanda}  
                        
                        df = pd.DataFrame(datos_dict)
                        df_total = pd.concat([df_total, df], ignore_index=True)
                        
                        df_total_activa = df_total 
                        df_total_activa['Activa'] = pd.to_numeric(df_total['Activa'], errors='coerce')  
                        tabla_1_activa = df_total_activa.pivot_table(index=["fecha_hora"], columns="Nombre", values="Activa", aggfunc="sum") 
                        
                        df = pd.DataFrame()
                        contenido.clear()
                        datos_dict = {}
                        fecha_1, lineas_a_leer = [], []
                    tabla_1_activa = tabla_1_activa.reset_index()
                    tabla_1_activa['fecha_hora'] = pd.to_datetime(tabla_1_activa['fecha_hora'])
                    tabla_1_activa = tabla_1_activa.sort_values(by='fecha_hora')
                    #tabla_1_activa = tabla_1_activa.set_index('fecha_hora')
                                  
                          
                    fecha_1 = datetime(int(archivo[5:9]), int(archivo[9:11]), int(archivo[11:13]))
            
            #tabla_1_activa = tabla_1_activa.sort_values(by='fecha_hora')
            #print(tabla_1_activa)
            fecha_1 = fecha_1.date()
            dia_semana = fecha_1.strftime('%A')
         
            if dia_semana in ['Tuesday']:
                hojas = ['martes']
            if dia_semana in ['Wednesday']:
                hojas = ['miercoles']
            if dia_semana in ['Thursday']:
                hojas = ['jueves']
            if dia_semana in ['Friday']:
                hojas = ['viernes']
                
            if dia_semana == 'Monday':
                hojas = ['lunes']
            if dia_semana == 'Saturday':
                hojas = ['sabado']
            if dia_semana == 'Sunday':
                hojas = ['domingo']
            
            #ruta_archivo = r'C:\Users\lportilla\Downloads\PSO\PROYECTO\Filtro_Por_DiaSemana\ACTIVA_DEM.xlsx'  
            directorio = os.getcwd()
            
            ruta_archivo = os.path.join(directorio, hojas[0] + '.csv')
            
            for hoja in hojas: 
                
                df_n = pd.read_csv(ruta_archivo)
                #df_n = df_n.reset_index()
                            
                df_n['fecha_hora'] = pd.to_datetime(df_n['fecha_hora'])
            
            if tabla_1_activa['fecha_hora'].isin(df_n['fecha_hora']).any():
                df_n = df_n[~df_n['fecha_hora'].isin(tabla_1_activa['fecha_hora'])]
                print('datos actualizados')
                print(tabla_1_activa)
                #return print('Cargado Exitosamente'), [html.A("¡Actualizado Exitosamente!", style = {'fontSize': '24px', 'color': '#00ffff', 'fontWeight': 'bold', 'textDecoration': 'none'})]
            
                
            datos_completos = pd.DataFrame()
            datos_completos = pd.concat([datos_completos, df_n, tabla_1_activa], ignore_index=True)
            
            datos_completos['fecha_hora'] = pd.to_datetime(datos_completos['fecha_hora'])

            datos_completos = datos_completos.sort_values(by='fecha_hora')
            """
            with pd.ExcelWriter(ruta_archivo, engine='openpyxl', mode='a') as writer:
                writer.book.remove(writer.book[hojas[0]])
                datos_completos.to_excel(writer, sheet_name=hojas[0], index=False)
            """
            datos_completos.to_csv(ruta_archivo, index=False)
           # print(datos_completos)
            
            return print('Cargado Exitosamente'), [html.A("¡Cargado Exitosamente!", style = {'fontSize': '24px', 'color': '#00ffff', 'fontWeight': 'bold', 'textDecoration': 'none'})]
        else:
            return print('Error'), None   
            
                    


        
        
    @app.callback(
              Output('informacion-1', 'data', allow_duplicate=True),
              Output("mensaje-2", "children"),
              Output("loading-output-2", "children"),
              Input('upload-data-2', 'contents'),
              State('upload-data-2', 'filename'),
              )            
                
    
    def extraer(list_of_contents, nombre_archivo):
        if list_of_contents is not None:
            df_total = pd.DataFrame(columns=['Fecha', 'Nombre', 'Codigo', 'Activa', 'Reactiva'])
            cont = 0
            
            for contenido_1, archivo in zip(list_of_contents, nombre_archivo):
                if isinstance(contenido_1, str):
                    content_type, content_string = contenido_1.split(',')
                    decoded = base64.b64decode(content_string)
                    
                    if archivo.endswith('.raw'):
                        with io.StringIO(decoded.decode('utf-8')) as raw:
                            contenido = raw.readlines()
                           
                        cont = cont + 1
                        ind_line = 0
                        ind_inicio = 0
                        ind_fin = 0
                            
                        for line in range(len(contenido)):
                                
                            if contenido[line] == ' 0 / END OF BUS DATA, BEGIN LOAD DATA\n':
                                ind_inicio = ind_line + 1 
                            if contenido[line] == ' 0 / END OF LOAD DATA, BEGIN FIXED BUS SHUNT DATA\n':
                                ind_fin = ind_line - 1
                                break
                            ind_line = ind_line + 1
                            
                        fecha = datetime(int(archivo[5:9]), int(archivo[9:11]), int(archivo[11:13]), int(archivo[13:15]), int(archivo[15:17]))
                           
                        lineas_a_leer = contenido[ind_inicio:ind_fin + 1] 
                        
                        nombre_demanda = [linea[87-1:95].strip() if linea[87-1:95].strip() else None for linea in lineas_a_leer]
                        activa_demanda = [linea[23-1:30].strip() if linea[23-1:30].strip() else None for linea in lineas_a_leer]
                        reactiva_demanda = [linea[32-1:39].strip() if linea[32-1:39].strip() else None for linea in lineas_a_leer]
                        
                        fecha_demanda = [fecha.strftime('%Y-%m-%d %H:%M') if fecha else None for linea in lineas_a_leer]       
                        

                        datos_dict = {           
                                    'fecha_hora': fecha_demanda,
                                    'Nombre': nombre_demanda,
                                    'Activa': activa_demanda,
                                    'Reactiva': reactiva_demanda}  
                        
                        df = pd.DataFrame(datos_dict)
                        df_total = pd.concat([df_total, df], ignore_index=True)
                        
                        df_total_activa = df_total 
                        df_total_activa['Activa'] = pd.to_numeric(df_total['Activa'], errors='coerce')  
                        tabla_activa = df_total_activa.pivot_table(index=["fecha_hora"], columns="Nombre", values="Activa", aggfunc="sum") 
                        
                        df = pd.DataFrame()
                        contenido.clear()
                        datos_dict = {}
                        fecha_1, lineas_a_leer = [], []
                    tabla_activa = tabla_activa.reset_index()
                    tabla_activa['fecha_hora'] = pd.to_datetime(tabla_activa['fecha_hora'])
                    tabla_activa = tabla_activa.sort_values(by='fecha_hora')
                    tabla_activa = tabla_activa.set_index('fecha_hora')
                    tabla_1_activa = tabla_activa.melt(ignore_index=False, var_name='usuario', value_name='demanda') 
                    tabla_1_activa = tabla_1_activa.reset_index()
                    
                    
                    
                    tabla_1_activa['fecha_hora'] = pd.to_datetime(tabla_1_activa['fecha_hora'], errors='coerce')
                    tabla_1_activa['fecha'] = tabla_1_activa['fecha_hora'].dt.date 
                    tabla_1_activa['hora'] = tabla_1_activa['fecha_hora'].dt.time  
                          
                    fecha_1 = datetime(int(archivo[5:9]), int(archivo[9:11]), int(archivo[11:13]))

            df_curva_historicas_dem_a   = cargar_datos_excel(fecha_1)
            
            directorio = os.getcwd()
            
            usuarios   = pd.read_csv(os.path.join(directorio, 'usuarios.csv'))    
            usuarios_filtrados = [usuario for usuario in tabla_activa.columns.tolist() if not usuarios['usuarios'].isin([usuario]).any()]
            usuarios_filtrados = pd.DataFrame(usuarios_filtrados, columns=['usuarios'])    
              

                    
            contar_usuarios = len(usuarios_filtrados)           
            usuario_3_a = list()

            fecha_1 = pd.to_datetime(fecha_1)

            tabla_datos_desviaciones_final_dem_a, tabla_datos_totales_final_dem_a, datos_frecuencia_desviaciones_dem_a, usuario_atipico, data_atipico, df_curva_analisis = ejecucion_demanda_parte_2(fecha_1, df_curva_historicas_dem_a, tabla_1_activa, usuarios_filtrados, usuario_3_a, contar_usuarios)

            tabla_dem_total_a= dem_total(tabla_datos_totales_final_dem_a)
            tabla_dem_total_a = tabla_dem_total_a[tabla_datos_totales_final_dem_a.drop(columns=['usuario']).columns]
            option_usuario_3_a = [{'label': nombre, 'value': nombre} for nombre in  (usuario_atipico)]    


            return {
                    #'tabla_datos_desviaciones_final_dem_a': tabla_datos_desviaciones_final_dem_a.to_dict(),
                    'tabla_datos_totales_final_dem_a': tabla_datos_totales_final_dem_a.to_dict(),
                    #'datos_frecuencia_desviaciones_dem_a': datos_frecuencia_desviaciones_dem_a.to_dict(),
                    'usuario_atipico': usuario_atipico, 
                    'data_atipico': data_atipico.to_dict(), 
                    'df_curva_analisis': df_curva_analisis.to_dict(), 
                    'tabla_dem_total_a': tabla_dem_total_a.to_dict(),
                    'option_usuario_3_a': option_usuario_3_a
                    } , [html.A("¡Ha terminado de compilar!", style = {'fontSize': '24px', 'color': '#00ffff', 'fontWeight': 'bold', 'textDecoration': 'none'})], None
        else:
                return no_update, no_update, no_update 
                            
  

    @app.callback(
        Output('lista_desplegable_dem_activa', 'options', allow_duplicate=True),
        Output('informacion-1', 'data', allow_duplicate=True),
        Input('informacion-1', 'data'),
    )
     
    def data_2(data):
        
        if data is not None:     
            return data['option_usuario_3_a'], data
        
        else:
            return no_update, no_update
        
    @app.callback(
    Output('informacion-1', 'data', allow_duplicate=True),
    Input('lista_desplegable_dem_activa', 'value'),
    Input('informacion-1', 'data')
    )
    def actualizar_output(valor, data):
        if data is not None:
            return data
        else:
            return no_update
            



    @app.callback(
    Output('informacion-1', 'data', allow_duplicate=True),
    Output('lista_desplegable_dem_activa', 'value'),
    Output('continuar-boton', 'n_clicks'),
    Output('retroceder-boton', 'n_clicks'),
    Input('continuar-boton', 'n_clicks'),
    Input('retroceder-boton', 'n_clicks'),
    Input('lista_desplegable_dem_activa', 'value'),
    State('informacion-1', 'data'))
    
    def update_drop(avanzar, anterior, valor, data):
        if avanzar != 0 or anterior !=0:
            if avanzar !=0:
                pos_init= data['usuario_atipico'].index(valor)
                pos_final=pos_init+1
     
                if pos_final == len(data['usuario_atipico']):
                    pos_final=pos_init
                valor_final=data['usuario_atipico'][pos_final]
     
                return data, valor_final, 0, 0
     
            else:
                pos_init= data['usuario_atipico'].index(valor)
                pos_final=pos_init-1
     
                if pos_final == -1:
                    pos_final=pos_init
                valor_final=data['usuario_atipico'][pos_final]
     
                return data, valor_final, 0, 0        
        else:
            return data, valor, 0, 0



    @app.callback(
                  [Output('tabs_content_grafico_dem_activa', 'children'),
                  Output('tabla_2_value', 'children')],
                  
                  [Input('lista_desplegable_dem_activa', 'value'),
                  State('informacion-1', 'data')],
                  [Input('tabs_grafico_dem_activa', 'value'),
                  #Input('tabs_tabla_dem_activa', 'value'),
                  ]
                  )

    def actualizar_dashboard_dem_a(nombre_seleccionado_dem_a, data, tab_linea_dem_a,):
        
         if data is not None:

             tabla_datos_totales_final_dem_a = pd.DataFrame(data['tabla_datos_totales_final_dem_a'])
             #datos_frecuencia_desviaciones_dem_a = pd.DataFrame(data['datos_frecuencia_desviaciones_dem_a'])
             #tabla_datos_desviaciones_final_dem_a = pd.DataFrame(data['tabla_datos_desviaciones_final_dem_a'])
             tabla_dem_total_a = pd.DataFrame(data['tabla_dem_total_a'])
            
             df_tabla_total_datos_dem_a = tabla_datos_totales_final_dem_a[tabla_datos_totales_final_dem_a['usuario'] == nombre_seleccionado_dem_a]
             #df_grafico_histograma_dem_a = datos_frecuencia_desviaciones_dem_a[datos_frecuencia_desviaciones_dem_a['usuario'] == nombre_seleccionado_dem_a]
             #df_tabla_desv_dem_a = tabla_datos_desviaciones_final_dem_a[tabla_datos_desviaciones_final_dem_a['usuario'] == nombre_seleccionado_dem_a]
             df_dem_total_a = tabla_dem_total_a
                     
             for col in df_tabla_total_datos_dem_a.columns:
                 if col not in ['usuario', 'hora']:
                     df_tabla_total_datos_dem_a[col] = pd.to_numeric(df_tabla_total_datos_dem_a[col], errors='coerce').map("{:.2f}".format)
                
             #for col in df_tabla_desv_dem_a.columns:
             #    if col not in ['fecha', 'usuario', 'hora', 'num']:
             #        df_tabla_desv_dem_a[col] = pd.to_numeric(df_tabla_desv_dem_a[col], errors='coerce').map("{:.2f}".format)
                     
             for col in df_dem_total_a.columns:
                 if col not in ['hora']:
                     df_dem_total_a[col] = pd.to_numeric(df_dem_total_a[col], errors='coerce').map("{:.2f}".format)
            
             
             grafico_linea_dem_a, grafico_tabla_dem_a = tab_graficos_dem_a(tab_linea_dem_a, app_color,tabla_formato_encabezado, tabla_formato_celdas,  df_dem_total_a, df_tabla_total_datos_dem_a, nombre_seleccionado_dem_a)
           

             return grafico_linea_dem_a, grafico_tabla_dem_a
         
         else:
             return no_update, no_update
         
url = 'http://127.0.0.1:8050'
#http://10.0.5.36:8054/

    
app = dash.Dash(  __name__, meta_tags=[{"name": "viewport", "content": "width=device-width, initial-scale=1"}], prevent_initial_callbacks='initial_duplicate',  suppress_callback_exceptions=True)
server = app.server

Aplicativo(app, url, ejecucion_demanda_parte_2, dem_total)

if __name__ == "__main__":
    app.run_server(host="10.0.5.36",port = 8054 ,debug=True)







