#!/usr/bin/env python
# -*- coding: utf-8 -*-

###MODIFICADO Por JeanKashmir

#Aun esta en proceso de Modificacion
#18-01-2022

"""
PyFingerprint
Copyright (C) 2015 Bastian Raschke <bastian.raschke@posteo.de>
All rights reserved.

"""
import os
import hashlib
import uno
import sys
import time
from colorama import init, Fore
from tqdm import tqdm
from pyfingerprint.pyfingerprint import PyFingerprint
init()
os.system("/usr/lib/libreoffice/program/soffice.bin --headless --invisible --nocrashreport --nodefault --nofirststartwizard --nologo --norestore --accept='socket,Host=localhost,port=2002,tcpNoDelay=1;urp;StarOffice.ComponentContext'") 
  ##AGREGAR CARGA
os.system('clear')

print('\033[1m' + Fore.YELLOW + "1.- Enrolar nuevo Usuario" + '\033[0m')
print('\033[1m' + Fore.GREEN + "2.- Registrar Llegada / Salida" + '\033[0m' )
print('\033[1m' + Fore.GREEN +"3.- Verificar usuario" + '\033[0m' )
print('\033[1m' + Fore.RED + "4.- Eliminar Usuario" + '\033[0m' )
print('\033[1m' + Fore.CYAN + "5.- Obtener Datos de Usuario" + '\033[0m')
opcion=int(input('\033[1m' + Fore.BLUE + "Escoja una opcion: " + '\033[0m'))

os.system('clear')

if opcion== 1 :
    
    nombre= str(input('\033[1m' + Fore.YELLOW + "Ingrese nombre(s) del colaborador: " + '\033[0m' ))
    apellido= str(input('\033[1m' + Fore.YELLOW + 'Ingrese apellido(s) del colaborador: ' + '\033[0m' ))
    identificador= str(input('\033[1m' + Fore.YELLOW + 'Ingrese RUN/DNI/Pasaporte del colaborador: ' + '\033[0m' ))
    os.system('clear')
    print ('\033[1m' + Fore.YELLOW + "Cargando..."+ '\033[0m' )
    time.sleep(1)
    ## Registrar huella
    
  
    try:
        f = PyFingerprint('/dev/ttyUSB0', 57600, 0xFFFFFFFF, 0x00000000)

        if ( f.verifyPassword() == False ):
            raise ValueError('Ups! Existe un error')

    except Exception as e:
        print('Ups! El error no inicio correctamente')
        print('ID: ' + str(e))
        exit(1)

    try:
        print('\033[1m' + Fore.GREEN +'Ingrese Huella...' + '\033[0m')

     
        while ( f.readImage() == False ):
            pass

    
        f.convertImage(0x01)

        result = f.searchTemplate()
        positionNumber = result[0]

        if ( positionNumber >= 0 ):
            print('La huella ya existe. Se encuentra en la posicion' + str(positionNumber))
            exit(0)

        print('\033[1m' + Fore.GREEN +'Saque el dedo' + '\033[0m')
        time.sleep(2)

        print('\033[1m' + Fore.GREEN +'Vuelva a colocar la huella'  + '\033[0m')

        while ( f.readImage() == False ):
            pass

        f.convertImage(0x02)

      
        if ( f.compareCharacteristics() == 0 ):
            raise Exception('Ups! no se pudo sincronizar')

       
        f.createTemplate()

  
        positionNumber = f.storeTemplate()
        print('\033[1m' + Fore.GREEN +'Enrolado exitoso' + '\033[0m')

    except Exception as e:
        print('Fallo encontrado')
        print('ID: ' + str(e))
        exit(1)




elif opcion== 2:
    
    try:
        f = PyFingerprint('/dev/ttyUSB0', 57600, 0xFFFFFFFF, 0x00000000)

        if ( f.verifyPassword() == False ):
            raise ValueError('Ups! Hubo un problema')

    except Exception as e:
        print('Error encontrado!')
        print('ID: ' + str(e))
        exit(1)

    print('\033[1m' + Fore.YELLOW + 'Sistema de Control de Acceso: PARRONALES DE NOS 2' + '\033[0m')

    try:
        print('\033[1m' + Fore.BLUE + 'Ingrese su huella...' +'\033[0m')

        while ( f.readImage() == False ):
            pass


        f.convertImage(0x01)

        result = f.searchTemplate()

        positionNumber = result[0]
        accuracyScore = result[1]

        if ( positionNumber == -1 ):
            print('\033[1m' + Fore.RED + 'Usuario no identificado' + '\033[0m')
            exit(0)
        else:
            print('\033[1m' + Fore.GREEN + 'Verificando identidad...' + '\033[0m')
      


        f.loadTemplate(positionNumber, 0x01)

        characterics = str(f.downloadCharacteristics(0x01)).encode('utf-8')


    
        excel=hashlib.sha256(characterics).hexdigest()
    
        localContext= uno.getComponentContext()
        resolver= localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
        ctx= resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        smgr = ctx.ServiceManager

        #Obteniendo Titulo de Excel abierto
        desktop = smgr.createInstanceWithContext ("com.sun.star.frame.Desktop", ctx)
        document= desktop.getCurrentComponent()
        document.getTitle()

        #Planilla
        sheets= document.getSheets()
        sheets.getByIndex(0)

        #Seleccionando Celda
        sheets.getByIndex(0).getCellRangeByName("J4")
        sheets.getByIndex(0).getCellRangeByName("J4").setString(excel)
    
        time.sleep(1)
        os.system('clear')
        print('\033[1m' + Fore.GREEN + 'Usuario registrado correctamente' + '\033[0m')
    


    except Exception as e:
        print('Ups! Error hubo un error')
        print('ID: ' + str(e))
        exit(1)
