#DHL Team
'''
import asyncio
from pyppeteer import launch
import time

async def main():

    browser = await launch(headless=False)
    page = await browser.newPage()
    await page.setViewport({'width': 1200, 'height': 800, 'deviceScaleFactor': 1})
    page.setDefaultNavigationTimeout(70000)
    await page.goto('https://dsctmsr2.dhl.com/GC3/glog.webserver.servlet.umt.Login?ct=MTU1MTEwNjc5NDE0NzAyMzM2OQ%3D%3D')
    await page.waitFor('body > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > form > table > tbody > tr:nth-child(2) > td:nth-child(2) > input[type=text]')
    await page.waitFor('body > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > form > table > tbody > tr:nth-child(3) > td:nth-child(2) > input[type=password]')
    username = await page.J('body > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > form > table > tbody > tr:nth-child(2) > td:nth-child(2) > input[type=text]')  #Hace objeto la etiqueta de campo para poner usuario
    password = await page.J('body > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > form > table > tbody > tr:nth-child(3) > td:nth-child(2) > input[type=password]')  #Hace objeto la etiqueta para colocar contraseña
    await username.type('MXPLAN.YBARROSOJ')    #se rellena
    await password.type('Yen2414...')
    await page.waitFor('body > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > form > table > tbody > tr:nth-child(5) > td:nth-child(2) > input')
    await page.waitFor(6000)
    await page.click('body > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td > form > table > tbody > tr:nth-child(5) > td:nth-child(2) > input')   #se hace click en el elemento
    await page.goto('https://dsctmsr2.dhl.com/GC3/glog.webserver.finder.FinderServlet?ct=MTU1MTEwNjc5NDE0NzAyMzM2OQ%3D%3D&query_name=glog.server.query.shipment.BuyShipmentQuery&finder_set_gid=MXPLAN.MX%20ST%20BUY%20SHIPMENT%20CENTRAL%20BC')
    await page.waitFor('#bodyDataDiv > table > tbody > tr > td:nth-child(1) > table > tbody > tr:nth-child(1) > td > div > input[type=text]:nth-child(3)')
    CeUve=await page.J('#bodyDataDiv > table > tbody > tr > td:nth-child(1) > table > tbody > tr:nth-child(1) > td > div > input[type=text]:nth-child(3)')
    await CeUve.type('01564113')
    await page.waitFor('#search_button')
    await page.click('#search_button')
    await page.waitFor('#rgSGSec\\.2')
    await page.waitFor('#rgSGSec\\.1\\.1\\.1\\.1\\.check')
    await page.click('#rgSGSec\\.1\\.1\\.1\\.1\\.check')
    await page.waitFor('#editButton')
    await page.click('#editButton')
    popI = len(browser.targets())
    print(popI)
    await page.waitFor(2000)
    pop = browser.targets()[2]
    popup= await pop.page()
    print(popup.url)
    print(len(popup.frames))
    frame = popup.frames[0]
    print(type(frame))
    await frame.waitFor("#ShipmentFinancials")   
    await frame.click("#ShipmentFinancials")
    await frame.waitFor('#\\31 3502054 > td.gridBodyCell.gridBodyBtnsCell > table > tbody > tr > td:nth-child(1) > a > img')
    await frame.click('#\\31 3502054 > td.gridBodyCell.gridBodyBtnsCell > table > tbody > tr > td:nth-child(1) > a > img', clickCount=3)


asyncio.get_event_loop().run_until_complete(main())

'''
import win32com.client as win32
import datetime
import time
import pandas as pd
from csv import writer
import csv
import easygui
import schedule
import os.path
import math
from glob import glob

path='//mxmex1-fipr01/public$/Nave 1/LPC/audit/'            #direccion donde se guarda el csv creado
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
account= win32.Dispatch("Outlook.Application").Session.Accounts
arrDest=[]
arrLin=[]

def confirmacion(fieldValues, valoresPick):
    msg = 'CV: '+fieldValues[0]+'\n'+'Cliente: '+valoresPick[1]+'\n'+'Destino: '+fieldValues[2]+'\n'+'Origen: '+valoresPick[2]+'\n'+'Tipo de Unidad: '+valoresPick[0]+'\n'+'Tarifa: '+fieldValues[2]+'\n'+'Linea: '+fieldValues[3]+'\n'+'Numero de referencia: '+fieldValues[4]+'\n'+'Correo: ' +fieldValues[5]
    title = "Prompt de confrmación"
    if easygui.ccbox(msg, title):     # show a Continue/Cancel dialog
        pass  # user chose Continue
        return fieldValues
    else:  # user chose Cancel
        retBucle=bucleConfirmacion(fieldValues,valoresPick)
        return retBucle
        
def tipodeUnidad():
    msg ="Selecciona el tipo de unidad"
    title = "Tipo de unidad"
    choicesT = ["FULL",'THNC','53PSO','THNM','53T','THN','48TC','THNB','48TP','TN6C','48T','48TB','THNP','48M','TN6','48PSO','MOTO','MDZ','15TR','TN6P','TN6M','RBN','RBNB','RBNC','RBNM','RBNP','10T','12T','15C','THNR','48TR','RBNR','35TP','35TR','35T','TEST','30T','15TB','15T','35TC','35TB']
    choiceT = easygui.choicebox(msg, title, choicesT)
    if choiceT==None:     # show a Continue/Cancel dialog
            MensajeBienv()
    else:  # user chose Cancel
        if easygui.ccbox("Escogiste: {}".format(choiceT)):
            return choiceT
        else:
            tipodeUnidad()
            return choiceT

def cliente():
    msg ="Selecciona el cliente"
    title = "Cliente"
    choicesC = ['MXPLAN/MXAAG','MXPLAN/MXATT','MXPLAN/MXABT','MXPLAN/MXACC','MXPLAN/MXACH','MXPLAN/MXALA', 'MXPLAN/MXAME','MXPLAN/MXAPP','MXPLAN/MXASE','MXPLAN/MXATT','MXPLAN/MXBAR','MXPLAN/MXBEA','MXPLAN/MXBIO','MXPLAN/MXCAD','MXPLAN/MXCAL','MXPLAN/MXCAM','MXPLAN/MXCIB','MXPLAN/MXCOK','MXPLAN/MXCON','MXPLAN/MXCOP','MXPLAN/MXCOR','MXPLAN/MXCOT','MXPLAN/MXCUE','MXPLAN/MXDAR','MXPLAN/MXDEV','MXPLAN/MXEFC','MXPLAN/MXEFL','MXPLAN/MXEFM','MXPLAN/MXEFZ','MXPLAN/MXENG','MXPLAN/MXESA','MXPLAN/MXEVE','MXPLAN/MXFOD','MXPLAN/MXGAT','MXPLAN/MXGLX','MXPLAN/MXGMI','MXPLAN/MXHAM','MXPLAN/MXHBS','MXPLAN/MXHER','MXPLAN/MXHPC','MXPLAN/MXHPE','MXPLAN/MXIDV','MXPLAN/MXIMD','MXPLAN/MXIMP','MXPLAN/MXIMX','MXPLAN/MXIMX','MXPLAN/MXISS','MXPLAN/MXJAN','MXPLAN/MXJCT','MXPLAN/MXJNJ','MXPLAN/MXKAZ','MXPLAN/MXKEL','MXPLAN/MXKFT','MXPLAN/MXKIT','MXPLAN/MXLEG','MXPLAN/MXLEN','MXPLAN/MXLGE','MXPLAN/MXMAP','MXPLAN/MXMAT','MXPLAN/MXMAV','MXPLAN/MXMER','MXPLAN/MXMOK','MXPLAN/MXNOV','MXPLAN/MXNXT','MXPLAN/MXOGL','MXPLAN/MXPIL','MXPLAN/MXPIL','MXPLAN/MXPLAN','MXPLAN/MXPNG','MXPLAN/MXPSO','MXPLAN/MXRYC','MXPLAN/MXSBS','MXPLAN/MXSDK','MXPLAN/MXSHW','MXPLAN/MXSOL','MXPLAN/MXSON','MXPLAN/MXSUN','MXPLAN/MXSWM','MXPLAN/MXTAJ','MXPLAN/MXTCT','MXPLAN/MXTCY','MXPLAN/MXTEV','MXPLAN/MXTRN','MXPLAN/MXTYC','MXPLAN/MXUHU','MXPLAN/MXUNI','MXPLAN/MXWKT','MXPLAN/MXZUU']
    choiceC = easygui.choicebox(msg, title, choicesC)
    if choiceC==None:     # show a Continue/Cancel dialog
            tipodeUnidad()
    else:  # user chose Cancel
        conf=easygui.ccbox("Escogiste: {}".format(choiceC))
        if conf:
            return choiceC
        else:
            cliente()
            return choiceC

def origen():
    msg ="Selecciona el origen"
    title = "Origen"
    choicesO = ['37D002','37D004','37D005','37D009','37D014','37D015','37D016','37D019','37D024','37D027','37D031','37D035','37D037','37D050','37D051','37D065','37D079','37D082','37D090','37D091','37D092','37D100','37D110','37D112','37D115','37D116','37D117','37D118','37D120','37D121','37D122','37D123','37D124','37D128','37D129','37D130','37D131','37D132','37D133','37D138','37D139','37D140','37D142','37D143','37D145','37D146','37D147','37D151']
    choiceO = easygui.choicebox(msg, title, choicesO)
    if choiceO==None:     # show a Continue/Cancel dialog
            cliente()
    else:  # user chose Cancel
        conf=easygui.ccbox("Escogiste: {}".format(choiceO))
        if conf:
            return choiceO
        else:
            origen()
            return choiceO

def start():
    tipo=tipodeUnidad()
    cli=cliente()
    ori=origen()
    #'cliente','origen','tipo de unidad'
    msg = "Echa las papas al fuego"
    title = "Control de correos"
    fieldNames = ["CV","Destino","Tarifa",'Linea','Numero de referencia','Correo','CC','CC2 (no es necesario)','CC3(no es necesario)']
    fieldValues = easygui.multenterbox(msg, title, fieldNames)
    while fieldValues==None:
        fieldValues = easygui.multenterbox(msg, title, fieldNames)
    parseD=fieldValues[1]
    parseL=fieldValues[3]
    for i in parseD:
        arrDest.append(i)
    for i in parseL:
        arrLin.append(i)
    while (arrDest[0]!='M'or arrDest[1]!='X'):
        errmsg = "{}: Usa formato MX/#.\n\n".format(fieldNames[1])
        fieldValues = easygui.multenterbox(errmsg, title, fieldNames, fieldValues)
        parseDNew=fieldValues[1]
        for i in range(len(arrDest)):
            arrDest.pop(0)
        for i in parseDNew:
            arrDest.append(i)
    while (arrLin[0]!='M'or arrLin[1]!='X'):
        errmsg = "{}: Usa formato MX/#.\n\n".format(fieldNames[3])
        fieldValues = easygui.multenterbox(errmsg, title, fieldNames, fieldValues)
        parseLNew=fieldValues[3]
        for i in range(len(arrLin)):
            arrLin.pop(0)
        for i in parseLNew:
            arrLin.append(i)
    cont=0
    while 1:
        errmsg = ""
        for i, name in enumerate(fieldNames):
            if cont<=5:
                if fieldValues[i].strip() == "":
                    errmsg += "{} is a required field.\n\n".format(name)
                cont+=1

        if errmsg == "":
            break # no problems found
        fieldValues = easygui.multenterbox(errmsg, title, fieldNames, fieldValues)
        if fieldValues is None:
            break
    valoresPick=[tipo,cli,ori]
    datosConf=confirmacion(fieldValues, valoresPick)
    fieldValues.insert(1,cli)
    fieldValues.insert(3,ori)
    fieldValues.insert(4,tipo)
    return datosConf                                                 #Return del arreglo


def bucleConfirmacion(fValues, valoresPick):
    msg = "Echa las papas al fuego"
    title = "Control de correos"
    fieldNames = ["CV","Destino","Tarifa",'Linea','Numero de referencia','Correo','CC','CC2 (no es necesario)','CC3 (no es necesario)']
    fieldValues = easygui.multenterbox(msg, title, fieldNames,fValues)
    cont=0
    while 1:
        errmsg = ""
        for i, name in enumerate(fieldNames):
            if cont<=5:
                if fieldValues[i].strip() == "":
                    errmsg += "{} is a required field.\n\n".format(name)
                cont+=1
        if errmsg == "":
            break # no problems found
        fieldValues = easygui.multenterbox(errmsg, title, fieldNames, fieldValues)
        if fieldValues is None:
            break
    print("Reply was:{}".format(fieldValues))
    confirmacion(fieldValues, valoresPick)
    return fieldValues

def succesfullSend():                      #Funcion que manda el mail. Recibe la dirección de correo, datos necesarios p/registro, num de referencia y el nombre del archivo a crear (csv)
    outlook = win32.Dispatch('outlook.application')
    baseDatos = pd.read_csv(fname)
    with open(fname,"r") as f:
        reader = csv.reader(f,delimiter = ",")
        data = list(reader)
        row_count = len(data)                   #Determina el numero de filas
    for i in range(row_count-1):
        cv=baseDatos.values[i][1]
        cliente=baseDatos.values[i][2]
        Destino=baseDatos.values[i][3]
        Origen=baseDatos.values[i][4]
        TipodeU=baseDatos.values[i][5]
        tarifa=baseDatos.values[i][6]
        linea=baseDatos.values[i][7]
        email_data = [cv,cliente,Destino,Origen,TipodeU,tarifa,linea]
 
        mail = outlook.CreateItem(0) 
        cadena_destinatarios = str(baseDatos.at[i,'Correo'])
        if len(str(baseDatos.at[i,'CC2']))>7:
            cadena_destinatarios += ';' + str(baseDatos.at[i, 'CC2'])
        if len(str(baseDatos.at[i,'CC1']))>7:
            cadena_destinatarios += ';' + str(baseDatos.at[i, 'CC1'])
        if len(str(baseDatos.at[i,'CC3']))>7:
            cadena_destinatarios += ';' + str(baseDatos.at[i, 'CC3'])

        mail.To = cadena_destinatarios
        mail.VotingOptions = "Accept;Decline"                       #Activate voting options
        mail.Subject = 'Autorización spot: '+str(baseDatos.at[i,'Referencia'])                 #Num de referencia se agrega al asunto del correo
        body = txt_to_str("mail_format.html")
        for i in range(len(email_data)):
            body = body.replace("<!--SPLITCONTROL("+str(i)+")-->","<td>{0}</td>".format(str(email_data[i])))        

        mail.HTMLBody = body
        images_path = "C:\\Users\\aperalda\\Documents\\RAudit\\RateAudit\\img\\"
        
        camino = os.path.realpath(__file__)
        print("EL CAMINO------------------: ", camino)
        mail.Attachments.Add(Source= images_path+"voting_buttons.png")
        mail.Attachments.Add(Source= images_path+"llpc.PNG")                                 #Los datos se agregan al cuerpo del correo
        mail.Send()                                                 #Se manda
        print("Successful send")                                    #Wuuuuu

def txt_to_str(route):
    f = open(route, mode="r", encoding="utf-8")
    content = f.read()
    f.close()
    return str(content)

def valid_email():
    return True
                                                 
def Envio():
    baseDatos = pd.read_csv(fname)
    baseDatos.at['Hora de envio']=datetime.datetime.now()

def succesfullRetrieval():                                 #Funcion que clasifica aceptados y rechazados
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)                         # "6" refers to the index of a folder - in this case the inbox.                                      
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    for message in messages:                                        #Hace un barrido por cada correo en bandeja
        subject_content = message.Subject
        subjectSplit=subject_content.split()                      #Separa el asunto del correo por espacios
        with open(fname ,"r") as f:                                 #Se abre el archivo en cuestión, en forma de lectura para capturar el numero de filas existentes
            reader = csv.reader(f,delimiter = ",")
            data = list(reader)
            row_count = len(data)                                   #Lee el numero de filas
        baseDatos = pd.read_csv(fname)                         #Se define al archivo en forma de lectura, para poder obtener datos
        if (len(subjectSplit)==4):                                  #Se valida que el asunto del correo tenga dos palabras  
            if (subjectSplit[0]=="Accept:"):                        #Condicion que determina si la tarifa fue aceptada
                horaAcept=datetime.datetime.now()                   #Hora en que se captura el correo
                csvstr = datetime.datetime.strftime(horaAcept, '%Y, %m, %d, %H, %M, %S')
                arrsinsp=csvstr.split()
                strv=''
                for i in range(len(arrsinsp)):
                    strv+=arrsinsp[i]
                arrcsv=strv.split(',')
                strcsv=''
                for i in range(len(arrcsv)):
                    if i<=1:
                        strcsv+=arrcsv[i]+'-'
                    elif i==2:
                        strcsv+=arrcsv[i]+'/'
                    elif i>2 and i<=4:
                        strcsv+=arrcsv[i]+':'
                    else:
                        strcsv+=arrcsv[i]
                respuestaPor=str(message.Sender)        
                respSplit=respuestaPor.split()
                respStr=''
                for i in range (0,3):
                    respStr+=respSplit[i]+' '
                for i in range(row_count-1):                        #Ciclo que barre la base de datos csv en busca de la tarifa encontrada en el correo
                    if (baseDatos['Referencia'][i]==int(subjectSplit[3])):           #Se busca el match de la referencia del correo con alguno de la base de datos
                        if(baseDatos['Respuesta'][i]=='Aceptado' or baseDatos['Respuesta'][i]=='Rechazado'):
                            pass
                        else:
                            baseDatos.at[i,'Hora de respuesta']=strcsv               #Se escribe la hora de respuesta
                            baseDatos.at[i,'Respuesta']='Aceptado' 
                            baseDatos.at[i,'Persona que responde']=respStr                    #Se cambia el valor de la columna 'Respuesta' de la referecia en cuestión a 'Aceptado'
                            baseDatos.to_csv(fname, index=False)                         #Se guarda el archivo 
                        

            if (subjectSplit[0]=="Decline:"):                       #Mismo proceso pero para rechazados
                horaDecl=datetime.datetime.now()
                csvstr = datetime.datetime.strftime(horaDecl, '%Y, %m, %d, %H, %M, %S')
                arrsinsp=csvstr.split()
                strv=''
                for i in range(len(arrsinsp)):
                    strv+=arrsinsp[i]
                arrcsv=strv.split(',')
                strcsv=''
                for i in range(len(arrcsv)):
                    if i<=1:
                        strcsv+=arrcsv[i]+'-'
                    elif i==2:
                        strcsv+=arrcsv[i]+'/'
                    elif i>2 and i<=4:
                        strcsv+=arrcsv[i]+':'
                    else:
                        strcsv+=arrcsv[i]
                respuestaPor=str(message.Sender)
                respSplit=respuestaPor.split()
                respStr=''
                for i in range (0,3):
                    respStr+=respSplit[i]+' '
                for i in range(row_count-1):
                    if (baseDatos['Referencia'][i]==int(subjectSplit[3])):
                        if(baseDatos['Respuesta'][i]=='Rechazado' or baseDatos['Respuesta'][i]=='Aceptado'):
                            pass
                        else:
                            baseDatos.at[i,'Hora de respuesta']=strcsv
                            baseDatos.at[i,'Respuesta']='Rechazado'
                            baseDatos.at[i,'Persona que responde']=respStr
                            baseDatos.to_csv(fname, index=False)


def creaCSV(datos):                               #Funcion que crea la base de datos csv
    #if (datos[0]=='' or datos[1]=='' or datos[2]=='' or datos[3]=='' or datos[4]=='' or datos[5]=='' or datos[6]=='' or datos[7]=='' or datos[8]=='' or datos[9]==''):
        #break
    mydate=datetime.datetime.now()                                    #Dato que indica la hora a la que se mandó el correo                             
    csvstr = datetime.datetime.strftime(mydate, '%Y, %m, %d, %H, %M, %S')
    arrsinsp=csvstr.split()
    strv=''
    for i in range(len(arrsinsp)):
        strv+=arrsinsp[i]
    arrcsv=strv.split(',')
    strcsv=''
    for i in range(len(arrcsv)):
        if i<=1:
            strcsv+=arrcsv[i]+'-'
        elif i==2:
            strcsv+=arrcsv[i]+'/'
        elif i>2 and i<=4:
            strcsv+=arrcsv[i]+':'
        else:
            strcsv+=arrcsv[i]
    cv=datos[0]
    cliente=datos[1]
    destino=datos[2]
    origen=datos[3]
    tipodeu=datos[4]
    tarifa=datos[5]
    linea=datos[6]
    referencia=datos[7]
    correo=datos[8]
    if len(datos)<10:
        correo1=""
        correo2=""
        correo3=""
    elif len(datos)<11:
        correo1=datos[9]
        correo2=""
        correo3=""
    elif len(datos)<12:
        correo1=datos[9]
        correo2=datos[10]
        correo3=""
    else:
        correo1=datos[9]
        correo2=datos[10]
        correo3=datos[11]

    addToInit=[referencia,cv, cliente,destino,origen,tipodeu,tarifa,linea,strcsv,'-','-',correo,correo1,correo2,correo3,'-']   #Se agregan los valores obtenidos al csv con la columna de respuesta con valor '-'
    # Se abre el archivo en modo append
    with open(fname, 'a+', newline='') as write_obj:
        csv_writer = writer(write_obj)  
        csv_writer.writerow(addToInit)  # Se añade el contenido a la última fila del archivo

    

def correosAceptadosPorTiempo():           #Funcion que acepta correos por tiempo excedido
    baseDatos = pd.read_csv(fname)
    with open(fname,"r") as f:
        reader = csv.reader(f,delimiter = ",")
        data = list(reader)
        row_count = len(data)                   #Determina el numero de filas
    for i in range(row_count-1):               #Barrido de datos en la base de datos en la columna 'Hora de envio'
        if (((pd.to_datetime(baseDatos['Hora de envio'][i])+datetime.timedelta(minutes=2))<datetime.datetime.now())):      #Si la hora de envio excedió 30 mins,
            baseDatos.at[i,'Hora de respuesta']=datetime.datetime.now()                                                           
            baseDatos.at[i,'Respuesta']='Aceptado por tiempo'                                                               #Se modifica la columna 'Respuesta' a 'Aceptado por tiempo'
            baseDatos.to_csv(fname, index=False)                                                                            #Se guarda el archivo

    
def nombreArchivo():                                           #Funcion que crea el nombre del archivo
    fileName=datetime.datetime.now()                           #El nombre va en funcion de la hora y fecha que se crea el archivo
    fileNameS=str(fileName)                                    #La fecha se convierte a tipo string
    splitCol=fileNameS.split(':')                              #Se separa por puntos, dado que es un caracter no valido en el nombre de los archivos
    splitSpc=splitCol[0].split()                               #El primer elemento se separa para obtener fecha y los primeros dos dígitos de la hora
    acFileName=splitSpc[0]+'-'+splitSpc[1]+'-'+splitCol[1]+'.csv' #Se agrega fecha, hora, minuto y 'csv' para darle dicha extensión
    fullPath=path+acFileName                                      #El nombre completo del archivo es la dirección en la que se quiere guardar + el nombre creado
    with open(fullPath, 'a+',newline='') as csvfile:              #Se crean los headers del archivo
        fieldnames = ['Referencia', 'CV','Cliente','Destino','Origen','Tipo de Unidad','Tarifa','Linea','Hora de envio','Respuesta','Hora de respuesta','Correo','CC1','CC2','CC3','Persona que responde']
        csv_writter = writer(csvfile)  
        csv_writter.writerow(fieldnames)  # Se añaden los datos a la ultima fila del archivo (en este caso la primera)
    return fullPath                       #Se regresa el nombre del archivo


def MensajeBienv():
    Bienvenido=easygui.ccbox(msg='Bienvenido al sistema DHL', title='DHL Supply Chain', choices=('Agregar', 'Enviar'), image='img\\dhl_logo.png',default_choice='Agregar', cancel_choice='Enviar')
    if  Bienvenido==True:
            datos=start()                                                   #Se piden los datos
            creaCSV(datos)
            MensajeBienv()
    elif Bienvenido==False:
        succesfullSend()                                         #Cuando exista valor en la variable fileName, se toma el tiempo actual
        #horaDeEnvio()

txt_to_str("mail_format.html")
fname=nombreArchivo()                                                #Se crea el nombre del archivo
MensajeBienv()                           

tiempoControl=datetime.datetime.now()                               #Para determinar que el proceso de enviar ha terminado
while True:
    succesfullRetrieval()                                       #Se leen los correos en bandeja
    correosAceptadosPorTiempo()                                 #Se modifican los que exceden el tiempo de respuesta
    if (tiempoControl+datetime.timedelta(minutes=35)<datetime.datetime.now()):  #Se cuentan 40 minutos a partir de que terminó el proceso de enviar correos                                                    
        break                                                              
