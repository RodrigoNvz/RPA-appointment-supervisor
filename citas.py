''' Automate appointments prototype 2.0
    Authors:
        Jesus Heriberto Vasquez Sanchez
        Jose Rodrigo Narvaez Berlanga'''

import asyncio, os, time, pyppeteer, pandas, csv
import numpy,calendar,schedule, datetime,shutil,win32com.client as win32
import os, sys
from pyppeteer import launch
from datetime import datetime,timedelta
from subprocess import STDOUT, check_output

master_citas = [] # all appointments
outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
account= win32.Dispatch("Outlook.Application").Session.Accounts

expiredAccounts = []

def launchQlik(route, name, retries = 1):
    now = datetime.now()
    cmd = r'"C:\Program Files\QlikView\Qv.exe" /r ' + route
    for i in range(retries):
        try:
            output = check_output(cmd, stderr=STDOUT, timeout = 1200)
            print('Generado',now.strftime("%Y-%m-%d %H:%M"),': ', name)
            return 1
        except:
            print('Timeout',now.strftime("%Y-%m-%d %H:%M"),': ', name)
    return 0 

#Walmart appointment extraction method
async def wm_appointment_portal(user,passwd,account_name):
    #browser = await launch(headless=False)
    browser = await launch()
    strusername = "body > div > div > div > div.page-container > div.main-container > div.content-container > div > div > form > span:nth-child(1) > span > span > input"
    strpass = "body > div > div > div > div.page-container > div.main-container > div.content-container > div > div > form > span:nth-child(2) > span > span > input"
    strbtn = "body > div > div > div > div.page-container > div.main-container > div.content-container > div > div > form > button"
    entregasEncontradas="#mc >  tbody > tr:nth-child(2) > td.contentPanel > table > tbody > tr:nth-child(4) > td > form > table > tbody > tr > td > table > tbody > tr.contentBodyRow > td.contentBody > table.formTable > tbody > tr:nth-child(2) > td.valueTd"

    page = await browser.newPage()
    await page.setViewport({"width": 1024, "height": 768, "deviceScaleFactor": 1})
    page.setDefaultNavigationTimeout(30000) 
    await page.goto("https://retaillink.login.wal-mart.com/?ServerType=IIS1&CTAuthMode=BASIC&language=en&utm_source=retaillink&utm_medium=redirect&utm_campaign=FalconRelease&CT_ORIG_URL=/&ct_orig_uri=/")
    await page.waitFor(strusername)

    username = await page.querySelector(strusername)
    password = await page.querySelector(strpass)

    await username.type(user)
    await password.type(passwd)
    await page.click(strbtn)
    try: 
        await page.waitForNavigation()
    except:
        print("Failed in",account_name,"login.")
        expiredAccounts.append(account_name)
        await browser.close()
        return 0
    
    # Opening query session
    testpage = await browser.newPage()
    await testpage.goto("https://retaillink.wal-mart.com/navis/default.aspx")
    await testpage.waitForNavigation()
    await testpage.waitForNavigation()
    await testpage.waitForNavigation()

    # Opening this week Deliveries
    await testpage.goto("https://logistics-scheduler-www9.wal-mart.com/trips_mx/quickQuery.do?type=thisWeeksDeliveries")
    xtabla = await testpage.waitForXPath('//*[@id="SortTable0"]/tbody')
    
    # Extracting table size
    tabla = await testpage.evaluate("(xtabla) => xtabla.children", xtabla)
    
    #Use try to manage when there are no near appointments 
    noAppointFound= await testpage.querySelector(entregasEncontradas)
    noAppointments=await testpage.evaluate('(noAppointFound) => noAppointFound.textContent',noAppointFound)
    if noAppointments=='0':
        await page.waitFor(5000)
        await browser.close()
    else:
        for i in range(1, len(tabla) + 1):
            index = str(i)
            x1 = await testpage.waitForXPath('//*[@id="SortTable0"]/tbody/tr[{}]/td[1]/a'.format(index))
            no_entrega = await testpage.evaluate("(x1) => x1.innerText", x1)
            x2 = await testpage.waitForXPath('//*[@id="SortTable0"]/tbody/tr[{}]/td[8]'.format(index))
            cita = await testpage.evaluate("(x2) => x2.innerText", x2)
            clean_cita = datetime.strptime(cita,'%m/%d/%y %I:%M %p')#.strftime('%m/%d/%Y %I:%M:%S %p')
            master_citas.append([no_entrega, clean_cita])
        
        print("Succesful",account_name,"extraction")
        await page.waitFor(5000)
        try:
            await browser.close()
        except:
            print("Error al cerrar, same as always")
        return master_citas

#-----------------------------------------------------------------------------------------------------
#Freko-portal
async def fsk_appointment_portal(user,passwd,account_name):
    browser = await launch(headless=False)
    strusr="body > table > tbody > tr:nth-child(5) > td > table > tbody > tr > td:nth-child(4) > form > table > tbody > tr:nth-child(2) > td:nth-child(2) > input"
    strpass="body > table > tbody > tr:nth-child(5) > td > table > tbody > tr > td:nth-child(4) > form > table > tbody > tr:nth-child(3) > td:nth-child(2) > input"
    enterbtn="body > table > tbody > tr:nth-child(5) > td > table > tbody > tr > td:nth-child(4) > form > table > tbody > tr:nth-child(4) > td > input[type=IMAGE]"
    citasProgramadas="miTabla1 > tbody > tr:nth-child(4) > td.menuSub > a"
    
    page = await browser.newPage()
    await page.setViewport({"width": 1024, "height": 768, "deviceScaleFactor": 1})
    page.setDefaultNavigationTimeout(30000) # maybe 60000
    
    await page.goto("http://www.provecomer.com.mx/htmlProvecomer/provecomer.html")
    await page.waitFor(strusr)
    await page.waitFor(strpass)
    username = await page.querySelector(strusr)
    password = await page.querySelector(strpass)
    await username.type(user)
    await password.type(passwd)
    await page.click(enterbtn) 

    # try: 
    #     await page.waitForNavigation()
    #     print("Succesful login...Navigating")
    # except:
    #     print("Failed in",account_name,"login.")
    #     await browser.close()
    #     return 0

    txt=await page.content()
    frames=page.frames
    frame01=frames[2] #this is the one for miTabla1
    
    #frame02=frames[4]
    #txt2=await frame01.content() #Uncomment soon
    #print(txt2)

    await frame01.waitFor("[id='miTabla1']")
    await frame01.click("[id='miTabla1']")

    #frametest=await frames[1].content()
    #print(frametest)
    #awaited=await frames[0].childFrames[0].content()

    citasProg=await frame01.waitForXPath('//*[@id="miTabla1"]/tbody/tr[4]/td[2]/a')
    await citasProg.click()
    ref='GeneraReporteFrm > table > tbody > tr:nth-child(12) > td:nth-child(8)'
    #Now extract appointments
    frame02=frames[3]
    txt3= await frame02.content()

    xtabla = await frame02.waitForXPath('//*[@id="GeneraReporteFrm"]/table/tbody')
    # Extracting table size
    tabla = await frame02.evaluate("(xtabla) => xtabla.children", xtabla)

    #ref1=await frame02.waitForXPath('//*[@id="GeneraReporteFrm"]/table/tbody/tr[12]/td[8]')
    #refF=await frame02.evaluate("(ref1) => ref1.innerText", ref1)
    #print(refF)

    #print(len(tabla))
    # if noAppointments=='0':
    #     print("NO",account_name,"APPOINTMENTS IN PORTAL")
    #     await page.waitFor(5000)
    #     await browser.close()
    # else:
    master_temporal=[]
    #make the initial range dinamic, make it start when it does not find it.
    for i in range(1, len(tabla) + 1):
        index = str(i)
        try:
            x1 = await frame02.waitForXPath('//*[@id="GeneraReporteFrm"]/table/tbody/tr[{}]/td[8]'.format(index))
            refCita = await frame02.evaluate("(x1) => x1.innerText", x1)
            x2 = await frame02.waitForXPath('//*[@id="GeneraReporteFrm"]/table/tbody/tr[{}]/td[13]'.format(index))
            refFecha= await frame02.evaluate("(x2) => x2.innerText",x2)            
            if(refCita !='Num. Ref.'):
            #clean_cita = datetime.strptime(refFecha,'%m/%d/%y %I:%M %p')
                master_temporal.append([refCita,refFecha])
                #master_citas.append()
        except:
            print('xpath does not exists')
        
        # if x1!=' ':
        #     refI= await frame02.evaluate("(x1) => x1.innerText", x1)
        #     #x2 = await testpage.waitForXPath('//*[@id="SortTable0"]/tbody/tr[{}]/td[8]'.format(index))
        #     #cita = await testpage.evaluate("(x2) => x2.innerText", x2)
        #     # Conversión a datetime
        #     #clean_cita = datetime.strptime(cita,'%m/%d/%y %I:%M %p')#.strftime('%m/%d/%Y %I:%M:%S %p')
        #     print(refI)
        #     master_temporal.append(refI)
    print(master_temporal)
    
    await page.waitFor(6000)

#-----------------------------------------------------------------------------------------------------
#Validate appointments in OTM 
async def captureOTM(arrCR,arrLate):
    # browser = await launch({'args': ['--disable-dev-shm-usage']})
    #browser = await launch(headless=False)  # headless false means open the browser in the operation
    browser = await launch()
    page = await browser.newPage()
    await page.setViewport({"width": 1024, "height": 768, "deviceScaleFactor": 1})
    page.setDefaultNavigationTimeout(60000)
    await page.goto("https://dsctmsr2.dhl.com/GC3/glog.webserver.servlet.umt.Login")

    data=readFile(r"appointData.txt", "txt")
    passw=await page.waitFor("[name='userpassword']")
    usernN=await page.waitFor("[name='username']")
    await passw.type(data[1])
    await usernN.type(data[0])       
    await page.click("[name='submitbutton']")  

    returned=[]
    for i in range(len(arrCR)):
        valuei=[]
        await page.goto("https://dsctmsr2.dhl.com/GC3/glog.webserver.finder.FinderServlet?ct=NzY5Nzg2NjExNDQwNjgzNTIyMg%3D%3D&query_name=glog.server.query.order.OrderReleaseQuery&finder_set_gid=MXCORP.MX%20OM%20ORDER%20RELEASE")
        await page.waitFor(1000)
        await page.waitFor("[name='order_release/xid']")  # Wait for the order release
        # await page.waitFor("[name='orrOrderReleaseRefnumValue59']") #Wait for the CR
        #await page.type("[name='order_release/xid']", oR[i])
        await page.type("[name='orrOrderReleaseRefnumValue59']",arrCR[i])
        await page.keyboard.press("Enter")

        await page.waitFor("[tabindex='201']")
        confirm = await page.querySelector("[id='rgSGSec.2.2.1.22']")
        folio = await page.evaluate('(confirm)=> confirm.textContent', confirm)
        early = await page.querySelector("[id='rgSGSec.2.2.1.18']")
        eDate = await page.evaluate('(early)=> early.textContent', early)
        late = await page.querySelector("[id='rgSGSec.2.2.1.29']")
        lDate = await page.evaluate('(late)=> late.textContent', late)
        consigna = await page.querySelector("[id='rgSGSec.2.2.1.24']")
        consignatario = await page.evaluate('(consigna)=> consigna.textContent', consigna)
        #containedTxt=content.split(" ")
        #if(len(containedTxt)>=2):
            #print(containedTxt[1])
        
        # valuei.append(folio)
        # valuei.append(eDate)
        # valuei.append(lDate)
        # valuei.append(consignatario)
        # returned.append(valuei)

        await page.waitFor("[name='rgSGSec.1.1.1.1.check']")
        await page.waitFor(1000)
        await page.click("[name='rgSGSec.1.1.1.1.check']")
        await page.waitFor("[title='Mass Update']")
        await page.waitFor(1000)
        await page.click("[title='Mass Update']")
        frames=page.frames
        temp= len(frames)
        while temp < 3 : #Wait until the frame is loaded
            temp= len(frames)
        frame = page.frames[3]

        await frame.waitFor("[name='order_release/delivery_is_appt']")
        await frame.waitFor("[name='order_release/early_delivery_date']")
        await frame.waitFor("[name='order_release/late_delivery_date']") #Wait for the order release)
        await frame.waitFor("[name='order_release/ship_with_group']")

        # checked= await frame.querySelector("[name='order_release/delivery_is_appt']")
        # buttonstatus=await(await checked.getProperty('checked')).jsonValue()
        # if buttonstatus==True:
        #     await frame.type("[name='order_release/late_delivery_date']",lateDate)
        # else:

        await frame.type("[name='order_release/ship_with_group']",arrCR[i])
        #await frame.type("[name='order_release/early_delivery_date']",firstDate)
        await frame.type("[name='order_release/late_delivery_date']",arrLate[i])
        await frame.click("[name='order_release/delivery_is_appt']")
        await frame.waitFor(1000)

    '''return returned
    await page.waitFor(3000)
    await browser.close()'''

#-----------------------------------------------------------------------------------------------------
#Method to send Email
def sendEmail(address,body,subject):
    o = win32.Dispatch("Outlook.Application")
    oacctouse = None
    for oacc in o.Session.Accounts:
        if oacc.SmtpAddress == "rpa.transport_@dhl.com":
            oacctouse = oacc
            break
    Msg = o.CreateItem(0)
    if oacctouse:
        Msg._oleobj_.Invoke(*(64209, 0, 8, 0, oacctouse))  # Msg.SendUsingAccount = oacctouse

    Msg.To = address
    Msg.HTMLBody = body
    Msg.Subject= subject
    images_path = "C:\\Users\\jesushev\\Documents\\LaunchScripts\\Ramses\\"
    Msg.Attachments.Add(Source= images_path+"DHL.png")    
    Msg.Send()

#-----------------------------------------------------------------------------------------------------
# Here we do the verification between the walmart site and OTM, consolidating data
def verificacionCita():
    print("Cargando Qlikview...")
    launchQlik(r'C:\Users\jesushev\Documents\QV\citas.qvw', 'Prime Light', 3)

    #extract all apointments on master all citas
    now = datetime.now()
    enviarCorreo = "No"
    anySinLateDelivery = 'No'
    anyLateDeliveryDiferente = 'No'
      
    # with open(r'S:\TRANSPORTE\LPC\ApptUser\USUARIO.csv') as credentials:
    #     gen_reader = csv.reader(credentials, delimiter = ',')
    #     next(gen_reader, None) #Skips headers
    #     for row in gen_reader:
    #         account_name = row[0]
    #         user = row[1]
    #         password = row[2]
    #         try:
    #             asyncio.get_event_loop().run_until_complete(wm_appointment_portal(user,password,account_name))
    #         except:
    #             print('Timeout',now.strftime("%Y-%m-%d %H:%M"),": ","on account:",account_name)
    #asyncio.get_event_loop().run_until_complete(wm_appointment_portal("yuriria.perez@dhl.com","Mayonesa2020","MONDELEZ"))
    asyncio.get_event_loop().run_until_complete(wm_appointment_portal("c7cs1bg","GL2021mx","LENOVO"))
    asyncio.get_event_loop().run_until_complete(wm_appointment_portal("4oy63w5 ","Julio*2020","OTHER"))
    light=lightReading(r'S:\TRANSPORTE\LPC\TEMP\Beto\Prime_Light.csv')#Now read prime light
    clienteDestinoCSV = csv.reader(open(r'S:\TRANSPORTE\LPC\ApptUser\CLIENTE DESTINO.csv'))

    clienteDestinoDict = {} #Fill Cliente Destino values in dictionary
    for row in clienteDestinoCSV:
        key = row[0]
        if key in clienteDestinoDict:
            pass
        clienteDestinoDict[key] = row[1:] 
    print("Comparing Portal-OTM appointments...")

    body = ''
    body +="<img src='DHL.png' width='287' height='82'>"
    if len(expiredAccounts) > 0:
        body += "<p style='font-family:sans-serif;'><b>Usuarios desactualizados</b></p>"
        for i in range (len(expiredAccounts)):
            body += "<p style='font-family:sans-serif;'>{0}</p>".format(expiredAccounts[i])
        body += "<p style='font-family:sans-serif;'>Favor de actualizar el archivo  deptos$\\MXCUTWS0001(S:)\\TRANSPORTE\\LPC\\ApptUser\\Usuarios.csv con las cuentas correspondientes.</p>"
        enviarCorreo="Si"    

    orderEmail = []
    sinLateDelivery = "<p style='font-family:sans-serif;'><b>Citas sin Late Delivery Date en OTM</b></p>"
    lateDeliveryDiferente = "<p style='font-family:sans-serif;'><b>Citas con Late Delivery Diferente</b></p>"
   
    tableTemp=light[(light["CONFIRMATION"]==(int)(master_citas[9][0]))]
    if (not tableTemp["LATE DELIVERY DATE"].iloc[12] =='nan'):
        print("Not nan")
    else:
        print("nan")
    for i in range(len(master_citas)): #generate a new table with the one that matched then use that table on the other comparision
        tableTemp=light[(light["CONFIRMATION"]==(int)(master_citas[i][0]))]
                
        if (len(tableTemp)==1):            
            if (not (tableTemp['ORDER_RELEASE_GID'].values[0] in orderEmail)):   
                if(tableTemp['LATE DELIVERY DATE'].values[0]=="nan"):
                    sinLateDelivery +="<p style='font-family:sans-serif;'>Order Release: {0} Shipment: {1} Cuenta: {2} Confirmacion:\
                        {3}<br><br></p>".format(tableTemp['ORDER_RELEASE_GID'].values[0],tableTemp['SHIPMENT_XID'].values[0],tableTemp['CUENTA'].values[0],tableTemp['CONFIRMATION'].values[0])
                    
                    orderEmail.append(tableTemp['ORDER_RELEASE_GID'].values[0])
                    anySinLateDelivery = 'Si'
                    enviarCorreo="Si"   
                
                else:
        #         tableTemp['LATE DELIVERY DATE']= str(datetime.strptime(tableTemp['LATE DELIVERY DATE'].values[0],'%d/%m/%Y %I:%M:%S %p'))

        #         #print(tableTemp['LATE DELIVERY DATE'].values[0])
        #         #print(master_citas[i][1])

                    if tableTemp['DESTINO FINAL'].values[0] in clienteDestinoDict:
                        horaLimite = datetime.strptime(clienteDestinoDict[tableTemp['DESTINO FINAL'].values[0]][1],'%H:%M')
    #             if datetime.strptime(tableTemp['LATE DELIVERY DATE'].values[0],'%Y-%m-%d %H:%M:%S') == datetime.strptime(str(master_citas[i][1]),'%Y-%m-%d %H:%M:%S'):
    #                 print("Equals")
    #             else:
                        #lateDeliveryST= datetime.strptime(str(tableTemp['LATE DELIVERY DATE'].values[0]),'%d/%m/%Y %I:%M:%S %p')
                        #if latedeliverySt=='nan'
                        if  datetime.strptime('00:00','%H:%M').hour < datetime.strptime(str(tableTemp['LATE DELIVERY DATE'].values[0]),'%d/%m/%Y %I:%M:%S %p').hour > horaLimite.hour:
                            tableTemp['LATE DELIVERY DATE']= pandas.to_datetime(tableTemp['LATE DELIVERY DATE'].values[0])+timedelta(days=1) #If it's between the range substract one day
                    else:
                        horaLimite = datetime.strptime(clienteDestinoDict['LOS DEMAS'][1],'%H:%M')
    #             if datetime.strptime(tableTemp['LATE DELIVERY DATE'].values[0],'%Y-%m-%d %H:%M:%S') == datetime.strptime(str(master_citas[i][1]),'%Y-%m-%d %H:%M:%S'):
    #                 print("Equals")
    #             else:
                        if datetime.strptime('00:00','%H:%M').hour < datetime.strptime(str(tableTemp['LATE DELIVERY DATE'].values[0]),'%d/%m/%Y %I:%M:%S %p').hour > horaLimite.hour:
                            tableTemp['LATE DELIVERY DATE'].values[0]=pandas.to_datetime(tableTemp['LATE DELIVERY DATE'].values[0])+timedelta(days=1)
            
    #         #If mismatch found add to the body of the mail 
                    lightLateDelivery = datetime.strftime(pandas.to_datetime(tableTemp['LATE DELIVERY DATE'].values[0]),'%d/%m/%Y %I:%M:%S %p')
                    portalLateDelivery = datetime.strftime(master_citas[i][1], '%d/%m/%Y %I:%M:%S %p')
                    if (not lightLateDelivery == portalLateDelivery):  
                        print(lightLateDelivery, portalLateDelivery, tableTemp['ORDER_RELEASE_GID'].values[0]) 
        #             #tableTemp['LATE DELIVERY DATE']= str(pandas.to_datetime(tableTemp['LATE DELIVERY DATE'].values[0])-timedelta(days=1))
                        lateDeliveryDiferente +="<p style='font-family:sans-serif;'>Order Release: {0} <br> Shipment: {1} Destino Final: {2} <br>Late Delivery Date en OTM: {3} <br> \
                        Late Delivery Date en Portal Walmart: {4} <br>Tipo Viaje: {5} <br> Cuenta: {6} \
                        <br> Confirmacion: {7}<br><br></p> ".format(tableTemp['ORDER_RELEASE_GID'].values[0],tableTemp['SHIPMENT_XID'].values[0],tableTemp['DESTINO FINAL'].values[0],tableTemp['LATE DELIVERY DATE'].values[0],master_citas[i][1],tableTemp['TIPO VIAJE'].values[0],tableTemp['CUENTA'].values[0],tableTemp['CONFIRMATION'].values[0])          
                   
                    orderEmail.append(tableTemp['ORDER_RELEASE_GID'].values[0])
                    anyLateDeliveryDiferente = 'Si'
                    enviarCorreo="Si"

        elif(len(tableTemp)>1):
            #print(len(tableTemp))
            for j in range(len(tableTemp)):
                #if (not (tableTemp['ORDER_RELEASE_GID'].values[j] in orderEmail)):s 
                if(not tableTemp['LATE DELIVERY DATE'].iloc[j] == 'nan'):
                    print("Is DATE",tableTemp['LATE DELIVERY DATE'].iloc[j],i,j)
                    # print(tableTemp['LATE DELIVERY DATE'])

                    # print(tableTemp['LATE DELIVERY DATE'].values[0],"values[0]")
                    # print(tableTemp['LATE DELIVERY DATE'].values[1],"[1]")
                                      
                    # if(tableTemp['LATE DELIVERY DATE'].values[j] is None):
                    #     print(tableTemp['LATE DELIVERY DATE'].values[j], "isnull")
                    
                    # else:
                    #     print(tableTemp['LATE DELIVERY DATE'].values[j], "notnull")
                        # sinLateDelivery +="<p style='font-family:sans-serif;'>Order Release: {0} Shipment: {1} Cuenta: {2} Confirmacion:\
                        # {3}<br><br></p>".format(tableTemp['ORDER_RELEASE_GID'].values[j],tableTemp['SHIPMENT_XID'].values[j],tableTemp['CUENTA'].values[j],tableTemp['CONFIRMATION'].values[j])

                        # orderEmail.append(tableTemp['ORDER_RELEASE_GID'].values[j])
                        # anySinLateDelivery = 'Si'
                        # enviarCorsreo="Si"

        #             else:                        
        #                 if tableTemp['DESTINO FINAL'].values[j] in clienteDestinoDict:
        #                     #print(tableTemp['DESTINO FINAL'].values[j])
        #                     #print("In destino final",j)
        #                     horaLimite = datetime.strptime(clienteDestinoDict[tableTemp['DESTINO FINAL'].values[j]][1],'%H:%M')    

                                              
        #                     #if datetime.strptime(tableTemp['LATE DELIVERY DATE'].values[0],'%d/%m/%Y %H:%M:%S %p') == datetime.strptime(str(master_citas[i][1]),'%Y-%m-%d %H:%M:%S'):
        #                     #    print("Equals")
        # #                     else:
        #                     #print(datetime.strptime(tableTemp['LATE DELIVERY DATE'].values[j],'%d/%m/%Y %H:%M:%S %p').hour,"hoursss")
                            
        #                     #if  datetime.strptime('00:00','%H:%M').hour < datetime.strptime(str(tableTemp['LATE DELIVERY DATE'].values[j]),'%d/%m/%Y %I:%M:%S %p').hour > horaLimite.hour:
        #                     if  datetime.strptime('00:00','%H:%M').hour < pandas.to_datetime(tableTemp['LATE DELIVERY DATE'].values[j]).hour > horaLimite.hour:
        #                             tableTemp['LATE DELIVERY DATE'].values[j]= pandas.to_datetime(tableTemp['LATE DELIVERY DATE'].values[j])+timedelta(days=1)
        #                 #print(tableTemp['LATE DELIVERY DATE'].values[j],master_citas[i][1])
        #                 else:
        #                     horaLimite = datetime.strptime(clienteDestinoDict['LOS DEMAS'][1],'%H:%M')
        # #                     if datetime.strptime(tableTemp['LATE DELIVERY DATE'].values[0],'%d/%m/%Y %H:%M:%S %p') == datetime.strptime(str(master_citas[i][1]),'%Y-%m-%d %H:%M:%S'):
        # #                         print("Equals")
        # #                     else:
        #                     #if datetime.strptime('00:00','%H:%M').hour < datetime.strptime(str(tableTemp['LATE DELIVERY DATE'].values[j]),'%d/%m/%Y %I:%M:%S %p').hour > horaLimite.hour:
        #                     if datetime.strptime('00:00','%H:%M').hour < pandas.to_datetime(tableTemp['LATE DELIVERY DATE'].values[j]).hour > horaLimite.hour:
        #                         tableTemp['LATE DELIVERY DATE'].values[j]= pandas.to_datetime(tableTemp['LATE DELIVERY DATE'].values[j])+timedelta(days=1)
                        
        #                 #print(tableTemp['LATE DELIVERY DATE'].values[j])
        #                 #tableTemp['LATE DELIVERY DATE'].values[j] = datetime.strftime(tableTemp['LATE DELIVERY DATE'].values[j],'%Y-%m-%d %H:%M:%S')
        #                 #sprint(tableTemp['LATE DELIVERY DATE'].values[j],"check")
        #                 #lightLateDelivery = datetime.strftime(pandas.to_datetime(tableTemp['LATE DELIVERY DATE'].values[j]),'%d/%m/%Y %I:%M:%S %p')
        #                 lightLateDelivery = tableTemp['LATE DELIVERY DATE'].values[j]
        #                 portalLateDelivery = datetime.strftime(master_citas[i][1], '%d/%m/%Y %I:%M:%S %p')
        #                 if (not lightLateDelivery == portalLateDelivery):  
        #                     print(lightLateDelivery, portalLateDelivery, tableTemp['ORDER_RELEASE_GID'].values[j])                         
        # #                     #tableTemp['LATE DELIVERY DATE']= str(pandas.to_datetime(tableTemp['LATE DELIVERY DATE'].values[0])-timedelta(days=1))
        #                     lateDeliveryDiferente +="<p style='font-family:sans-serif;'>Order Release: {0} <br> Shipment: {1} Destino Final: {2} <br>Late Delivery Date en OTM: {3} <br> \
        #                     Late Delivery Date en Portal Walmart: {4} <br>Tipo Viaje: {5} <br> Cuenta: {6} \
        #                     <br> Confirmacion: {7}<br><br></p> ".format(tableTemp['ORDER_RELEASE_GID'].values[0],tableTemp['SHIPMENT_XID'].values[0],tableTemp['DESTINO FINAL'].values[0],tableTemp['LATE DELIVERY DATE'].values[0],master_citas[i][1],tableTemp['TIPO VIAJE'].values[0],tableTemp['CUENTA'].values[0],tableTemp['CONFIRMATION'].values[0])                  
                            
        #                     anyLateDeliveryDiferente = 'Si'     
        #                     orderEmail.append(tableTemp['ORDER_RELEASE_GID'].values[j])  
        #                     enviarCorreo="Si"
        
    # if anyLateDeliveryDiferente == 'Si':
    #     body += lateDeliveryDiferente
    # if anySinLateDelivery == 'Si':
    #     body += sinLateDelivery    
    # if anySinLateDelivery == 'Si' or anyLateDeliveryDiferente == 'Si':
    #     body += "<br><p style='font-family:sans-serif;'>Validar estatus y capturar la información en OTM a la brevedad. ¡Muchas Gracias!</p>"

    # if (enviarCorreo == "Si"):
    #     #sendEmail("OM-LLPC@DHL.COM;Julio.VegaC@dhl.com;Alejandro.RiveraD@dhl.com;Diego.MartinezG@dhl.com; Alejandro.PeraltaD@dhl.com ",body,"Reporte Inconsistencias")
    #     sendEmail("jesus.vasquezsanchezs@dhl.com",body,"Reporte Inconsistencias")
    # else:
    #     print("Usuarios actualizados, no hay inconsistencias.")

#-----------------------------------------------------------------------------------------------------
# Method that filter the info requierd from the prime light
def lightReading(Route):
    data = pandas.read_csv(Route, encoding="ISO-8859-1")
    return data

#-----------------------------------------------------------------------------------------------------
# Example of route call --->data=readFile(r'C:\Users\...\file.txt')
def readFile(route, typeF):  # ReadFile Method
    if typeF == "txt":
        f = open(route, "r")
        usern = str(f.readline().strip()) #strip removes /t
        passwd = str(f.readline().strip())
        f.close()
        data = [usern, passwd]
        return data
    if typeF == "csv":
        data = []
        with open(route) as csvfile:
            contentreader = csv.reader(csvfile)
            cuentas = []
            users = []
            passws = []
            for row in contentreader:
                cuentas.append(row[0])
                users.append(row[1])
                passws.append(row[2])
            data = [cuentas, users, passws]
        return data

def validEstatus():
    time = os.path.getmtime(r'S:\TRANSPORTE\LPC\Power BI\Qlikview\Extract\Prime\Data\Prime.csv')
    dateTime = datetime.fromtimestamp(time)
    timeFormated= dateTime.strftime('%d/%m/%Y %H:%M:%S')
    before = datetime.now()-timedelta(hours=1)
    now = datetime.now()
    if before < dateTime < now:
        print("Prime se encuentra actualizado")
        verificacionCita()
    else:
        print("Generación de Prime atrasada")

def main():
    #try:
    validEstatus()
    #except Exception as e:
        #print('Excepción en main', main)

#-----------------------------------------------------------------------------------------------------
#asyncio.get_event_loop().run_until_complete(fsk_appointment_portal('10101071','DHL900821M4','Test')

schedule.every().day.at("09:35").do(main)

main()

while (True):
    schedule.run_pending()
    time.sleep(1)
