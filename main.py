''' Automate appointments prototype 0.5
    Authors:
        Heriberto Vasquez Sanchez
        Jose Rodrigo Narvaez Berlanga'''

import asyncio, os, time, pyppeteer, pandas, csv,numpy,calendar
from pyppeteer import launch
from datetime import datetime,timedelta

master_citas = [] #master all appointments
master_report = [] #master all reports by Control Vehicular

#Walmart appointment extraction method
async def wm_appointment_portal(user,passwd,account_name):

    browser = await launch(headless=False)
    strusername = "body > div > div > div > div.page-container > div.main-container > div.content-container > div > div > form > span:nth-child(1) > span > span > input"
    strpass = "body > div > div > div > div.page-container > div.main-container > div.content-container > div > div > form > span:nth-child(2) > span > span > input"
    strbtn = "body > div > div > div > div.page-container > div.main-container > div.content-container > div > div > form > button"
    entregasEncontradas="#mc > tbody > tr:nth-child(2) > td.contentPanel > table > tbody > tr:nth-child(4) > td > form > table > tbody > tr > td > table > tbody > tr.contentBodyRow > td.contentBody > table.formTable > tbody > tr:nth-child(2) > td.valueTd"

    page = await browser.newPage()
    await page.setViewport({"width": 1024, "height": 768, "deviceScaleFactor": 1})
    page.setDefaultNavigationTimeout(30000) # maybe 60000
    await page.goto("https://retaillink.login.wal-mart.com/?ServerType=IIS1&CTAuthMode=BASIC&language=en&utm_source=retaillink&utm_medium=redirect&utm_campaign=FalconRelease&CT_ORIG_URL=/&ct_orig_uri=/")
    await page.waitFor(strusername)

    username = await page.querySelector(strusername)
    password = await page.querySelector(strpass)

    print("Filling form...")
    await username.type(user)
    await password.type(passwd)
    await page.click(strbtn)
    try: 
        await page.waitForNavigation()
        print("Succesful login...Navigating")
    except:
        print("Failed in",account_name,"login.")
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
    #print("LENGTH", tabla)
    
    #Use try to manage when there are no near appointments 

    noAppointFound= await testpage.querySelector(entregasEncontradas)
    noAppointments=await testpage.evaluate('(noAppointFound) => noAppointFound.textContent',noAppointFound)
    if noAppointments=='0':
        print("NO",account_name,"APPOINTMENTS IN PORTAL")
        await page.waitFor(5000)
        await browser.close()
    else:
        for i in range(1, len(tabla) + 1):
            index = str(i)
            x1 = await testpage.waitForXPath('//*[@id="SortTable0"]/tbody/tr[{}]/td[1]/a'.format(index))
            no_entrega = await testpage.evaluate("(x1) => x1.innerText", x1)
            x2 = await testpage.waitForXPath('//*[@id="SortTable0"]/tbody/tr[{}]/td[8]'.format(index))
            cita = await testpage.evaluate("(x2) => x2.innerText", x2)
            # Conversión a datetime
            clean_cita = datetime.strptime(cita,'%m/%d/%y %I:%M %p')#.strftime('%m/%d/%Y %I:%M:%S %p')
            master_citas.append([no_entrega, clean_cita])
        
        print("Succesful",account_name,"extraction")
        # for i in range(len(master_citas)):
        #   print("VALUE: ",master_citas[i])
        await page.waitFor(5000)
        try:
            await browser.close()
        except:
            print("Error al cerrar, same as always")
        return master_citas


#-----------------------------------------------------------------------------------------------------
#Validate appointments in OTM 
async def captureOTM(arrCR):
    # try:
    # browser = await launch({'args': ['--disable-dev-shm-usage']})
    browser = await launch(headless=False)  # headless false means open the browser in the operation
    page = await browser.newPage()
    await page.setViewport({"width": 1024, "height": 768, "deviceScaleFactor": 1})
    page.setDefaultNavigationTimeout(60000)
    await page.goto("https://dsctmsr2.dhl.com/GC3/glog.webserver.servlet.umt.Login")
    # await page.waitFor(2000) 
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
        # Aqui checamos si la cita hace match con el sitio web de walmart
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
        #await page.waitFor(1000)
        await frame.waitFor("[name='order_release/delivery_is_appt']")
        await frame.waitFor("[name='order_release/early_delivery_date']")
        await frame.waitFor("[name='order_release/late_delivery_date']") #Wait for the order release)
        await frame.waitFor("[name='order_release/ship_with_group']")
        # #print("Success")
        # checked= await frame.querySelector("[name='order_release/delivery_is_appt']")
        # buttonstatus=await(await checked.getProperty('checked')).jsonValue()
        # if buttonstatus==True:
        #     await frame.type("[name='order_release/late_delivery_date']",lateDate)
        # else:
        await frame.type("[name='order_release/ship_with_group']",arrCR[i])
        #await frame.type("[name='order_release/early_delivery_date']",firstDate)
        await frame.type("[name='order_release/late_delivery_date']","lateDate")
        await frame.click("[name='order_release/delivery_is_appt']")
        await frame.waitFor(1000)
        #print("PASSED")
    #print("SUCCESS----")
    '''return returned
    await page.waitFor(3000)
    await browser.close()'''

# except:
#     await browser.close()
#     print("FAILED----")
#     print("RETRYING----")
#     await captureOTM()


def destinoFinal():
    masterClienteDestino=[]
    with open(r'C:\Users\jesushev\Desktop\CLIENTE DESTINO.csv') as credentials:
        gen_reader = csv.reader(credentials, delimiter = ',')
        next(gen_reader, None) #Skips headers        
        for row in gen_reader:
            tempCD=[]
            clienteDestino = row[0]
            destino = row[1]
            minTime = row[2]
            maxTime = row[3]
            tempCD.append(clienteDestino)
            tempCD.append(destino)
            tempCD.append(minTime)
            tempCD.append(maxTime)
            masterClienteDestino.append(tempCD)
    return masterClienteDestino

#-----------------------------------------------------------------------------------------------------
# Here we do the verification between the walmart site and OTM, consolidating data
def verificacionCita():
    extract all apointments on master all citas
    with open(r'\\Mxmex1-fipr01\public$\Nave 1\LPC\ApptUsers\USUARIO.csv') as credentials:
        gen_reader = csv.reader(credentials, delimiter = ',')
        next(gen_reader, None) #Skips headers
        for row in gen_reader:
            account_name = row[0]
            user = row[1]
            password = row[2]
            if account_name=='JDE COFFE':
                password='Agosto2020'
            asyncio.get_event_loop().run_until_complete(wm_appointment_portal(user,password,account_name))
    

    light=lightReading(r'\\Mxmex1-fipr01\public$\Nave 1\LPC\Prime_Light.csv')#Now read prime light
    lightArr=light.to_numpy()

    tabla=pandas.DataFrame() # data of table with just matchs of confirmation
    arrTabla=[]    
    for i in range(len(master_citas)): #generate a new table with the one that matched then use that table on the other comparision
        tableTemp=light[(light["CONFIRMATION"]==(int)(master_citas[i][0]))] #here match with the confirmation 
        tableTemp['LATE DELIVERY DATE']=tableTemp['LATE DELIVERY DATE'].apply(lambda x: datetime.strptime(x,'%d/%m/%Y %H:%M')) #Apply time format to late delivery date
            
        #Generate Hour Day and Day of week from otm 
        tableTemp['Hour']=pandas.to_datetime(tableTemp['LATE DELIVERY DATE']).dt.hour
        tableTemp['Day']=pandas.to_datetime(tableTemp['LATE DELIVERY DATE']).dt.day
        tableTemp['Day of the week']=pandas.to_datetime(tableTemp['LATE DELIVERY DATE']).dt.day_name()

        #extract hour day and day of week from Walmart
        wmHour=master_citas[i][1].hour
        wmDay=master_citas[i][1].day
        wmDayWeek=master_citas[i][1].strftime("%A")
        #print(tableTemp)
        tabla=tabla.append(tableTemp)
    
    #Now our table has all the matches by confirmation.
    clienteDestino=destinoFinal()  
    tablaCD=pandas.DataFrame() #table that will contain just the ones that match cliente destino.
    if tabla.empty:
        print("Tabla empty")
    else:
        for i in range(len(clienteDestino)): #now generate the match with clienteDestino
            temp=tabla[(tabla["CLIENTE DESTINO"]==clienteDestino[i][0])] #now we'll have our data filtred by cliente destino append that.
            tablaCD=tablaCD.append(temp)
    
    if tablaCD.empty:
        print("TablaCD empty")
    else:
        for j in range(len(tablaCD)): #Do it for every register in tablaCD
            hourInPrime=datetime.fromtimestamp(int((tablaCD['Hour'])[j])).strftime('%H:%M')
            for i in range(len(clienteDestino)): #now validate hour from cliente destino rules
                if (hourInPrime>=clienteDestino[i][2] and hourInPrime<clienteDestino[i][3]): #cliente destino has not the same lenght of masterCitas
                    #If the hour it's between the two of them add one to do the match with walmart
                    tablaCD['LATE DELIVERY DATE']=pandas.to_datetime(tablaCD['LATE DELIVERY DATE'])+timedelta(days=1)
                    tablaCD['Day']=pandas.to_datetime(tablaCD['LATE DELIVERY DATE']).dt.day
                    tablaCD['Day of the week']=pandas.to_datetime(tablaCD['LATE DELIVERY DATE']).dt.day_name()
    
    #Now that you have updated the table now you have to check the conditions if it is different.
    if tablaCD.empty:
        print("TablaCD empty nothing to compare")
    else:
        for i in range(len(tablaCD)): #len of tabCD!= master citas, we compare every i register of tablaCD to all the registers of master citas
            for j in range(len(master_citas)):
                if tablaCD['LATE DELIVERY DATE'][i]!=master_citas[j][1]:
                    print("Cita con confirmacion",tablaCD['CONFIRMATION'][i],"inconsistente")
                    print("Fecha en portal Walmart:",master_citas[j][1],"Fecha capturada en OTM:",tablaCD['LATE DELIVERY DATE'][i])
                    break
                    
                    #now let's updated in OTM
                #captureOTM(tabla['CR'])  
    

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
        usern = f.readline()
        passwd = f.readline()
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

#-----------------------------------------------------------------------------------------------------
verificacionCita()
#captureOTM()