''' Automate appointments prototype 0.5
    Authors:
        Heriberto Vasquez Sanchez
        Jose Rodrigo Narvaez Berlanga'''

import asyncio, os, time, pyppeteer, pandas, csv,numpy
from pyppeteer import launch
from datetime import datetime

#oR = ["33976-001", "34514-001", "5602993825-001"]
#confirmCita=["7865881","7955784","7944631"]
firstDate = "2019-11-4 13:04:00"
lateDate = "2019-11-4 13:04:00"

#Walmart appointment extraction method
#Consider change of user.
async def wm_appointment_portal(user,passwd):

    browser = await launch(headless=False)
    strusername = "body > div > div > div > div.page-container > div.main-container > div.content-container > div > div > form > span:nth-child(1) > span > span > input"
    strpass = "body > div > div > div > div.page-container > div.main-container > div.content-container > div > div > form > span:nth-child(2) > span > span > input"
    strbtn = "body > div > div > div > div.page-container > div.main-container > div.content-container > div > div > form > button"

    page = await browser.newPage()
    await page.setViewport({"width": 1024, "height": 768, "deviceScaleFactor": 1})
    page.setDefaultNavigationTimeout(60000)
    await page.goto("https://retaillink.login.wal-mart.com/?ServerType=IIS1&CTAuthMode=BASIC&language=en&utm_source=retaillink&utm_medium=redirect&utm_campaign=FalconRelease&CT_ORIG_URL=/&ct_orig_uri=/")
    await page.waitFor(strusername)

    username = await page.querySelector(strusername)
    password = await page.querySelector(strpass)

    #data = readFile(r"walmartD.txt", "txt")

    print("Filling form...")
    await username.type(user)
    await password.type(passwd)
    await page.click(strbtn)
    
    await page.waitForNavigation()
    print("Succesful login...Navigating")
    
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
    master_citas = []

    for i in range(1, len(tabla) + 1):
        index = str(i)
        x1 = await testpage.waitForXPath('//*[@id="SortTable0"]/tbody/tr[{}]/td[1]/a'.format(index))
        no_entrega = await testpage.evaluate("(x1) => x1.innerText", x1)
        x2 = await testpage.waitForXPath('//*[@id="SortTable0"]/tbody/tr[{}]/td[8]'.format(index))
        cita = await testpage.evaluate("(x2) => x2.innerText", x2)
        # Conversión a datetime
        clean_cita = datetime.strptime(cita, "%m/%d/%y %I:%M %p")
        master_citas.append([no_entrega, clean_cita])

    print("Succesful extraction...")
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
async def captureOTM(oR):
    # try:
    # browser = await launch({'args': ['--disable-dev-shm-usage']})
    browser = await launch(headless=False)  # headless false means open the browser in the operation
    page = await browser.newPage()
    await page.setViewport({"width": 1024, "height": 768, "deviceScaleFactor": 1})
    page.setDefaultNavigationTimeout(60000)
    await page.goto("https://dsctmsr2.dhl.com/GC3/glog.webserver.servlet.umt.Login")
    await page.waitFor(1000)
    data = readFile(r"appointData.txt", "txt")
    await page.waitFor("[name='userpassword']")
    await page.waitFor("[name='username']")
    await page.type("[name='userpassword']", data[1])
    await page.type("[name='username']", data[0])
    # await page.click("[name='submitbutton']") 
    returned=[]
    for i in range(len(oR)):
        valuei=[]
        await page.goto("https://dsctmsr2.dhl.com/GC3/glog.webserver.finder.FinderServlet?ct=NzY5Nzg2NjExNDQwNjgzNTIyMg%3D%3D&query_name=glog.server.query.order.OrderReleaseQuery&finder_set_gid=MXCORP.MX%20OM%20ORDER%20RELEASE")
        await page.waitFor(1000)
        await page.waitFor("[name='order_release/xid']")  # Wait for the order release
        # await page.waitFor("[name='orrOrderReleaseRefnumValue59']") #Wait for the CR
        await page.type("[name='order_release/xid']", oR[i])
        # await page.type("[name='orrOrderReleaseRefnumValue59']",cR[i])
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
        valuei.append(folio)
        valuei.append(eDate)
        valuei.append(lDate)
        valuei.append(consignatario)
        returned.append(valuei)

        # await page.waitFor("[name='rgSGSec.1.1.1.1.check']")
        # await page.waitFor(1000)
        # await page.click("[name='rgSGSec.1.1.1.1.check']")
        # await page.waitFor("[title='Mass Update']")
        # await page.waitFor(1000)
        # await page.click("[title='Mass Update']")
        # frames=page.frames
        # temp= len(frames)
        # while temp < 3 : #Wait until the frame is loaded
        #     temp= len(frames)
        # frame = page.frames[3]
        # #await page.waitFor(1000)
        # await frame.waitFor("[name='order_release/delivery_is_appt']")
        # await frame.waitFor("[name='order_release/early_delivery_date']")
        # await frame.waitFor("[name='order_release/late_delivery_date']") #Wait for the order release)
        # await frame.waitFor("[name='order_release/ship_with_group']")
        # #print("Success")
        # checked= await frame.querySelector("[name='order_release/delivery_is_appt']")
        # buttonstatus=await(await checked.getProperty('checked')).jsonValue()
        # if buttonstatus==True:
        #     await frame.type("[name='order_release/late_delivery_date']",lateDate)
        # else:
        #     await frame.type("[name='order_release/ship_with_group']",cR[i])
        #     await frame.type("[name='order_release/early_delivery_date']",firstDate)
        #     await frame.type("[name='order_release/late_delivery_date']",lateDate)
        #     await frame.click("[name='order_release/delivery_is_appt']")
        #print("PASSED")
    #print("SUCCESS----")
    return returned
    await page.waitFor(3000)
    await browser.close()

# except:
#     await browser.close()
#     print("FAILED----")
#     print("RETRYING----")
#     await captureOTM()

#-----------------------------------------------------------------------------------------------------
# Here we do the verification between the walmart site and OTM, consolidating data
def verificacionCita():
    usr=readFile(r'USUARIO WALMART.csv',"csv")
    data=[]
    oR=[]
    #falta que de click cuando #tip este en enabled.
    #Por ahora lo probamos con Lenovo
    data.append(asyncio.get_event_loop().run_until_complete(wm_appointment_portal(usr[1][9],usr[2][9])))
    #Then start comparing it with prime light dB
    #print(data[0][4])
    #How to check if your date is alright
    '''dda=data[0][1]
    print(dda[1])
    arrTempo=['11/12/2019 11:00:00 AM']
    if dda[1]==arrTempo[0]:
        print("True")'''
    master_light=[] #Master array of light db
    #print(len(data[0]))
    for i in range(len(data)):
        print(data[0][i][0])
        pands=lightReading('FOLIO 5328480')
        #pands=lightReading(data[0][i][0])  #the file is inside [[[]]] 3 that's why
        fecha=pands[['EARLY DELIVERY DATE']]#[['EARLY DELIVERY DATE']])
        print(fecha.to_numpy()[0])
        #Check if Folio is bad format Example Folio +7734453... (good format: 7734454...)
        '''if pands.empty:
            pands=lightReading("FOLIO "+ data[0][i][0]) 
            #1: Check if is empty:
            if pands.empty:
                master_light.append("Cita con folio: "+data[0][i][0]+" faltante, sin capturar en OTM")
            else:
                master_light.append("Formato de cita desactualizado: "+"FOLIO "+ data[0][i][0]+"||  Formato de cita adecuado: "+ data[0][i][0])
        else:
            #Now check date
            fecha=pands[['EARLY DELIVERY DATE']]#[['EARLY DELIVERY DATE']])
            #important check if there is serveral that share folio so made it of OR.
            fechaWal=(fecha)
            if data[0][i][1]!=fecha:
                master_light.append("Discordancia en fechas en cita con folio: "+data[0][i][0]+"\nFecha en portal: "+data[0][i][1]+"\nFecha en OTM: "+fecha)'''
    print(master_light)
    #pands['Month']=pandas.DatetimeIndex(fecha['EARLY DELIVERY DATE']).month #create a new column if needed, print(pands[["Month"]])  
    # Confirmation was not set 
    # 2:     
    #print(pands[["Mont"]].iloc[0])
    #if fecha==arrTemp:
     #   print("true")
    #fecha.datetime
    #fecha.day()
    #caso folio no en OTM
    #for i in range(len(pands)): #It should be done by every register.
     #   oR.append(pands.iloc[i]['ORDER_RELEASE_GID'])
    #otmData=asyncio.get_event_loop().run_until_complete(captureOTM(oR))

#-----------------------------------------------------------------------------------------------------
# Method that filter the info requierd from the prime light
def lightReading(cita):
    #Reading just confirmacion appointement needed
    data = pandas.read_csv("Prime_Light.csv", encoding="ISO-8859-1")
    tabla = data[(data["CONFIRMACION CITA"]== cita)]
    #appointment=tabla.iloc[1]['ORDER_RELEASE_GID'], tabla[['CONSIGNATARIO','ORDER_RELEASE_GID','EARLY DELIVERY DATE','LATE DELIVERY DATE','CUENTA','CR','CONFIRMACION CITA']]
    return tabla

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