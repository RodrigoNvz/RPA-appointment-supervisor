import asyncio, os, time, datetime, pyppeteer
from pyppeteer import launch

#Automate appointment version 2
oR = ["5227162139-002","682913782-001","681782891-002","5227162139-002","682913782-001","681782891-002"]
cR = ["5227162139","682913782" ,"681782891","5227162139","682913782" ,"681782891"]
firstDate = "2019-10-17 13:04:00"
lateDate="2019-10-18 13:04:00"


async def wm_appointment_portal():
    # print('\n'*60)
    clear = lambda: os.system('cls')  # on Linux System
    clear()
    if(True):
        start = time.time()
        browser = await launch(headless=False)
        strusername = "body > div > div > div > div.page-container > div.main-container > div.content-container > div > div > form > span:nth-child(1) > span > span > input"
        strpass = "body > div > div > div > div.page-container > div.main-container > div.content-container > div > div > form > span:nth-child(2) > span > span > input"
        strbtn = "body > div > div > div > div.page-container > div.main-container > div.content-container > div > div > form > button"

        page = await browser.newPage()
        await page.setViewport({'width': 1024, 'height': 768, 'deviceScaleFactor': 1})
        page.setDefaultNavigationTimeout(60000)
        await page.goto('https://retaillink.login.wal-mart.com/?ServerType=IIS1&CTAuthMode=BASIC&language=en&utm_source=retaillink&utm_medium=redirect&utm_campaign=FalconRelease&CT_ORIG_URL=/&ct_orig_uri=/')
        await page.waitFor(strusername)

        username = await page.querySelector(strusername)
        password = await page.querySelector(strpass)

        data=readFile(r'walmartD.txt',"txt")

        await username.type(data[0])
        await password.type(data[1])
        await page.click(strbtn)
        #await page.waitForSelector("Head1")
        print("Esperando")
        await page.waitForNavigation()
        #await page.waitFor(15000)
        print("YA")
        contenido = await page.content()
        print(contenido)
        print("HIa")

#bodyDataDiv > table > tbody > tr > td:nth-child(1) > table > tbody > tr:nth-child(1) > td > div > input[type=text]:nth-child(3)
        print("LOGIN SUCCESS")
        #targetsAntes = len(browser.targets())
        #print(targetsAntes)
        #await page.waitFor(5000)
        #Executing internal func
        #await browser.browserContext
        #await page.evaluate('() => openWin(\'/navis/default.aspx\',\'Appointment_Scheduling\',640,480)')
        
        #await page.waitFor(5000)

        testpage = await browser.newPage()
        await testpage.goto("https://retaillink.wal-mart.com/navis/default.aspx")
        await testpage.waitForNavigation()
        #await testpage.goto("https://logistics-scheduler-www9.wal-mart.com/retaillink_mx/ServerDecision.jsp")
        await testpage.waitForNavigation()
        #await testpage.goto("https://logistics-scheduler-www9.wal-mart.com/retaillink_mx/NavisDecrypt.jsp")
        await testpage.waitForNavigation()
        print("LISTOOOOOO")
        await testpage.goto("https://logistics-scheduler-www9.wal-mart.com/trips_mx/quickQuery.do?type=thisWeeksDeliveries")
        #await testpage.waitForNavigation()
        print("LISTOOOOOO 2")
        contenidotest = await testpage.content()
        print(contenidotest)
        #bodyDataDiv > table > tbody > tr > td:nth-child(2) > table > tbody > tr:nth-child(5) > td > div > input[type=text]:nth-child(3)
        xtabla = await testpage.waitForXPath("//*[@id=\"SortTable0\"]/tbody")
        tabla = await testpage.evaluate('(xtabla) => xtabla.children', xtabla)
        #print(len(tabla))

        for i in range(1,len(tabla)+1):
            index = str(i)
            x1 = await testpage.waitForXPath("//*[@id=\"SortTable0\"]/tbody/tr[{}]/td[1]/a".format(index))
            no_entrega = await testpage.evaluate('(x1) => x1.innerText', x1)
            x2 = await testpage.waitForXPath("//*[@id=\"SortTable0\"]/tbody/tr[{}]/td[8]".format(index))
            cita = await testpage.evaluate('(x2) => x2.innerText', x2)
            print("Resultado {}: ".format(index), no_entrega, cita)
        #await testpage.waitFor("body > mc > tbody > tr:nth-child(2) > td.contentPanel > table > tbody > tr:nth-child(4) > td > form > table > tbody > tr > td > table > tbody > tr.contentBodyRow > td.contentBody > SortTable0 > thead > tr.tableTitleRow > th > table > tbody > tr > td")
        print("LISTOOOOOO 3")
        print('SUCCESSSS!!!!')
        await page.waitFor(10000)
        await page.close()
        try:
            await browser.close()
        except:
            print('Error al cerrar, same as always')

#favoritesList
        end = time.time()
        
        print('PROCESS EXECUTION TIME: ', (end - start) / 60, ' MIN')
        now = datetime.datetime.now()
        print('\n\n\n\n\n ***** SUCCESSFULL ', now.strftime("%Y-%m-%d %H:%M"), ' ***** \n\n\n\n\n')


#Example of route call
#--->data=readFile(r'C:\Users\...\file.txt')
def readFile(route,typeF): #ReadFile Method
    if typeF=="txt":
        f = open(route,"r") 
        usern=f.readline()
        passwd=f.readline()
        f.close()
        data=[usern,passwd]
        return data
    ##wb = xlrd.open_workbook(route) 
    #sheet = wb.sheet_by_index(0) 
    #return sheet.cell_value(0,0)         

async def captureOTM():  
    #try:
    #browser = await launch({'args': ['--disable-dev-shm-usage']})  #headless false means open the browser in the operation
    browser =await launch(headless=False)
    page = await browser.newPage()  
    await page.setViewport({'width': 1024, 'height': 768, 'deviceScaleFactor': 1})
    page.setDefaultNavigationTimeout(60000)
    await page.goto('https://dsctmsr2.dhl.com/GC3/glog.webserver.servlet.umt.Login')    
    await page.waitFor(1000)
    data=readFile(r'C:\Users\jesushev\Documents\RPA-appointment-supervisor\appointData.txt',"txt")  
    await page.waitFor("[name='userpassword']") 
    await page.waitFor("[name='username']")          
    await page.type("[name='userpassword']", data[1])
    await page.type("[name='username']", data[0])
    #await page.click("[name='submitbutton']")  
    for i in range(len(cR)):
        await page.goto('https://dsctmsr2.dhl.com/GC3/glog.webserver.finder.FinderServlet?ct=NzY5Nzg2NjExNDQwNjgzNTIyMg%3D%3D&query_name=glog.server.query.order.OrderReleaseQuery&finder_set_gid=MXCORP.MX%20OM%20ORDER%20RELEASE')
        await page.waitFor(1000)
        await page.waitFor("[name='order_release/xid']") #Wait for the order release
        await page.waitFor("[name='orrOrderReleaseRefnumValue59']") #Wait for the CR        
        await page.type("[name='order_release/xid']",oR[i])
        await page.type("[name='orrOrderReleaseRefnumValue59']",cR[i])        
        await page.keyboard.press('Enter') 
        await page.waitFor("[name='rgSGSec.1.1.1.1.check']") 
        await page.click("[name='rgSGSec.1.1.1.1.check']")
        await page.waitFor("[title='Mass Update']") 
        await page.click("[title='Mass Update']") 
        frames=page.frames 
        temp= len(frames)  
        while temp < 3 : #Wait until the frame is loaded
            temp= len(frames) 
        frame = page.frames[3] 
        await page.waitFor(1000)
        await frame.waitFor("[name='order_release/delivery_is_appt']")         
        await frame.waitFor("[name='order_release/early_delivery_date']")  
        await frame.waitFor("[name='order_release/late_delivery_date']") #Wait for the order release)
        await frame.waitFor("[name='order_release/ship_with_group']")
        #print("Success")
        checked= await frame.querySelector("[name='order_release/delivery_is_appt']")
        buttonstatus=await(await checked.getProperty('checked')).jsonValue()
        if buttonstatus==True:
            await frame.type("[name='order_release/late_delivery_date']",lateDate)
        else:
            await frame.type("[name='order_release/ship_with_group']",cR[i])
            await frame.type("[name='order_release/early_delivery_date']",firstDate)
            await frame.type("[name='order_release/late_delivery_date']",lateDate)
            await frame.click("[name='order_release/delivery_is_appt']") 

        print("PASSED")
    print("SUCCESS----")
    await page.waitFor(2000)
    await browser.close()

    '''except:
        await browser.close() 
        print("FAILED----")        
        print("RETRYING----")         
        await captureOTM()'''
                                         
#asyncio.get_event_loop().run_until_complete(wm_appointment_portal()) 
asyncio.get_event_loop().run_until_complete(captureOTM())       