import asyncio, os, time, datetime, pyppeteer,pandas
from pyppeteer import launch

#Automate appointment version 2
oR = ["3015800169-001","3015800167-001","681782891-002","5227162139-002"]
cR = ["3015800169","3015800167" ,"681782891","5227162139"]
firstDate = "2019-11-4 13:04:00"
lateDate="2019-11-4 13:04:00"


async def wm_appointment_portal():

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

    data=readFile(r'C:\Users\jesushev\Documents\RPA-appointment-supervisor\walmartD.txt',"txt")

    print("Filling form...")
    await username.type(data[0])
    await password.type(data[1])
    await page.click(strbtn)

    await page.waitForNavigation()
    print("Succesful login...Navigating")

    #Opening query session
    testpage = await browser.newPage()
    await testpage.goto("https://retaillink.wal-mart.com/navis/default.aspx")
    await testpage.waitForNavigation()
    await testpage.waitForNavigation()
    await testpage.waitForNavigation()

    #Opening this week Deliveries
    await testpage.goto("https://logistics-scheduler-www9.wal-mart.com/trips_mx/quickQuery.do?type=thisWeeksDeliveries")
    xtabla = await testpage.waitForXPath("//*[@id=\"SortTable0\"]/tbody")
    #Extracting table size
    tabla = await testpage.evaluate('(xtabla) => xtabla.children', xtabla)
    master_citas = []

    for i in range(1,len(tabla)+1):
        index = str(i)
        x1 = await testpage.waitForXPath("//*[@id=\"SortTable0\"]/tbody/tr[{}]/td[1]/a".format(index))
        no_entrega = await testpage.evaluate('(x1) => x1.innerText', x1)
        x2 = await testpage.waitForXPath("//*[@id=\"SortTable0\"]/tbody/tr[{}]/td[8]".format(index))
        cita = await testpage.evaluate('(x2) => x2.innerText', x2)
        #Conversión a datetime
        clean_cita = datetime.datetime.strptime(cita, '%m/%d/%y %I:%M %p')
        master_citas.append([no_entrega,clean_cita])

    print("Succesful extraction... \nResults:\n",master_citas)
    await page.waitFor(5000)
    try:
        await browser.close()
    except:
        print('Error al cerrar, same as always')
    return master_citas


#clean cita se compara contra el late delivery date > o =
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
        
def csvReading():
    data=pandas.read_csv(r"C:\Users\jesushev\Documents\RPA-appointment-supervisor\light.csv",encoding="ISO-8859-1")
    #print(data.head(5))

async def captureOTM():  
    try:
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
        await page.waitFor(3000)
        await browser.close()

    except:
        await browser.close() 
        print("FAILED----")        
        print("RETRYING----")         
        await captureOTM()

#csvReading()                                          
data=asyncio.get_event_loop().run_until_complete(wm_appointment_portal()) 
for i in range(len(data)):
    print("VALUE: ",data[i])
asyncio.get_event_loop().run_until_complete(captureOTM())  