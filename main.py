import asyncio, os, time, datetime, pyppeteer
from pyppeteer import launch

#Automate appointment version 2
usern = ""
passwd = ""
cR = ["5227162139","682913782" ,"681782891"]
firstDate = "2019-10-17 13:04:00"
LateDate="2019-10-18 13:04:00"

async def puppet():
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
        await username.type('so5mu65')
        await password.type('Sanloro44')
        await page.click(strbtn)

        
        contenido = await page.content()
        print(contenido)
        await page.waitForNavigation()

        print("LOGIN SUCCESS")
        print(browser.targets())
        #Executing internal func
        title = await page.evaluate('() => openWin(\'/navis/default.aspx\',\'Appointment_Scheduling\',640,480)')
        await page.waitFor(5000)
        print(browser.targets())
        popup_page = await(browser.targets()[len(browser.targets()) - 1]).page()
        contenido = await popup_page.content()
        print(contenido)
        

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

def readFile():
    print("estoy imprimiendo")
    f = open(r'C:\Users\jesushev\repos\.git\RPA-appointment-supervisor-master\appointData.txt',"r") 
    usern=f.readline()
    passwd=f.readline()
    f.close()
    data=[usern,passwd]
    return data
    

#Method automatized login
async def login(page,usern,passwd):  
    #try:
    #await page.waitFor(1000)    
    await page.setViewport({ 'width':1200, 'height':720})
    await page.goto('https://dsctmsr2.dhl.com/GC3/glog.webserver.servlet.umt.Login')
    #await asyncio.sleep(5)
    await page.waitFor("[name='submitbutton']") # wait for the login button to continue
    
    
    await page.type("[name='username']", usern)
    await page.type("[name='userpassword']", passwd)
    await page.click("[name='submitbutton']")
    #await page.waitFor("[name='Link1_2']")
    await page.waitFor(2000)
    #call captureFunct
    #try:
    await captureOTM(page,cR,len(cR)-1,firstDate,lateDate)
    #except pyppeteer.errors.NetworkError:
    #await asyncio.sleep(2)
    #print("You should't be here")
            

#Start capture of appointments in OTM
async def captureOTM(page,cR,appointI,firstDate, lateDate):    
    #try:
    for i in range(appointI):
        await page.goto('https://dsctmsr2.dhl.com/GC3/glog.webserver.finder.FinderServlet?ct=NzY5Nzg2NjExNDQwNjgzNTIyMg%3D%3D&query_name=glog.server.query.order.OrderReleaseQuery&finder_set_gid=MXCORP.MX%20OM%20ORDER%20RELEASE')
        await page.waitFor("[name='orrOrderReleaseRefnumValue59']") #Wait for the order release)
        await page.type("[name='orrOrderReleaseRefnumValue59']",cR[i])
        await page.keyboard.press('Enter')
        await page.waitFor("[name='rgSGSec.1.1.1.1.check']")
        await page.click("[name='rgSGSec.1.1.1.1.check']")
        await page.click("[id='rgMassUpdateImg']") 
        #await page.waitFor(2000)
        frame = page.frames[3]    
        await frame.waitFor("[name='order_release/late_delivery_date']") #Wait for the order release)
        await frame.waitFor("[name='order_release/ship_with_group']")
        checked= await frame.querySelector("[name='order_release/delivery_is_appt']")
        buttonstatus=await(await checked.getProperty('checked')).jsonValue()
        #print(buttonstatus)
        if buttonstatus==True:
            await frame.type("[name='order_release/late_delivery_date']",lateDate)
        else:
            await frame.type("[name='order_release/ship_with_group']",cR[i])
            await frame.type("[name='order_release/early_delivery_date']",firstDate)
            await frame.type("[name='order_release/late_delivery_date']",lateDate)
            await frame.click("[name='order_release/delivery_is_appt']")               
    print("SUCCESS-----")
        #except pyppeteer.errors.NetworkError:
        #   await login(page)           
        #await asyncio.sleep(2)            

async def main():
    data=readFile() 
    usern=data[0]
    passwd=data[1]   
    #print(data[0])
    #print(data[1])
    browser = await launch(headless=False)  #headless false means open the browser in the operation
    page = await browser.newPage()  
    try:        
        await login(page,usern,passwd)
    except: 
        await main()
        print("FAILED----")        
        print("Retrying----")  
    await page.waitFor(1
                    
                
asyncio.get_event_loop().run_until_complete(puppet())            
asyncio.get_event_loop().run_until_complete(main())
