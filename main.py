import asyncio, os, time, datetime, pyppeteer
from pyppeteer import launch

#Automate appointment version 2
usern = "MXPLAN.JNARVAEZB"
passwd = "H3ll0w0rld1"
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

#Method automatized login
async def login(page,retries=0):  
    if retries > 10:
        print("10 reintentos fallidos")
        return await page.evaluate('document.body.innerHTML')
    else:
        try:
            await page.setViewport({ 'width':1200, 'height':720})
            await page.goto('https://dsctmsr2.dhl.com/GC3/glog.webserver.servlet.umt.Login')
            await page.waitFor("[name='submitbutton']") # wait for the login button to continue
            await page.type("[name='username']", usern)
            await page.type("[name='userpassword']", passwd)
            await page.click("[name='submitbutton']")
            await page.waitFor(2000)
        except pyppeteer.errors.NetworkError:
            #await asyncio.sleep(2)
            print("You should't be here")
            return await login(page, retries+1)

#Start capture of appointments in OTM
async def captureOTM(page,cR,appointI,firstDate, lateDate, retries=0):
    if retries > 10:
        print("10 reintentos fallidos")
        return await page.evaluate('document.body.innerHTML')
    else:
        try:
            for i in range(appointI):
                await page.goto('https://dsctmsr2.dhl.com/GC3/glog.webserver.finder.FinderServlet?ct=NzY5Nzg2NjExNDQwNjgzNTIyMg%3D%3D&query_name=glog.server.query.order.OrderReleaseQuery&finder_set_gid=MXCORP.MX%20OM%20ORDER%20RELEASE')
                await page.waitFor("[name='orrOrderReleaseRefnumValue59']") #Wait for the order release)
                await page.type("[name='orrOrderReleaseRefnumValue59']",cR[i])
                await page.keyboard.press('Enter')
                await page.waitFor("[name='rgSGSec.1.1.1.1.check']")
                await page.click("[name='rgSGSec.1.1.1.1.check']")
                await page.click("[id='rgMassUpdateImg']") 
                await page.waitFor(2000)
                frame = page.frames[3]    
                await frame.waitFor("[name='order_release/late_delivery_date']") #Wait for the order release)
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
                #appointI +=1   
                #print("appointment:",i," success") 
                #await frame.waitFor(1000)
                #await page.waitFor(1000)
            #await frame.close()
            #await page.close()
            print("SUCCESS-----")
        except pyppeteer.errors.NetworkError:
            #await asyncio.sleep(2)
            print("You should't be here")
            return await captureOTM(page,cR,appointI,firstDate,lateDate, retries+1)

async def main():
    for i in range(3):
        browser = await launch(headless=False)  #headless false means open the browser in the operation
        page = await browser.newPage()
        retries=0 
        # try:
        #     await login(page,retries)
        # except: 
        #     print("FAILED IN LOGIN----")        
        # try:
        #     await captureOTM(page,cR,len(cR),firstDate,lateDate,retries)
        # except:
        #     print("FAILED IN CAPTURE DATA----")   
        await login(page,retries)
        await captureOTM(page,cR,len(cR),firstDate,lateDate,retries)
    await page.waitFor(1000)
    await page.close()     
                    
                
asyncio.get_event_loop().run_until_complete(puppet())            
asyncio.get_event_loop().run_until_complete(main())
