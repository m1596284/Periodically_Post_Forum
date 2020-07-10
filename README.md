# Periodically_Post_Forum
Periodically login to a forum and post a article. [User simulation, Numeric captcha, Windows control]

## User :
Prepare your article in word and account info in txt file. Launch the program and it will automatically start.  
<img src="https://i.imgur.com/o8kj6EQ.gif" width="647" height="426">

## Backstage : Python + Selenium + Google Tesseract
With: Selenium, Chromedriver, pytesseract, docx, win32gui, win32api, win32con  

#### Selenium: Simulate user controlling in webdriver (for me, chromedriver)
Use win32api to simulate typing data. "win32api.keybd_event(0x0D, 0, 0, 0)" for "Enter"  
Use win32com to simulate openning file in word. "win32com.client.Dispatch('Word.Application')"  
Use ActionChains to simulate clicking. "actionRight.context_click(imgPic).perform()"

#### Tesseract: image process and verify
Image Process: "img = Image.open(pyPath + "/cha.png").convert("L")", include binarizing, converting...etc.  
Image Verify: "logAdd = pytesseract.image_to_string(img, lang='eng')"  
<img src="https://i.imgur.com/YkUebX0.png" width="50" height="20"> => 4QHK
