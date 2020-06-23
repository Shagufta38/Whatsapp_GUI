from tkinter import *
from tkinter import messagebox
from functools import partial
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from openpyxl import load_workbook, cell
import time
import sys
from selenium.webdriver.chrome.options import Options
import time
import sys
from tkinter import filedialog
  
top = Tk()
top.title("WhatsApp Application")

top.geometry("600x400")
unsaved_Contacts = []
filename=""

def openfile():
    filename=filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("Excel file","*.xlsx"),("all files","*.*")))
    messagebox.showinfo("Hello", filename)
    workbook = load_workbook(filename)
    sheet = workbook["contact"]
    for cell in sheet['A']:
        target = cell.value

        unsaved_Contacts.append(target)
    return

    
def sendno(s1,s2):
    str1 = (s1.get())  
    str2 = (s2.get())
    message = str1
    target = str2
    link = "https://wa.me/" + str2
    chrome_options = Options()
    chrome_options.add_argument('--user-data-dir=./User_Data')
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 600)
    driver.get(link)
    driver.maximize_window()
    print("QR scanned")
    try:
        time.sleep(20)
        input_box = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
        for ch in message:
            if ch == "\n":
                ActionChains(browser).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.ENTER).key_up(Keys.SHIFT).key_up(Keys.BACKSPACE).perform()
            else:
                input_box.send_keys(ch)
        input_box.send_keys(Keys.ENTER)
        print("Message sent successfuly")
    except NoSuchElementException:
        print("Failed to send message")
    driver.close()
    return

def sendcon(s1,s3):
    str1 = (s1.get())  
    str3 = (s3.get())
    message = str1
    target= '"' + str3 + '"'
    link = "https://web.whatsapp.com/"
    chrome_options = Options()
    chrome_options.add_argument('--user-data-dir=./User_Data')
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 600)
    driver.get(link)
    driver.maximize_window()
    print("QR scanned")
    print(target)
    try:
        x_arg = '//span[contains(@title,' + target + ')]'
        try:
            time.sleep(20)
            group_title = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
            group_title.click()
        except:
            print("contact not found")
            time.sleep(4)
        input_box = driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
        input_box.send_keys(str1)
        input_box.send_keys(Keys.ENTER)
        print("Message sent successfuly")
        time.sleep(1)
    except NoSuchElementException:
        return
    driver.close()
    return

def sendall(s1): 
    str1 = (s1.get())
    for i in unsaved_Contacts:
        message=str1
        link = "https://wa.me/" + str(i)
        chrome_options = Options()
        chrome_options.add_argument('--user-data-dir=./User_Data')
        driver = webdriver.Chrome(options=chrome_options)
        wait = WebDriverWait(driver, 600)
        driver.get(link)
        driver.maximize_window()
        print("QR scanned")
        try:
            time.sleep(20)
            input_box = driver.find_elements_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')[0]
            for ch in message:
                if ch == "\n":
                    ActionChains(browser).key_down(Keys.SHIFT).key_down(Keys.ENTER).key_up(Keys.ENTER).key_up(Keys.SHIFT).key_up(Keys.BACKSPACE).perform()
                else:
                    input_box.send_keys(ch)
            input_box.send_keys(Keys.ENTER)
            print("Message sent successfuly")
        except NoSuchElementException:
            print("Failed to send message")
        driver.close()
    return

string1 = StringVar()  
string2 = StringVar()
string3 = StringVar()

mno = Label(top, text = "Enter Message").place(x = 30,y = 50)
msg = Label(top, text = "Enter Mobile No(91xxxxxxxxxx)").place(x = 30, y = 90)
rule = Label(top, text = "Mobile no should have prefix of country code e.g. 91 for India").place(x = 30, y = 130)
cname = Label(top, text = "Enter Contact Name").place(x = 30, y = 170)
ques = Label(top, text = "Want to send to multiple mobile numbers?? If yes, then").place(x = 30, y = 210)
fileup = Label(top, text = "Select file(.xlsx) to upload").place(x = 30, y = 250)
quitapp = Label(top, text = "Click on QUIT to cancel the App").place(x = 30, y = 290)


txt1 = Entry(top,textvariable=string1).place(x = 250, y = 50)  
txt2 = Entry(top,textvariable=string2).place(x = 250, y = 90)
txt3 = Entry(top,textvariable=string3).place(x = 250, y = 170)

sendno = partial(sendno, string1, string2)
sendcon = partial(sendcon, string1, string3)
sendall = partial(sendall, string1)

sendtono = Button(top, text = "Send to Number",command = sendno,activebackground = "green", activeforeground = "white").place(x =400, y = 90)
sendtocon = Button(top, text = "Send to Contact",command = sendcon,activebackground = "green", activeforeground = "white").place(x = 400, y = 170)
sendtoall = Button(top, text = "Send to All",command = sendall,activebackground = "green", activeforeground = "white").place(x = 400, y = 250)
selbtn = Button(top, text = "Select File",command = openfile,activebackground = "green", activeforeground = "white").place(x = 250, y = 250)
cancelbtn = Button(top, text = "QUIT",command = top.destroy,activebackground = "green", activeforeground = "white").place(x = 250, y = 290)
   
top.mainloop()  
