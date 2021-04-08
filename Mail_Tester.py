from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys

import smtplib, socket, sys, os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email import encoders

import time
import sys
import re
from datetime import datetime
import bs4 as bs
import argparse
import requests
import configparser

try:
    config = configparser.ConfigParser()
    config.read('NotasMail.ini')
except:
    print("No se encontro el archivo de configuración")


def enviar_correo(supermensaje):
    # Nos conectamos al servidor SMTP de mail
    try:
        smtpserver = smtplib.SMTP(config['SMTP']['HOST'], config['SMTP']['PORT'])
        smtpserver.ehlo()
        smtpserver.starttls()
        smtpserver.ehlo()
        print( "Conexion exitosa con mail" + '\n')
    except(socket.gaierror, socket.error, socket.herror, smtplib.SMTPException):
        os.system("cls")
        print( "Fallo en la conexion con mail")
        

    # logeamos el ususario
    try:
        user = config['SMTP']['USER']
        pwd = config['SMTP']['PASS']
        smtpserver.login(user, pwd)


        print("Autentificación correcta" + "\n")
    except smtplib.SMTPException:
        os.system("cls")
        print( "Autentificación incorrecta" + "\n")
        smtpserver.close()
        

    # configuramos correo
    to_list = []
    for key,val in config.items('MAILS'):
        to_list.append(val)
    to_addr = to_list


    print("Se enviara correo a: " + str(to_addr) + '\n' )

    From = config['MAIL_CFG']['from']
    mime_message=MIMEMultipart(_subtype='mixed')
    message = supermensaje
    mime_message["From"] = From
    mime_message["To"] = ', '.join(to_addr)
    mime_message["Subject"] = config['MAIL_CFG']['subject']
    mime_message.attach(MIMEText(message, 'plain'))


    try:
        smtpserver.sendmail(from_addr=From, to_addrs=to_addr, msg=mime_message.as_string())
        print( "El correo se envio correctamente:" + "\n")
    except smtplib.SMTPException:
        print( "El correo no pudo ser enviado" + "\n")
        smtpserver.close()
    
    smtpserver.close()


fecha = time.strftime("%d/%m/%y")
browser = webdriver.Chrome(executable_path=r'C:\webdriver\chromedriver.exe')

browser.get('https://www.mail-tester.com/')
browser.maximize_window()
browser.switch_to.default_content()

cc = browser.find_element_by_class_name('cc-compliance')
browser.execute_script("arguments[0].click();", cc)

delay = 3 # seconds

while True:
    try:
        wait = WebDriverWait(browser, 10)

        wait.until(EC.visibility_of_element_located((By.ID, "lang_selected")))

        email = browser.find_element_by_id("email")

        if email.get_attribute('value') != "":

            mail_to_test = email.get_attribute('value')

            break

    except TimeoutException:

        print ("Loading took too much time!-Try again")

        break

print ("se enviara un mail a: " + str(mail_to_test))

###enviamos el email por mailWizz 
sendMail = config['MAIL_CFG']['sendMail']
requests.get(sendMail+str(mail_to_test))

time.sleep(25)

log = browser.find_element_by_id('submitbutton')
browser.execute_script("arguments[0].click();", log)

time.sleep(120)



score = browser.find_element_by_class_name("score")
score = str(score.text)

os.system('cls')

print('El puntaje obtenido fue: ' + score)


puntaje = score.split('/')[0]
puntaje = float(puntaje)


link = browser.find_element_by_class_name("permalink-input")
link = link.text

print('El link para ver el resultado es: ' + link)

browser.close()

### aviso nota ###
nota_min = config.getint("NOTA", "min")
if puntaje <= nota_min:
    mensaje = ('El resultado de la prueba de mail tester, realizada el dia ' + fecha + ', dio de resultado: ' + score + '\nSe adjunta link: \n' + link)
    enviar_correo(mensaje)
    sys.exit(0)

else:
    sys.exit(0)
    


