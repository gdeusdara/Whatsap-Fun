# Note: For proper working of this Script Good and Uninterepted Internet Connection is Required
# Keep all contacts unique
# Can save contact with their phone Number

# Import required packages
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import datetime
import time
import openpyxl as excel

# function to read contacts from a text file
def readContacts(fileName):
    lst_name = []
    lst_number = []
    file = excel.load_workbook(fileName)
    sheet = file.active
    firstCol = sheet['A']
    secCol = sheet['B']
    for cell in range(len(firstCol)):
        contact_name = str(firstCol[cell].value)
        contact_name = "\"" + contact_name.strip() + "\""
        contact_number = str(secCol[cell].value)
        contact_number = "\"" + contact_number + "\""

        lst_name.append(contact_name)
        lst_number.append(contact_number)
    return lst_name, lst_number

# Target Contacts, keep them in double colons
# Not tested on Broadcast
target_names, target_numbers = readContacts("contacts.xlsx")

# can comment out below line
print(target_names, target_numbers)

# Driver to open a browser
driver = webdriver.Chrome()

#link to open a site
driver.get("https://web.whatsapp.com/")

# 10 sec wait time to load, if good internet connection is not good then increase the time
# units in seconds
# note this time is being used below also
wait = WebDriverWait(driver, 10)
wait5 = WebDriverWait(driver, 5)
input("Scan the QR code and then press Enter")

# Message to send list
# 1st Parameter: Hours in 0-23
# 2nd Parameter: Minutes
# 3rd Parameter: Seconds (Keep it Zero)
# 4th Parameter: Message to send at a particular time
# Put '\n' at the end of the message, it is identified as Enter Key
# Else uncomment Keys.Enter in the last step if you dont want to use '\n'
# Keep a nice gap between successive messages
# Use Keys.SHIFT + Keys.ENTER to give a new line effect in your Message
msgToSend = [
    [13, 16, 0,
        '''
     Querid@ artista e produtor(a) cultural!  ÚLTIMOS DIAS pra se inscrever no curso completo Viver de Arte! Aproveite essa oferta exclusiva: foram muitos pedidos da classe artística, então mantivemos o BÔNUS DA lista com 15 mil empresas patrocinadoras pela lei rouanet divididas por estado). https://viverdearte.arteemcurso.com/inscricoes-abertas
     As inscrições se encerram segunda-feira, dia 08 de junho, 23:59, e não sabemos quando abriremos novamente e o preço vai subir! O curso conta com 10h gravadas, suporte diário (inclusive ajuda pra responder diligências) e lives quinzenais comigo! Um abraço. Te espero no curso Viver de Arte!\n
     '''
    ],
    [12, 51, 0,
        '''
     Olá, artista e produtor(a) cultural!  ÚLTIMOS DIAS pra se inscrever no curso completo Viver de Arte! Aproveite essa oferta exclusiva: foram muitos pedidos da classe artística, então mantivemos o BÔNUS DA lista com 15 mil empresas patrocinadoras pela lei rouanet divididas por estado). https://viverdearte.arteemcurso.com/inscricoes-abertas
     As inscrições serão encerradas na segunda-feira, dia 08 de junho, 23:59, e não sabemos quando abriremos novamente e o preço vai subir! O curso conta com 10h gravadas, suporte diário (inclusive ajuda pra responder diligências) e lives quinzenais comigo! Um abraço. Te espero no curso Viver de Arte!\n
        '''
    ],
    [12, 51, 0,
        '''
     Olá novamente, querido artista e produtor(a) cultural!  ÚLTIMOS DIAS pra se inscrever no curso completo Viver de Arte! Aproveite essa oferta exclusiva: foram muitos pedidos da classe artística, então mantivemos o BÔNUS DA lista com 15 mil empresas patrocinadoras pela lei rouanet divididas por estado). https://viverdearte.arteemcurso.com/inscricoes-abertas
     As inscrições se encerram segunda-feira, dia 08 de junho, 23:59. Nós não sabemos quando abriremos novamente, e o preço vai subir! O curso conta com *10h gravadas, suporte diário (inclusive ajuda pra responder diligências) e lives quinzenais comigo!* Um abraço. Te espero no curso Viver de Arte!\n
        '''
    ],
    [12, 51, 0,
        '''
     Como vai? Querid@ artista e produtor(a) cultural!  ÚLTIMOS DIAS pra se inscrever no curso completo Viver de Arte! Aproveite essa oferta exclusiva: foram muitos pedidos da classe artística, então mantivemos o BÔNUS DA lista com 15 mil empresas patrocinadoras pela lei rouanet divididas por estado). https://viverdearte.arteemcurso.com/inscricoes-abertas
     Você pode se inscrever até segunda-feira, dia 08 de junho, 23:59. Após isso, não sabemos quando abriremos vagas novamente, e o preço vai subir! O curso conta com 10h gravadas, suporte diário (inclusive ajuda pra responder diligências) e lives quinzenais comigo! Um abraço. Te espero no curso Viver de Arte!\n
        '''
    ],
    [12, 51, 0,
        '''
     Alô alô, meu querid@!  ÚLTIMOS DIAS pra se inscrever no *curso completo* Viver de Arte! Aproveite essa oferta exclusiva: foram muitos pedidos da classe artística, então mantivemos o BÔNUS DA lista com 15 mil empresas patrocinadoras pela lei rouanet divididas por estado). https://viverdearte.arteemcurso.com/inscricoes-abertas
     As inscrições se encerram segunda-feira, dia 08 de junho, 23:59, e não sabemos quando abriremos novamente e o preço vai subir! O curso vai te fornecer 10h de aulas gravadas, suporte diário (inclusive ajuda pra responder diligências) e lives quinzenais comigo! Um abraço. Estarei te esperando no curso Viver de Arte!\n
        '''
    ],
    [12, 51, 0,
        '''
     Hoje é um excelente dia para te lembrar que estamos nos  ÚLTIMOS DIAS pra se inscrever no curso completo Viver de Arte! Aproveite essa oferta exclusiva: foram muitos pedidos da classe artística, então mantivemos o BÔNUS DA lista com 15 mil empresas patrocinadoras pela lei rouanet divididas por estado). https://viverdearte.arteemcurso.com/inscricoes-abertas
     É possível se inscrever até segunda-feira, dia 08 de junho, 23:59. Nós não sabemos quando abriremos o curso novamente, e o preço vai subir! O curso conta com 10h gravadas, suporte diário (inclusive ajuda pra responder diligências) e lives quinzenais comigo! Um abraço. Te espero no curso Viver de Arte!\n
        '''
    ],
    [12, 51, 0,
        '''
     Oi, novamente! Venho te lembrar que estes são os ÚLTIMOS DIAS pra se inscrever no curso completo Viver de Arte! Aproveite essa oferta exclusiva: foram muitos pedidos da classe artística, então mantivemos o BÔNUS DA lista com 15 mil empresas patrocinadoras pela lei rouanet divididas por estado). https://viverdearte.arteemcurso.com/inscricoes-abertas
     As inscrições se encerram segunda-feira, dia 08 de junho, 23:59, e não sabemos quando abriremos novamente e o preço vai subir! Você ainda tem a chance de contar com 10h de aulas gravadas, um suporte diário (inclusive ajuda pra responder diligências) e lives quinzenais comigo! Um abraço. Te espero no curso Viver de Arte!\n
        '''
    ],
    [12, 51, 0,
        '''
     Olá, olá, artista e produtor(a) cultural! estes são os ÚLTIMOS DIAS pra se inscrever no curso completo Viver de Arte! Aproveite essa oferta exclusiva: foram muitos pedidos da classe artística, então mantivemos o BÔNUS DA lista com 15 mil empresas patrocinadoras pela lei rouanet divididas por estado). https://viverdearte.arteemcurso.com/inscricoes-abertas
     As inscrições estarão abertas até segunda-feira, dia 08 de junho, 23:59, e não sabemos quando abriremos novamente e o preço vai subir! O curso conta com 10h gravadas, suporte diário (inclusive ajuda pra responder diligências) e lives quinzenais comigo! Um abraço. Te espero no curso Viver de Arte!\n
        '''
    ],
    [12, 51, 0,
        '''
     Prezado artista e produtor(a) cultural!  ÚLTIMOS DIAS pra se inscrever no curso completo Viver de Arte! Aproveite essa oferta exclusiva: foram muitos pedidos da classe artística, então mantivemos o BÔNUS DA lista com 15 mil empresas patrocinadoras pela lei rouanet divididas por estado). https://viverdearte.arteemcurso.com/inscricoes-abertas
     Até segunda-feira, dia 08 de junho, 23:59, você ainda poderá se inscrever. Aproveite, pois nós não sabemos quando abriremos novamente e o preço vai subir! O curso conta com 10h gravadas, suporte diário (inclusive ajuda pra responder diligências) e lives quinzenais comigo! Um abraço. Te espero no curso Viver de Arte!\n
        '''
    ],
    [12, 51, 0,
        '''
     Passei aqui para te avisar que nós estamos oficialmente nos  ÚLTIMOS DIAS pra se inscrever no curso completo Viver de Arte! Aproveite essa oferta exclusiva: foram muitos pedidos da classe artística, então mantivemos o BÔNUS DA lista com 15 mil empresas patrocinadoras pela lei rouanet divididas por estado). https://viverdearte.arteemcurso.com/inscricoes-abertas
     As inscrições poderão ser feitas até segunda-feira, dia 08 de junho, 23:59, após isso, não sabemos quando abriremos novamente e o preço vai subir! O curso conta com *10h gravadas, suporte diário (inclusive ajuda pra responder diligências) e lives quinzenais comigo!* Um grande abraço! Te espero no curso Viver de Arte.\n
        '''
    ],
]

# Count variable to identify the number of messages to be sent
count = 0

sended = False

# Identify time
curTime = datetime.datetime.now()
curHour = curTime.time().hour
curMin = curTime.time().minute
curSec = curTime.time().second

# utility variables to tract count of success and fails
sended = True

success = 0
sNo = 1
failList = []

# Iterate over selected contacts
for index in range(len(target_names)):
# for target in targets:
    print(sNo, ". Target is: " + target_names[index])
    sNo+=1
    try:
        # Select the target_names[index]
        # x_arg = '//span[@title=' + target_names[index] + ' and contains(@class, "_3ko75")]'
        x_arg = '//span[@class="_357i8" and //span[@title=' + target_names[index] + ' and contains(@class, "_3ko75")]]'
        print(x_arg)
        # try:
        #     wait5.until(EC.presence_of_element_located((
        #         By.XPATH, x_arg
        #     )))
        # except:
        # If contact not found, then search for it
            # Select the Input Box
        inp_xpath = "//div[@contenteditable='true']"
        search_box = wait.until(EC.presence_of_element_located((
            By.XPATH, inp_xpath)))
        time.sleep(1)

        # Send message
        # taeget is your target_names[index] Name and msgToSend is you message
        search_box.clear()
        time.sleep(1)
        search_box.send_keys(target_numbers[index].replace('"', '')) # (Uncomment it if your msg doesnt contain '\n')
        # Link Preview Time, Reduce this time, if internet connection is Good
        # search_box.send_keys(Keys.ENTER)
        # searBoxPath = "//div[@data-tab='3']"
        # wait5.until(EC.presence_of_element_located((
        #     By.ID, "input-chatlist-search"
        # )))
        # inputSearchBox = driver.find_element_by_id("input-chatlist-search")
        # time.sleep(0.5)
        # # click the search button
        # driver.find_element_by_xpath('/html/body/div/div/div/div[2]/div/div[2]/div/button').click()
        # time.sleep(1)
        # inputSearchBox.clear()
        # inputSearchBox.send_keys(target_names[index][1:len(target_names[index]) - 1])
        # print('Target Searched')
        # # Increase the time if searching a contact is taking a long time
        # time.sleep(4)

        time.sleep(6)
        try:
            driver.find_element_by_xpath(x_arg).click()
            print("Target Successfully Selected")
            time.sleep(2)
        except:
            testNumber = target_numbers[index].replace('"', '')
            search_box.clear()
            time.sleep(1)
            search_box.send_keys(testNumber[3:])
            time.sleep(6)
            driver.find_element_by_xpath(x_arg).click()
            print("Target Successfully Selected")

        # Select the target_names[index]
        time.sleep(1)
        # Select the Input Box
        inp_xpath = '//div[@spellcheck="true"]'
        input_box = wait.until(EC.presence_of_element_located((
            By.XPATH, inp_xpath)))
        time.sleep(1)

        # Send message
        # taeget is your target_names[index] Name and msgToSend is you message
        input_box.send_keys(msgToSend[count][3]) # + Keys.ENTER (Uncomment it if your msg doesnt contain '\n')
        # Link Preview Time, Reduce this time, if internet connection is Good
        time.sleep(6)
        # input_box.send_keys(Keys.ENTER)
        print("Successfully Send Message to : "+ target_names[index] + '\n')
        success+=1
        time.sleep(0.5)

        count+=1

        if count==len(msgToSend):
            count=0
    except:
        # If target_names[index] Not found Add it to the failed List
        f = open("errorList.txt", "a")
        f.write("%s - %s\n" % (target_names[index], target_numbers[index]))
        f.close()
        print("Cannot find Target: " + target_names[index] + " " + target_numbers[index])
        failList.append(target_names[index])


print("\nSuccessfully Sent to: ", success)
print("Failed to Sent to: ", len(failList))
print(failList)
print('\n\n')
driver.quit()