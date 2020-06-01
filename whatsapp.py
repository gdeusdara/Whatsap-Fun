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
        contact_name = "\"" + contact_name + "\""
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
    [19, 33, 0,
        '''
        Peço a sua licença, artista e produtor cultural, pois você se cadastrou na Semana Viver de Arte! Aqui é o Flavio Nardelli, fundador da Arte em Curso.
        
        Está no ar a *Semana Viver de Arte* e já com muito engajamento, mais de 600 comentários de artistas do país, querendo descobrir como captar muitos R$ em patrocínios para seu projeto cultural. Você que é artista/produtor(a) cultural e quer conhecer ó ÚNICO JEITO de bombar sua carreira ou projeto sem depender da sorte ou de um investidor milionário, você PRECISA assistir às aulas! Então clica nos links abaixo, e escolha a aula que você ainda não viu:

        https://viverdearte.arteemcurso.com/aula-1

        https://viverdearte.arteemcurso.com/aula-2

        https://viverdearte.arteemcurso.com/aula-3

        Caso os links nãos estejam ativos, responda essa mensagem com um “ok” para ativá-los.
        Corre lá! Abs,\n
        '''
    ],
    [12, 51, 0,
        '''
        Peço a sua licença, artista e produtor cultural, pois você se cadastrou na Semana Viver de Arte! Aqui é o Flavio Nardelli, fundador da Arte em Curso.
        Está no ar a Semana Viver de Arte e já com muito engajamento, mais de 600 comentários de artistas do país, querendo descobrir como captar muitos R$ em patrocínios para seu projeto cultural. Você que é artista/produtor(a) cultural e quer conhecer ó ÚNICO JEITO de bombar sua carreira ou projeto sem depender da sorte ou de um investidor milionário, você PRECISA assistir às aulas! Então clica nos links abaixo, e escolha a aula que você ainda não viu:
        https://viverdearte.arteemcurso.com/aula-1
        https://viverdearte.arteemcurso.com/aula-2
        https://viverdearte.arteemcurso.com/aula-3

        Caso os links nãos estejam ativos, responda essa mensagem com um “ok” para ativá-los

        Confere lá! Abraço,\n
        '''
    ],
    [12, 51, 0,
        '''
        Boa noite! Aqui é o Flavio Nardelli, fundador da Arte em Curso.
        Você viu que está no ar a Semana Viver de Arte? E já com muito engajamento, mais de 600 comentários de artistas do país, querendo descobrir como captar muitos R$ em patrocínios para seu projeto cultural. Você que é artista/produtor(a) cultural e quer conhecer ó ÚNICO JEITO de bombar sua carreira ou projeto sem depender da sorte ou de um investidor milionário, você PRECISA assistir às aulas! Então clica nos links abaixo, e escolha a aula que você ainda não viu:
        https://viverdearte.arteemcurso.com/aula-1
        https://viverdearte.arteemcurso.com/aula-2
        https://viverdearte.arteemcurso.com/aula-3

        Responda “ok” a essa mensagem para que os links sejam ativados. Abs,\n
        '''
    ],
    
]

# Count variable to identify the number of messages to be sent
count = 0

sended = False

while not sended:

    # Identify time
    curTime = datetime.datetime.now()
    curHour = curTime.time().hour
    curMin = curTime.time().minute
    curSec = curTime.time().second

    # if time matches then move further
    if msgToSend[count][0]==curHour and msgToSend[count][1]==curMin and msgToSend[count][2]==curSec:
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
                x_arg = '//span[@title=' + target_names[index] + ' and contains(@class, "_1wjpf")]'
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

                time.sleep(2)
                try:
                    driver.find_element_by_xpath(x_arg).click()
                    print("Target Successfully Selected")
                    time.sleep(2)
                except:
                    testNumber = target_numbers[index].replace('"', '')
                    search_box.clear()
                    time.sleep(1)
                    search_box.send_keys(testNumber[3:])
                    time.sleep(2)
                    driver.find_element_by_xpath(x_arg).click()
                    print("Target Successfully Selected")

                # Select the target_names[index]
                time.sleep(1)
                # Select the Input Box
                inp_xpath = "//div[@spellcheck='true']"
                input_box = wait.until(EC.presence_of_element_located((
                    By.XPATH, inp_xpath)))
                time.sleep(1)

                # Send message
                # taeget is your target_names[index] Name and msgToSend is you message
                input_box.send_keys(msgToSend[count][3] + Keys.SPACE) # + Keys.ENTER (Uncomment it if your msg doesnt contain '\n')
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