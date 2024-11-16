# Python Script for Duplicate Deletion - Endpoints Beta Version

# Dependencies:
from selenium import webdriver
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import datetime 
import pandas as pd
import maskpass

# Initialization :
print(" \n Initiating Script ...... \n")
sleep(1)

# User Input 
username = input(' Enter Username : ')
password = maskpass.askpass(prompt=' Enter Password : ', mask="*")
csv_path=input(' Enter CSV Path : ')

# EPM Console URL:
url="https://in.epm.cyberark.com"


# Chrome Driver:
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(options=options)

try : 
    driver.maximize_window()
    driver.get(url)

    # Enter Username & Password
    print(" \n Attempt to login ...... \n")
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "login_input_username"))).send_keys(username)
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "login_input_password"))).send_keys(password)
    sleep(4)
    # Click on the submit button
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "login_button_submit"))).click()
    sleep(4)
    
    # Set Selection
    sleep(4)
    WebDriverWait(driver, 60).until(EC.visibility_of_element_located((By.XPATH,'//*[@title="Airtel_Production(bharti airtel limited ibm so account)"]'))).click()

    # Go to Endpoints -> My Computers :
    sleep(2)
    endpoint_tab= driver.find_element(By.XPATH,'//*[@title="Endpoints"]')
    # endpoint_tab.click()

    # click on arrow button
    arrow_btn =endpoint_tab.find_element(By.XPATH,'.//span[contains(@class,"epm-panelmenu-icon")]')
    arrow_btn.click()
    sleep(2)
    #Enter My Computers 
    WebDriverWait(endpoint_tab, 60).until(EC.element_to_be_clickable((By.XPATH,'//*[@title="My Computers"]'))).click()
    sleep(3)

    # Expand My Computers :
    WebDriverWait(driver,60).until(EC.element_to_be_clickable((By.ID,'close-menu-button'))).click()
    sleep(3)

    # Search for Hostname: 
    # frames switched:
    iframe = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.NAME, "frameName")))
    driver.switch_to.frame(iframe)

# # Search for Hostname :
        # Pandas Logic
    dataFrame=pd.read_csv(csv_path)
        # convert to string:
    dataFrame['Hostname']=dataFrame['Hostname'].astype(str)

    # print(dataFrame)
    # print("Length of Dataframe: ",len(dataFrame))

        # iteration of dataframe:
    for i in range(len(dataFrame)):
        WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "fCompNm"))).clear()
        host=dataFrame['Hostname'].values[i]
        # print(host)
        WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "fCompNm"))).send_keys(dataFrame['Hostname'].values[i])
        sleep(4)
        # Wait to search for hostname entry : 
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "ctl00_cphButtonFilter_btnFilter_lblText"))).click()
        sleep(3)

        # Get the duplicate entry details - for date & state comparison
        delElement0= WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@title="Click to Open Computer; Right-click for more options"]'))).text
        delElement0=delElement0.lstrip(" ")
        # print(delElement0)
        disconnectedElement0 = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "ctl00_cphData_MainGrid_cell_0_3"))).text
        

        # Exception handling in case we receive only single entry :
        try: 
            delElement1= WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@title="Click to Open Computer; Right-click for more options"]'))).text
            delElement1=delElement1.lstrip(" ")
            disconnectedElement1 =WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, "ctl00_cphData_MainGrid_cell_1_3"))).text
            # print(disconnectedElement1)
        except Exception as e:
            # print("Not Found")
            continue

        # Exception entry for triple entries : 
        try: 
            delElement2= WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, '//*[@title="Click to Open Computer; Right-click for more options"]'))).text
            delElement2=delElement2.lstrip(" ")
            disconnectedElement2 = WebDriverWait(driver, 2).until(EC.visibility_of_element_located((By.ID, "ctl00_cphData_MainGrid_cell_2_3"))).text
            
            # Determine all 3 entries - state & date to compare :
            state0=disconnectedElement0.split("(")[0]
            lastSeen0=disconnectedElement0.split("(")[1]
            lastSeen0=lastSeen0.replace(')','')
            if(state0=="Alive " or state0=="Upgrading "):
                complt_date0= lastSeen0[20:]
            else:
                complt_date0= lastSeen0[11:]
            format = "%d-%b-%y %H:%M:%S"
            complt_date0=datetime.datetime.strptime(complt_date0,format)
            
            state1=disconnectedElement1.split("(")[0]
            lastSeen1=disconnectedElement1.split("(")[1]
            lastSeen1=lastSeen1.replace(')','')
            if(state1=="Alive " or state1=="Upgrading "):
                 complt_date1= lastSeen1[20:]
            else:
                 complt_date1= lastSeen1[11:]
            format = "%d-%b-%y %H:%M:%S"
            complt_date1=datetime.datetime.strptime(complt_date1,format)

            state2=disconnectedElement2.split("(")[0]
            lastSeen2=disconnectedElement2.split("(")[1]
            lastSeen2=lastSeen2.replace(')','')
            if(state2=="Alive " or state2=="Upgrading "):
                 complt_date2= lastSeen2[20:]
            else:
                 complt_date2= lastSeen2[11:]
            format = "%d-%b-%y %H:%M:%S"
            complt_date2=datetime.datetime.strptime(complt_date2,format)

            # Compare all 3 entries date & state & delete 1st duplicate value :
            # print(complt_date0)
            if(host == delElement0 and state0=="Disconnected " and (complt_date0<complt_date1 and complt_date0<complt_date2)):
                    # Deletion of Disconnected Hostname: Right click:
                    disconnectedElement = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "ctl00_cphData_MainGrid_cell_0_3")))
                    #print(disconnectedElement)
                    actions = ActionChains(driver)
                    # Action:Mouseup
                    actions.context_click(disconnectedElement)
                    actions.release()
                    actions.perform()
                    # Deletion of the Entry: 
                    sleep(4)
                    # Hover to - Delete :
                    table= WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "ctl00_cphEnd_menuComputers_5")))
                    nobr_del=table.find_element(By.TAG_NAME,"nobr")
                    ActionChains(driver).move_to_element(nobr_del).perform()
                    # print("Delete Hovered!!")
                    # Click on - Delete selected:
                    delSelct=WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "ctl00_cphEnd_menuComputers_6")))
                    delSelct.click()
                    sleep(4)
                    # print("Delete Selected")
                    # Click on Delete Selected:
                    WebDriverWait(driver,30).until(EC.visibility_of_element_located((By.ID, "btnConfirmDlgOk"))).click()
                    sleep(5)

            elif(host == delElement1 and state1=="Disconnected " and (complt_date1<complt_date0 and complt_date1<complt_date2)):
                    # Deletion of Disconnected Hostname: Right click:
                    disconnectedElement = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "ctl00_cphData_MainGrid_cell_1_3")))
                    #print(disconnectedElement)
                    actions = ActionChains(driver)
                    # Action:Mouseup
                    actions.context_click(disconnectedElement)
                    actions.release()
                    actions.perform()
                    # Deletion of the Entry: 
                    sleep(4)
                    # Hover to - Delete :
                    table= WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "ctl00_cphEnd_menuComputers_5")))
                    nobr_del=table.find_element(By.TAG_NAME,"nobr")
                    ActionChains(driver).move_to_element(nobr_del).perform()
                    # print("Delete Hovered!!")
                    # Click on - Delete selected:
                    delSelct=WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "ctl00_cphEnd_menuComputers_6")))
                    delSelct.click()
                    sleep(4)
                    # print("Delete Selected")
                    # Click on  Delete Selected:
                    WebDriverWait(driver,30).until(EC.visibility_of_element_located((By.ID, "btnConfirmDlgOk"))).click()
                    sleep(5)

            elif(host == delElement2 and state2=="Disconnected " and (complt_date2<complt_date0 and complt_date2<complt_date1)):
                    # Deletion of Disconnected Hostname: Right click:
                    disconnectedElement = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "ctl00_cphData_MainGrid_cell_2_3")))
                    #print(disconnectedElement)
                    actions = ActionChains(driver)
                    # Action:Mouseup
                    actions.context_click(disconnectedElement)
                    actions.release()
                    actions.perform()
                    # Deletion of the Entry: 
                    sleep(4)
                    # Hover to - Delete :
                    table= WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "ctl00_cphEnd_menuComputers_5")))
                    nobr_del=table.find_element(By.TAG_NAME,"nobr")
                    ActionChains(driver).move_to_element(nobr_del).perform()
                    # print("Delete Hovered!!")
                    # Click on - Delete selected:
                    delSelct=WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "ctl00_cphEnd_menuComputers_6")))
                    delSelct.click()
                    sleep(4)
                    # print("Delete Selected")
                    # Click on Cancel for Delete Selected:
                    WebDriverWait(driver,30).until(EC.visibility_of_element_located((By.ID, "btnConfirmDlgOk"))).click()
                    sleep(5)

            # After removing 1st duplicate ,search the hostname again & act as double duplicate:
            
            WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "fCompNm"))).clear()
            WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "fCompNm"))).send_keys(host)
            # Wait to search for hostname entry : 
            WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "ctl00_cphButtonFilter_btnFilter_lblText"))).click()
            sleep(5)
            disconnectedElement0 = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "ctl00_cphData_MainGrid_cell_0_3"))).text
            disconnectedElement1 =WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "ctl00_cphData_MainGrid_cell_1_3"))).text

        except:
             pass
    

        # Python logic to compare hostname entries & click on one with oldest last seen:
        
        #Hostname -1 : Last seen & date 
        state0=disconnectedElement0.split("(")[0]
        lastSeen0=disconnectedElement0.split("(")[1]
        lastSeen0=lastSeen0.replace(')','')
        if(state0=="Alive " or state0=="Upgrading "):
            complt_date0= lastSeen0[20:]
        else:
             complt_date0= lastSeen0[11:]

        format = "%d-%b-%y %H:%M:%S"
        complt_date0=datetime.datetime.strptime(complt_date0,format)

        # print(state0)
        # print(lastSeen0)
        # print(complt_date0)

        #Hostname -2 : Last seen & date 
        state1=disconnectedElement1.split("(")[0]
        lastSeen1=disconnectedElement1.split("(")[1]
        lastSeen1=lastSeen1.replace(')','')
        if(state1=="Alive " or state1=="Upgrading "):
            complt_date1= lastSeen1[20:]
        else:
             complt_date1= lastSeen1[11:]

        format = "%d-%b-%y %H:%M:%S"
        complt_date1=datetime.datetime.strptime(complt_date1,format)

        # print(state1)
        # print(lastSeen1)
        # print(complt_date1)

        # print(complt_date0)
        if(host == delElement0 and state0=="Disconnected " and (complt_date0<complt_date1 or state1 == "Alive " or state1=="Upgrading ")):
                # Deletion of Disconnected Hostname: Right click:
                disconnectedElement = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "ctl00_cphData_MainGrid_cell_0_3")))
                #print(disconnectedElement)
                actions = ActionChains(driver)
                # Action:Mouseup
                actions.context_click(disconnectedElement)
                actions.release()
                actions.perform()
                # Deletion of the Entry: 
                sleep(4)
                # Hover to - Delete :
                table= WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "ctl00_cphEnd_menuComputers_5")))
                nobr_del=table.find_element(By.TAG_NAME,"nobr")
                ActionChains(driver).move_to_element(nobr_del).perform()
                # print("Delete Hovered!!")
                # Click on - Delete selected:
                delSelct=WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "ctl00_cphEnd_menuComputers_6")))
                delSelct.click()
                sleep(4)
                # print("Delete Selected")
                # Click on Cancel for Delete Selected:
                WebDriverWait(driver,30).until(EC.visibility_of_element_located((By.ID, "btnConfirmDlgOk"))).click()
                sleep(5)

        elif(host == delElement1 and state1=="Disconnected " and (complt_date0>complt_date1 or state0 == "Alive " or state0=="Upgrading ")):
                # Deletion of Disconnected Hostname: Right click:
                disconnectedElement = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "ctl00_cphData_MainGrid_cell_1_3")))
                #print(disconnectedElement)
                actions = ActionChains(driver)
                # Action:Mouseup
                actions.context_click(disconnectedElement)
                actions.release()
                actions.perform()
                # Deletion of the Entry: 
                sleep(4)
                # Hover to - Delete :
                table= WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "ctl00_cphEnd_menuComputers_5")))
                nobr_del=table.find_element(By.TAG_NAME,"nobr")
                ActionChains(driver).move_to_element(nobr_del).perform()
                # print("Delete Hovered!!")
                # Click on - Delete selected:
                delSelct=WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "ctl00_cphEnd_menuComputers_6")))
                delSelct.click()
                sleep(4)
                # print("Delete Selected")
                # Click on Cancel for Delete Selected:
                WebDriverWait(driver,30).until(EC.visibility_of_element_located((By.ID, "btnConfirmDlgOk"))).click()
                sleep(5)
        else:
                continue 
   # Additional sleep to keep the screen awake:
    sleep(5)
    print("\n Thank You !!! \t Duplicates have been deleted sucessfully . \n You may close the window now.... \tCreated by: Shivam Arora , AUUID - 23168324 \n")
    
    # Exception Handling
except Exception as e:
    print(f"\n Script encountered an error !!  Please check the code  & run the script again .\n In case of any help , kindly connect with script developer : Shivam Arora , AUUID : 23168324  \n") 
    # print(e)

# Default Case:
finally : 
    # Quit the driver
    driver.quit()
    # Quit the driver
    driver.quit()
