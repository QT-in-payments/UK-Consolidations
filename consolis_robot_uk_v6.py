# -*- coding: utf-8 -*-
"""
Created on Wed Nov 14 15:43:35 2018

@author: quang.tran
"""

import win32com.client
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import pandas as pd
import time
import datetime
import numpy as np
import tkinter as tk



today = datetime.datetime.today().strftime('%Y-%m-%d')
yesterday = (datetime.datetime.today() - datetime.timedelta(1)).strftime('%Y-%m-%d')

master = tk.Tk()
master.title("Consolis Robot")
tk.Label(master, text = "Input Email: ", font = ("Helvetica", 9, "bold")).grid(row = 0, sticky = "e")
tk.Label(master, text = "Input Password: ", font = ("Helvetica", 9, "bold")).grid(row = 1, sticky = "e")
text = tk.Text(master, width = 65, height = 20)
text.grid(row = 7, column = 0, columnspan = 7, padx = 10, pady = 10)
e1 = tk.Entry(master, width = 20)
e2 = tk.Entry(master, width = 20, show = '*')
e1.grid(row = 0, column = 1)
e2.grid(row = 1, column = 1)
e1.insert(tk.END, '@fundingcircle.com')

#--------------------- Grab Login info for Bilcas ----------------------#
bilcas_email = e1.get()
bilcas_password = e2.get()

#----------------- Filepaths to save and load files --------------------#

consolidations_filepath = r"N:\Operations\Operations\BILCAS Processes\Consolidations Folder\Templates\Consolidations Macro Template - v2.xlsm"
combined_data_path = r"N:\Operations\Operations\BILCAS Processes\Consolidations Folder\Consolidations Data\Consolidations Data " + datetime.datetime.today().strftime('%d%m%Y') + ".xlsx"
old_loan_path = r"N:\Operations\Operations\BILCAS Processes\Consolidations Folder\Consolidations Data\Old Loan Consolis " + datetime.datetime.today().strftime('%d%m%Y') + ".csv"
new_loan_path = r"N:\Operations\Operations\BILCAS Processes\Consolidations Folder\Consolidations Data\New Loan Consolis " + datetime.datetime.today().strftime('%d%m%Y') + ".csv"
chrome_path = r"N:\Operations\Operations\BILCAS Processes\Payments Wizard\chromedriver.exe"

#---------------- Run VBA Macros separately via Python --------------------#

#Below 3 macros allow us to build the consolidations data sheet, which we use to compile data for all the Processes
#Run a macro the goes into FCA Backend, grab any additional funds required for any old cashfacs we're Settling
#And finally, update the loan amount within FCA Backend, which updates in Bilcas automatically

start_time = datetime.datetime.now()

incomplete_consolis = []


def runConsolidationsSheetMacro():
    xl = win32com.client.Dispatch("Excel.Application") #Loads excel application in background
    xl.Workbooks.Open(Filename = consolidations_filepath) #Opens the file with our 3 Macros
    xl.Visible = 1 #Makes excel screen visible so the macro works properly
    xl.Application.Run("'Consolidations Macro Template - v2.xlsm'!BuildConsolidationsSheet.FormatSheets") #Runs the macro. Name of macro is after exclamation mark
    xl.Application.Quit() #Quits excel
    text.insert(tk.END, "Step 1 complete. You can check the Consolidations Data " + datetime.datetime.today().strftime('%d%m%Y') + " file for accuracy. \n")
    start_time = datetime.datetime.now()
    
def runBackendUpdateMacro():
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Workbooks.Open(Filename = consolidations_filepath, ReadOnly = 1)
    xl.Visible = 1
    xl.Application.Run("'Consolidations Macro Template - v2.xlsm'!UpdateBackendLoanAmounts.UpdateBackend")
    xl.Application.Quit()
    text.insert(tk.END, "Step 4 complete. Loan amounts in backend have been updated. \n")

def grabAdditionalFundsRequired():
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Workbooks.Open(Filename = consolidations_filepath, ReadOnly = 1)
    xl.Visible = 1
    xl.Application.Run("'Consolidations Macro Template - v2.xlsm'!GrabAdditionalFundsToSettle.GrabAFRBackend")
    xl.Application.Quit()
    text.insert(tk.END, "Step 3 complete. All old loans have data on additional funds required. \n")
    
def settleInBackend():
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Workbooks.Open(Filename = consolidations_filepath, ReadOnly = 1)
    xl.Visible = 1
    xl.Application.Run("'Consolidations Macro Template - v2.xlsm'!SettleInBackend.settleLoan")
    xl.Application.Quit()
    text.insert(tk.END, "Step 8 complete. Loans have been settled in the Backend. \n")

#------------------------- Checks to make sure cashfacs are found in Exports/Paused page -----------------------------#

#This logs into the Bilcas page, navigates to exports, then paused page
#Once there, it looks at all the new cashfac ids in the consolidations data sheet from step 1,
#And assigns a true or false note in the 'Exports/Paused Good?' column, with true meaning we can find the new cashfac id in the page_source

def check_bilcas_exports(excel_sheet_name, save_name):

    options = webdriver.ChromeOptions()
    browser = webdriver.Chrome(chrome_path, options = options) #Loads chrome
    browser.get('https://bilcas.fundingcircle.co.uk/users/sign_in') #Navigates to Bilcas page

    browser.maximize_window() #Maximizes window because sometimes you can't click on elements if screen is too small
    email = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, "user_email"))) #browser.find_element_by_id('user_email') #Finds the email input box
    password = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'user_password'))) #Finds the password input box

    email.send_keys(e1.get())
    password.send_keys(e2.get())

    browser.find_element_by_xpath('//*[@id="new_user"]/div[3]/input').click()

    def scroll_down_pages():
        #The below code is to SCROLL DOWN the bilcas page because sometimes there are a lot of transactions that won't load unless you scroll down the page
        last_height = browser.execute_script("return document.body.scrollHeight")
        while True:
            #Scrolls to bottom of the page
            browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(0.825)
            #Gets the new length of the page after scrolling down
            new_height = browser.execute_script("return document.body.scrollHeight")
            #If you scroll to the end and there's no more data, this code will stop
            if new_height == last_height:
                break
            #updates the last height until last height is equal to new height
            last_height = new_height

    consolidations_df = pd.read_excel(combined_data_path, sheet_name = excel_sheet_name)

    browser.get('https://bilcas.fundingcircle.co.uk/bacs_export_bank_files')
    browser.find_element_by_xpath('//*[@id="exportfile-status"]/tbody/tr[3]/td[1]/a').click()

    column_names = ['Holder Reference', 'Type', 'Payee Sort Code', 'Payee Account Number', 'Payee Name', 'Transaction Ref', 'Amount', 'Status', 'Last Updated', 'Action']
    scroll_down_pages()
    time.sleep(1)

    try:
        #Selenium hands page source to beautiful soup
        soup_level = BeautifulSoup(browser.page_source, 'lxml')
        #Find table in bilcas
        table = soup_level.find_all('table')
        #Reads table
        df = pd.read_html(str(table), header = 0)
        table = df[0]
        #Gives table columns names as defined above
        table.columns = column_names
        #Pushes table into a Pandas dataframe
        dataframe = pd.DataFrame(table)

        consolidations_df['Exports/Paused Good?'] = np.bool

    except NoSuchElementException:
                pass

    for index, row in consolidations_df.iterrows():
        #Identifies if cashfac is in holder_reference column
        id_count = dataframe['Holder Reference'].str.contains(row['New Cashfac IDs']).any()

        consolidations_df.at[index, 'Exports/Paused Good?'] = id_count

    consolidations_df.to_csv(r"N:\Operations\Operations\BILCAS Processes\Consolidations Folder\Consolidations Data" + r"\\" + save_name + " " + datetime.datetime.today().strftime('%d%m%Y') + ".csv", index = None)

    browser.close() 
    

def run_exports_check():
    check_bilcas_exports("New Cashfac Data", "New Loan Consolis")
    check_bilcas_exports("Old Cashfac Data", "Old Loan Consolis")
    text.insert(tk.END, "Step 2 complete. All loans have been checked in the exports/paused page. \n")
    

#-------------------------- End code to check exports ---------------------------------------#

def post_new_cashfacs():

    consolidations_df = pd.read_csv(new_loan_path)
    options = webdriver.ChromeOptions()
    browser = webdriver.Chrome(chrome_path, options = options)
    browser.get('https://bilcas.fundingcircle.co.uk/users/sign_in')

    browser.maximize_window() #Maximizes window because sometimes you can't click on elements if screen is too small
    email = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, "user_email"))) #browser.find_element_by_id('user_email') #Finds the email input box
    password = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'user_password'))) #Finds the password input box

    email.send_keys(e1.get())
    password.send_keys(e2.get())

    browser.find_element_by_xpath('//*[@id="new_user"]/div[3]/input').click()
    i = 0
    
    for index, row in consolidations_df.iterrows():

        #----------------- LOGIC FOR POSTING COMMENTS WITH MULTIPLE CASHFACS ------------------#
        if (
            pd.isnull(row['Old Cashfac ID 10']) == True and
            pd.isnull(row['Old Cashfac ID 9']) == True and
            pd.isnull(row['Old Cashfac ID 8']) == True and
            pd.isnull(row['Old Cashfac ID 7']) == True and
            pd.isnull(row['Old Cashfac ID 6']) == True and
            pd.isnull(row['Old Cashfac ID 5']) == True and
            pd.isnull(row['Old Cashfac ID 4']) == True and
            pd.isnull(row['Old Cashfac ID 3']) == True and
            pd.isnull(row['Old Cashfac ID 2']) == True
            ):
            comment = "Settling previous loan " + str(row['Old Cashfac ID 1']) + " with new loan"
        elif (
            pd.isnull(row['Old Cashfac ID 10']) == True and
            pd.isnull(row['Old Cashfac ID 9']) == True and
            pd.isnull(row['Old Cashfac ID 8']) == True and
            pd.isnull(row['Old Cashfac ID 7']) == True and
            pd.isnull(row['Old Cashfac ID 6']) == True and
            pd.isnull(row['Old Cashfac ID 5']) == True and
            pd.isnull(row['Old Cashfac ID 4']) == True and
            pd.isnull(row['Old Cashfac ID 3']) == True
            ):
            comment = "Settling previous loans " + str(row['Old Cashfac ID 1']) + " and " + str(row['Old Cashfac ID 2']) + " with new loan"
        elif (
            pd.isnull(row['Old Cashfac ID 10']) == True and
            pd.isnull(row['Old Cashfac ID 9']) == True and
            pd.isnull(row['Old Cashfac ID 8']) == True and
            pd.isnull(row['Old Cashfac ID 7']) == True and
            pd.isnull(row['Old Cashfac ID 6']) == True and
            pd.isnull(row['Old Cashfac ID 5']) == True and
            pd.isnull(row['Old Cashfac ID 4']) == True
            ):
            comment = "Settling previous loans " + str(row['Old Cashfac ID 1']) + ", " + str(row['Old Cashfac ID 2']) + " and " + str(row['Old Cashfac ID 3']) + " with new loan"
        elif (
            pd.isnull(row['Old Cashfac ID 10']) == True and
            pd.isnull(row['Old Cashfac ID 9']) == True and
            pd.isnull(row['Old Cashfac ID 8']) == True and
            pd.isnull(row['Old Cashfac ID 7']) == True and
            pd.isnull(row['Old Cashfac ID 6']) == True and
            pd.isnull(row['Old Cashfac ID 5']) == True
            ):
            comment = "Settling previous loans " + str(row['Old Cashfac ID 1']) + ", " + str(row['Old Cashfac ID 2']) + ", " + str(row['Old Cashfac ID 3']) + " and " + str(row['Old Cashfac ID 4']) + " with new loan"
        elif (
            pd.isnull(row['Old Cashfac ID 10']) == True and
            pd.isnull(row['Old Cashfac ID 9']) == True and
            pd.isnull(row['Old Cashfac ID 8']) == True and
            pd.isnull(row['Old Cashfac ID 7']) == True and
            pd.isnull(row['Old Cashfac ID 6']) == True
            ):
            comment = "Settling previous loans " + str(row['Old Cashfac ID 1']) + ", " + str(row['Old Cashfac ID 2']) + ", " + str(row['Old Cashfac ID 3']) + ", " + str(row['Old Cashfac ID 4']) + " and " + str(row['Old Cashfac ID 5']) + " with new loan"
        elif (
            pd.isnull(row['Old Cashfac ID 10']) == True and
            pd.isnull(row['Old Cashfac ID 9']) == True and
            pd.isnull(row['Old Cashfac ID 8']) == True and
            pd.isnull(row['Old Cashfac ID 7']) == True
            ):
            comment = "Settling previous loans " + str(row['Old Cashfac ID 1']) + ", " + str(row['Old Cashfac ID 2']) + ", " + str(row['Old Cashfac ID 3']) + ", " + str(row['Old Cashfac ID 4']) + ", " + str(row['Old Cashfac ID 5']) + " and " + str(row['Old Cashfac ID 6']) + " with new loan"
        elif (
            pd.isnull(row['Old Cashfac ID 10']) == True and
            pd.isnull(row['Old Cashfac ID 9']) == True and
            pd.isnull(row['Old Cashfac ID 8']) == True
            ):
            comment = "Settling previous loans " + str(row['Old Cashfac ID 1']) + ", " + str(row['Old Cashfac ID 2']) + ", " + str(row['Old Cashfac ID 3']) + ", " + str(row['Old Cashfac ID 4']) + ", " + str(row['Old Cashfac ID 5']) + ", " + str(row['Old Cashfac ID 6']) + " and " + str(row['Old Cashfac ID 7']) + " with new loan"
        elif (
            pd.isnull(row['Old Cashfac ID 10']) == True and
            pd.isnull(row['Old Cashfac ID 9']) == True
            ):
            comment = "Settling previous loans " + str(row['Old Cashfac ID 1']) + ", " + str(row['Old Cashfac ID 2']) + ", " + str(row['Old Cashfac ID 3']) + ", " + str(row['Old Cashfac ID 4']) + ", " + str(row['Old Cashfac ID 5']) + ", " + str(row['Old Cashfac ID 6']) + ", " + str(row['Old Cashfac ID 7']) + " and " + str(row['Old Cashfac ID 8']) + " with new loan"
        elif (
            pd.isnull(row['Old Cashfac ID 10']) == True
            ):
            comment = "Settling previous loans " + str(row['Old Cashfac ID 1']) + ", " + str(row['Old Cashfac ID 2']) + ", " + str(row['Old Cashfac ID 3']) + ", " + str(row['Old Cashfac ID 4']) + ", " + str(row['Old Cashfac ID 5']) + ", " + str(row['Old Cashfac ID 6']) + ", " + str(row['Old Cashfac ID 7']) + ", " + str(row['Old Cashfac ID 8']) + " and " + str(row['Old Cashfac ID 9']) + " with new loan"
        else:
            comment = "Settling previous loans " + str(row['Old Cashfac ID 1']) + ", " + str(row['Old Cashfac ID 2']) + ", " + str(row['Old Cashfac ID 3']) + ", " + str(row['Old Cashfac ID 4']) + ", " + str(row['Old Cashfac ID 5']) + ", " + str(row['Old Cashfac ID 6']) + ", " + str(row['Old Cashfac ID 7']) + ", " + str(row['Old Cashfac ID 8']) + ", " + str(row['Old Cashfac ID 9']) + " and " + str(row['Old Cashfac ID 10']) + " with new loan"

        #----------------- END LOGIC FOR COMMENTS -----------------#

        #---------- Check to make sure we're not posting any transactions where there is no settlement amount --------#
        if (
            pd.isnull(row['Total Amount to Settle']) == False and
            row['Exports/Paused Good?'] == True and
            row['Loan Status Late Check'] == 0 and
            row['Backend Status'] == 0
            ):
            
            i += 1
            
            browser.get('https://bilcas.fundingcircle.co.uk/customers/borrowers')

            time.sleep(1.5)

            #Clears the search box
            browser.find_element_by_xpath('//*[@id="customer-table_filter"]/label/input').clear()

            #Searches for cashfac we need in borrower page
            browser.find_element_by_xpath('//*[@id="customer-table_filter"]/label/input').send_keys(row['New Cashfac IDs'])

            #Clicks the search button
            browser.find_element_by_xpath('//*[@id="customer-table_filter"]/label/i').click()

            time.sleep(3)

            #Clicks the view page button for borrowers
            browser.find_element_by_xpath('//*[@id="customer-table"]/tbody/tr/td[8]/a/i').click()

            soup_level = BeautifulSoup(browser.page_source, 'lxml')

            borrower_h1 = soup_level.find_all('h1')[0].text
            borrower_name = borrower_h1.split('\n')[1]

            #Clicks to create new transaction
            browser.find_element_by_xpath('/html/body/div[2]/div[2]/div/div[1]/div/a[2]').click()

            #Clicks on manual tab in new transactions
            browser.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div[2]/ul/li[3]/a').click()


            #Identifies comment box
            comment_box = browser.find_element_by_xpath('//*[@id="create_authorization_request_notes"]')

            #Clicks to debit cash
            Select(browser.find_element_by_xpath('//*[@id="account-to-debit"]')).select_by_visible_text(borrower_name + "'s cash account")

            #Clicks to credit Control
            Select(browser.find_element_by_xpath('//*[@id="account-to-credit"]')).select_by_visible_text('CON Control Account')

            #Inputs amount we want to post
            WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'manual-amount'))).send_keys(str(row['Total Amount to Settle']))

            comment_box.send_keys(comment)

            #Clicks save button in manual transactions page
            browser.find_element_by_xpath('//*[@id="new_create_authorization_request"]/div[7]/input').click()
        else:
            print("Cashfac " + row["New Cashfac IDs"] + " does not meet all the criteria to post a transaction.")
            incomplete_consolis.append(row['New Cashfac IDs'])
            pass
        
    text.insert(tk.END, "Step 5 complete. " + str(i) + " transactions should've posted. Cash for loan settlements have been debited. \n")
    print(incomplete_consolis)

#----------------- Code to post settlements in Bilcas for old loans that need to be consolidated ----------------------#

def post_old_cashfacs():

    consolidations_df = pd.read_csv(old_loan_path)
    options = webdriver.ChromeOptions()
    browser = webdriver.Chrome(chrome_path, options = options)
    browser.get('https://bilcas.fundingcircle.co.uk/users/sign_in')

    browser.maximize_window() #Maximizes window because sometimes you can't click on elements if screen is too small
    email = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, "user_email"))) #browser.find_element_by_id('user_email') #Finds the email input box
    password = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'user_password'))) #Finds the password input box

    email.send_keys(e1.get())
    password.send_keys(e2.get())

    browser.find_element_by_xpath('//*[@id="new_user"]/div[3]/input').click()
    i = 0
    for index, row in consolidations_df.iterrows():

        #---------- Check to make sure we're not posting any transactions where there is no settlement amount --------#

        if (
            pd.isnull(row['Settlement Amount']) == False and
            row['Exports/Paused Good?'] == True and
            row['Status Check'] == 0 and
            row['Loan Exists?'] == True and
            row['Loan Status'] == 'loan: repaying' and
            row['Backend Status'] == 'live' and
            pd.isnull(row['Additional Funds Required Amount']) == False
            ):
            
            i += 1

            browser.get('https://bilcas.fundingcircle.co.uk/customers/borrowers')

            time.sleep(1.5)

            #Clears the search box
            browser.find_element_by_xpath('//*[@id="customer-table_filter"]/label/input').clear()

            #Searches for cashfac we need in borrower page
            browser.find_element_by_xpath('//*[@id="customer-table_filter"]/label/input').send_keys(row['Old Cashfac IDs'])

            #Clicks the search button
            browser.find_element_by_xpath('//*[@id="customer-table_filter"]/label/i').click()

            time.sleep(3)

            #Clicks the view page button for borrowers
            browser.find_element_by_xpath('//*[@id="customer-table"]/tbody/tr/td[8]/a/i').click()

            time.sleep(2.5)
            
            soup_level = BeautifulSoup(browser.page_source, 'lxml')
            
            cash_balance_h1 = soup_level.find_all('h3')[0].text
            amount = cash_balance_h1.split('\n')[2]

            consolidations_df.at[index, 'Bilcas Cash Balance'] = amount


            #--------------- Start section to post manual transaction --------------#
            borrower_h1 = soup_level.find_all('h1')[0].text
            borrower_name = borrower_h1.split('\n')[1]

            #Clicks to create new transaction
            browser.find_element_by_xpath('/html/body/div[2]/div[2]/div/div[1]/div/a[2]').click()

            #Clicks on manual tab in new transactions
            browser.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div[2]/ul/li[3]/a').click()


            #Identifies comment box
            comment_box = browser.find_element_by_xpath('//*[@id="create_authorization_request_notes"]')

            #Clicks to debit cash
            Select(browser.find_element_by_xpath('//*[@id="account-to-debit"]')).select_by_visible_text('CON Control Account')

            #Clicks to credit Control
            Select(browser.find_element_by_xpath('//*[@id="account-to-credit"]')).select_by_visible_text(borrower_name + "'s cash account")

            #Inputs amount we want to post
            WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'manual-amount'))).send_keys(str(row['Settlement Amount']))

            comment_box.send_keys('Settling with new loan ' + row['New Cashfac IDs'])

            #clicks save button
            browser.find_element_by_xpath('//*[@id="new_create_authorization_request"]/div[7]/input').click()
        else:
            print("Cashfac " + row["Old Cashfac IDs"] + " does not meet all the criteria to post a transaction.")
            pass
        
        for index, row in consolidations_df.iterrows():            
            if (
                pd.isnull(row['Settlement Amount']) == False and 
                row['Exports/Paused Good?'] == True and
                row['Status Check'] == 0 and
                row['Loan Exists?'] == True and
                row['Loan Status'] == 'loan: repaying' and
                row['Backend Status'] == 'live' and
                pd.isnull(row['Additional Funds Required Amount']) == False
                ):
                consolidations_df.at[index, 'Test if Run Amount'] = float(row['Bilcas Cash Balance']) - float(row['Additional Funds Required Amount'])
            else:
                pass
        
    consolidations_df.to_csv(old_loan_path, index = None)
        
    text.insert(tk.END, "Step 6 complete. " + str(i) + " transactions should've posted. Cash to settle old loans have been credited and bilcas cash balance has been grabbed. \n")

def post_test_if_run():

    consolidations_df = pd.read_csv(old_loan_path)
    options = webdriver.ChromeOptions()
    browser = webdriver.Chrome(chrome_path, options = options)
    browser.get('https://bilcas.fundingcircle.co.uk/users/sign_in')

    browser.maximize_window() #Maximizes window because sometimes you can't click on elements if screen is too small
    email = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, "user_email"))) #browser.find_element_by_id('user_email') #Finds the email input box
    password = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'user_password'))) #Finds the password input box

    email.send_keys(e1.get())
    password.send_keys(e2.get())

    browser.find_element_by_xpath('//*[@id="new_user"]/div[3]/input').click()
    i = 0
    for index, row in consolidations_df.iterrows():

         if(
            pd.isnull(row['Settlement Amount']) == False and
            row['Exports/Paused Good?'] == True and
            row['Status Check'] == 0 and
            row['Loan Exists?'] == True and
            row['Loan Status'] == 'loan: repaying' and
            row['Backend Status'] == 'live' and
            pd.isnull(row['Additional Funds Required Amount']) == False and
            row['Test if Run Amount'] < 0
            ):
             
            i += 1
            
            browser.get('https://bilcas.fundingcircle.co.uk/customers/borrowers')

            time.sleep(1.5)

            #Clears the search box
            browser.find_element_by_xpath('//*[@id="customer-table_filter"]/label/input').clear()

            #Searches for cashfac we need in borrower page
            browser.find_element_by_xpath('//*[@id="customer-table_filter"]/label/input').send_keys(row['Old Cashfac IDs'])

            #Clicks the search button
            browser.find_element_by_xpath('//*[@id="customer-table_filter"]/label/i').click()

            time.sleep(3)

            #Clicks the view page button for borrowers
            browser.find_element_by_xpath('//*[@id="customer-table"]/tbody/tr/td[8]/a/i').click()

            time.sleep(2.5)

            soup_level = BeautifulSoup(browser.page_source, 'lxml')

            borrower_h1 = soup_level.find_all('h1')[0].text
            borrower_name = borrower_h1.split('\n')[1]

            #Clicks to create new transaction
            browser.find_element_by_xpath('/html/body/div[2]/div[2]/div/div[1]/div/a[2]').click()

            #Clicks on manual tab in new transactions
            browser.find_element_by_xpath('/html/body/div[2]/div[2]/div/div/div[2]/ul/li[3]/a').click()

            #Identifies comment box
            comment_box = browser.find_element_by_xpath('//*[@id="create_authorization_request_notes"]')

            #Clicks to debit cash
            Select(browser.find_element_by_xpath('//*[@id="account-to-debit"]')).select_by_visible_text('CON Control Account')

            #Clicks to credit Control
            Select(browser.find_element_by_xpath('//*[@id="account-to-credit"]')).select_by_visible_text(borrower_name + "'s cash account")

            #Inputs amount we want to post
            WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, 'manual-amount'))).send_keys(str(round(row['Test if Run Amount'] * -1,2)))

            comment_box.send_keys('Test if run ' + str(row['Old Cashfac IDs']))

            #Clicks save button
            browser.find_element_by_xpath('//*[@id="new_create_authorization_request"]/div[7]/input').click()

         else:
             pass
    
    text.insert(tk.END, "Step 7 complete. " + str(i) + " transactions should've posted. Test if run transactions have been posted if necessary. \nIt took " + str(datetime.datetime.now() - start_time) + " to run through the 8 steps. \n\n")
    text.insert(tk.END, "The consolis process was not completed for the following cashfacs...\n" + str(incomplete_consolis))
#------------------ Create buttons we can click to run the above functions ----------------------#

consolidations_sheet = tk.Button(master, text = 'Step 1: Build \n Consolidations Sheet', command = runConsolidationsSheetMacro, width = 20, padx = 2, pady = 3, font = ("Helvetica", 9, "bold")).grid(row = 4, column = 0, pady = 5, padx = 15)
export_check = tk.Button(master, text = 'Step 2: Check Exports -\n Paused Page', command = run_exports_check, width = 20, padx = 2, pady = 3, font = ("Helvetica", 9, "bold")).grid(row = 4, column = 1, pady = 5, padx = 15)
additional_funds = tk.Button(master, text = 'Step 3: Grab Additional \n Funds Required', command = grabAdditionalFundsRequired, width = 20, padx = 2, pady = 3, font = ("Helvetica", 9, "bold")).grid(row = 4, column = 2, pady = 5, padx = 15)
backend_update = tk.Button(master, text = 'Step 4: Update Backend \n Loan Amount', command = runBackendUpdateMacro, width = 20, padx = 2, pady = 3, font = ("Helvetica", 9, "bold")).grid(row = 5, column = 0, pady = 5, padx = 15)
new_loan_settling = tk.Button(master, text = 'Step 5: Post Settling \n Previous Loans', command = post_new_cashfacs, width = 20, padx = 2, pady = 3, font = ("Helvetica", 9, "bold")).grid(row = 5, column = 1, pady = 5, padx = 15)
old_loan_settling = tk.Button(master, text = 'Step 6: Post Settling \n With New Loan', command = post_old_cashfacs, width = 20, padx = 2, pady = 3, font = ("Helvetica", 9, "bold")).grid(row = 5, column = 2, pady = 5, padx = 15)
post_test_if_run = tk.Button(master, text = 'Step 7: Post "test if \n run" Transactions', command = post_test_if_run, width = 20, padx = 2, pady = 3, font = ("Helvetica", 9, "bold")).grid(row = 6, column = 0, pady = 5, padx = 15)
settle_in_backend = tk.Button(master, text = 'Step 8: Settle loan in \n Backend', command = settleInBackend, width = 20, padx = 2, pady = 3, font = ("Helvetica", 9, "bold")).grid(row = 6, column = 1, pady = 5, padx = 15)


master.mainloop()
