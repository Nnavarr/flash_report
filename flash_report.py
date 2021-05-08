import pandas as pd
import numpy as np
import datetime
import os
import win32com.client
import xlwings as xw
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from time import sleep
from textwrap import dedent
import glob
import re
import sys
from getpass import getuser
import downloader

"""
Author: Noe Navarro
Date: 11/16/2020
Objective:
    Process flash report related files for ease of aggregation
Update Log
----------
Version 0.1.0: Inception of life | 11/16/2020 | NN
Version 0.1.1: Created .net query bot which pulls Fridays BE/BP | 11/17/2020 | NN
Version 0.1.2: Created ability to modify webpages for PDF print | 12/2/2020 | NN
Version 0.1.3: Created ability to autogenerate emails to relevant parties | 12/3/2020 | NN
"""

# u-move pzt email processing
def pzt_email(user_email, desktop_path):

    """
    Author: Noe Navarro
    Date: 11/16/2020
    return1: email_attch_path; filepath for pzt u-move spreadsheet
    processes pzt email and saves where necessary
    """
    # establish outlook parameters
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    inbox = outlook.Folders[f'{user_email}'].Folders['Flash Report'].Items
    save_path = rf'{desktop_path}'
    email_attch_path = ''
    today = datetime.date.today().strftime(format='%Y-%m-%d')

    # check for pzt as sender (sent Fridays)
    for email in inbox:
        try:
            if (email.SenderEmailAddress == '<email>') & (email.SentOn.strftime(format='%Y-%m-%d') == today):

                # download attachment to desktop
                attachment = email.Attachments.item(1)
                attachment.SaveAsFile(save_path + '\\' + 'pzt_umv_temp.csv')

                # check for friday date
                wb = xw.Book(rf'{save_path}\pzt_umv_temp.csv')
                sheet1 = wb.sheets[0]
                create_date = sheet1.range('O2').value

                assert create_date.strftime('%Y-%m-%d') == datetime.date.today().strftime('%Y-%m-%d'),'The Excel file creation date does not match the Friday date'
                #if the friday days don't match, throw an error

                # continue with aggregation if date checks out
                week = int(sheet1.range('M2').value)
                if (datetime.date.today().month >= 4) & (datetime.date.today().month <= 12):
                    fiscal_year = datetime.date.today().year + 1
                else:
                    fiscal_year = datetime.date.today().year

                # name and save files
                josh_loc = rf'\\adfs01.uhi.amerco\departments\mia\group\MIA\Flash Report\FY{fiscal_year}\Other\FLASHUmvWk{week}.csv'
                wb.save(josh_loc)
                email_attch_path = josh_loc

                archive_loc = rf'\\adfs01.uhi.amerco\departments\mia\group\MIA\Flash Report\FY{fiscal_year}\Details\FLASH - FY{fiscal_year} - Wk{week}.csv'
                wb.save(archive_loc)
                wb.close()
            else:
                pass
        except:
            print(f'There was an issue with the {pzt_email.__name__} function.')

    return email_attch_path

# Auction wholesale PDF extraction
def brandons_auction_pdf(user_email):

    """
    Author: Noe Navarro
    Date: 12/2/2020
    return1: <rename to output>; filepath for auction wholesale group performance
        to be processed later within the PDF attachment emails
    """

    # calculate fiscal year
    if (datetime.date.today().month >= 4) & (datetime.date.today().month <= 12):
        fiscal_year = datetime.date.today().year + 1
    else:
        fiscal_year = datetime.date.today().year

    # establish outlook parameters
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    inbox = outlook.Folders[f'{user_email}'].Folders['Flash Report'].Items
    save_path = rf'\\adfs01.uhi.amerco\departments\mia\group\MIA\Flash Report\FY{fiscal_year}\Other'
    pdf_path = ''
    today = datetime.date.today()
    wed_date = today + datetime.timedelta(-2)

    for email in inbox:
        try:
            if (str(email.Sender) == 'Brandon Crim') & (email.SentOn.strftime(format='%Y-%m-%d') == wed_date.strftime(format='%Y-%m-%d')):
                # download pdf attachment
                attachment = email.Attachments.item(1)
                attachment.SaveAsFile(save_path + '\\' + str(attachment))
                pdf_path = save_path + '\\' + str(attachment)
            else:
                pass
        except:
            print(f'There was an issue with the {brandons_email.__name__} function.')
        # the function can return the pdf_path in order to satisfy the pdf filename

# moving help
def moving_help(user_email):

    """
    Author: Noe Navarro
    Date: 11/17/2020
    param1: NA
    return: current and last year moving help numbers. We only use last year for
        Josh's email.
    """
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    inbox = outlook.Folders[f'{user_email}'].Folders['Flash Report'].Items
    today = datetime.date.today()
    wed_date = today + datetime.timedelta(-2)

    # containers
    cy_arr, ly_arr = 0, 0

    # extract moving help numbers
    for email in inbox:
        if (str(email.Sender) == '<user name>') & (email.SentOn.strftime(format='%Y-%m-%d') == wed_date.strftime(format='%Y-%m-%d')):

            # split string
            str_list = email.body.split()

            for i in range(len(str_list)):
                if str_list[i] == 'CY':
                    cy_arr = int(str_list[i + 2])
                elif str_list[i] == 'LY':
                    ly_arr = int(str_list[i + 2])
                else:
                    pass
    return cy_arr

# BE/BP number
class MIQ_bot(object):

    def __init__(self):

        # pass MIQ query path
        url = '<url>'
        self.driver = webdriver.Chrome(r'<filepath>')

        self.driver.get(url)
        sleep(1)

        # login
        self.driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_tb_ssn"]').send_keys('1217543')
        self.driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_tb_pin"]').send_keys(os.getenv('pw'))
        self.driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_bt_login"]').click()
        sleep(2)

    def be_bp(self):

        """
        **Important**
        Apply report parameters. These WILL NOT SAVE even if a view is made
        within the MIQ portal. All relevant parameters will be passed here
        FYI: The dropdown menus can interfere with another's selection if the
        drop down overlaps. In order to avoid this issue, the bot will click inside
        search box up-top to close all active drop down menus.
        """
        # MCO selection (All)
        self.driver.find_element_by_xpath('//*[@id="reportinput"]/div/fieldset[1]/div/div[7]/label/div').click()
        sleep(1)
        self.driver.find_element_by_xpath('/html/body/div[7]/div/ul/li[1]/a/span').click()
        self.driver.find_element_by_xpath('//*[@id="Text"]').click()

        # Model Selection
        self.driver.find_element_by_xpath('//*[@id="reportinput"]/div/fieldset[2]/div/div[4]/label/div').click()
        sleep(1)
        self.driver.find_element_by_xpath('/html/body/div[12]/div/ul/li[1]/a').click()
        sleep(4)

        # Date Range (first click away from it to allow Selenium to click)
        self.driver.find_element_by_xpath('//*[@id="Text"]').click() # reset of windows so we can traverse through the site
        sleep(1)
        self.driver.find_element_by_xpath('//*[@id="reportinput"]/div/fieldset[3]/div/div[1]/label/div').click()
        self.driver.find_element_by_xpath('//*[@id="reportinput"]/div/fieldset[3]/div/div[1]/label/div/ul/li[3]').click()
        self.driver.find_element_by_xpath('//*[@id="Text"]').click() # reset windows
        sleep(4)

        # Prior Year
        self.driver.find_element_by_xpath('//*[@id="reportinput"]/div/fieldset[3]/div/div[9]/label/div').click()
        sleep(1)
        self.driver.find_element_by_xpath('//*[@id="reportinput"]/div/fieldset[3]/div/div[9]/label/div/ul/li[3]').click()
        sleep(4)
        self.driver.find_element_by_xpath('//*[@id="Text"]').click()  # reset windows
        sleep(1)

        # click "Query"
        self.driver.find_element_by_xpath('//*[@id="submitInput"]').click()
        sleep(15)

        # extract "IT + OW Gross" $
        number = self.driver.find_element_by_xpath('//*[@id="dtable0"]/tbody/tr/td[8]').text
        number = int(number.replace(',', ''))

        # close webbrowser
        self.driver.close()
        return number

# Flash PDF creation
class Flash_bot(object):

    """
    Extracts the preliminary flash html table to be pasted within the
        '.net Web Query' sheet in the FLash Report Excel aggregation.
        This is dont to check for any variance in compilation via the 'Variance Check' sheet.
    """

    def __init__(self):

        # flash URL
        url = '<url>'
        self.driver = webdriver.Chrome(r'C:\bin\chromedriver.exe')

        self.driver.get(url)
        sleep(1)

        # login
        self.driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_tb_ssn"]').send_keys('1217543')
        self.driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_tb_pin"]').send_keys(os.getenv('pw'))
        self.driver.find_element_by_xpath('//*[@id="ContentPlaceHolder1_bt_login"]').click()
        sleep(2)

        # Javascript injection
        # remove header
        js1 = "var aa=document.getElementsByClassName('fixed')[0];aa.remove()"
        js2 = "var user=document.getElementsByClassName('user')[0]; user.remove()"

        # remove flash report hyperlinks
        js3 = "var hyperLinks=document.getElementsByClassName('subnav')[0]; hyperLinks.remove()"

        # remove green box (top)
        js4 = "var greenBox = document.getElementsByClassName('action_form')[0]; greenBox.remove()"

        # remove green box (bottom)
        js5 = "var greenBoxBottom = document.getElementById('ctl00_cphMainBody_divChart'); var gBoxTwo = greenBoxBottom.getElementsByClassName('action_form')[0].remove();"

        # js6 = "var reportGraphSpace = document.getElementsByClassName('data_table')[0]; var childBreak = reportGraphSpace.getElementsByClassName('print_break'); childBreak.innerHTML = '<p>&nbsp</p><p>&nbsp</p';"

        # remove footer
        js6 = "var footer = document.getElementsByClassName('footer')[0].remove();"

        # remove header
        js_flash_banner = "var header = document.getElementsByClassName('description')[0].remove()"

        # Injection lists
        self.js_rev = [js1, js2, js3, js4, js5, js6]
        self.js_others = [js1, js2, js3, js4, js6, js_flash_banner]

    def prelim_flash(self):

        # click 'go' to generate flash table
        self.driver.find_element_by_xpath('//*[@id="ctl00_cphMainBody_btnSubmit"]').click()
        sleep(2)

        # extract HTML table of Flash Totals
        table = self.driver.find_element_by_xpath('//*[@id="aspnetForm"]/div[3]/div/div[1]/div/div[3]/table').get_attribute('outerHTML')
        df = pd.read_html(table, header=0, index_col=0)
        self.driver.close()

        return df

    def rev_key_exp(self):

        """
        Cycle through final Flash webpages and format HTML accordingly for pdf print
        """
        sleep(2)

        # click 'go' to generate flash table
        self.driver.find_element_by_xpath('//*[@id="ctl00_cphMainBody_btnSubmit"]').click()
        sleep(2)

        # loop and execute JS
        for js in self.js_rev:
            self.driver.execute_script(js)

    def truck_performance(self):
        """
        Navigate to the truck_performance page and remove HTML elements
        """
        truck_url = '<url>'
        self.driver.get(truck_url)
        sleep(2)

        # click 'go' to generate flash table
        self.driver.find_element_by_xpath('//*[@id="ctl00_cphMainBody_btnSubmit"]').click()
        sleep(2)

        # javascript injection
        for js in self.js_others:
            self.driver.execute_script(js)

    def trailer_performance(self):
        """
        Navigate to the truck_performance page and remove HTML elements
        """
        truck_url = '<url>'
        self.driver.get(truck_url)
        sleep(2)

        # click 'go' to generate flash table
        self.driver.find_element_by_xpath('//*[@id="ctl00_cphMainBody_btnSubmit"]').click()
        sleep(2)

        # javascript injection
        for js in self.js_others:
            self.driver.execute_script(js)

    def SRI_performance(self):
        """
        Navigate to the truck_performance page and remove HTML elements
        """
        truck_url = '<url>'
        self.driver.get(truck_url)
        sleep(2)

        # click 'go' to generate flash table
        self.driver.find_element_by_xpath('//*[@id="ctl00_cphMainBody_btnSubmit"]').click()
        sleep(2)

        # javascript injection
        for js in self.js_others:
            self.driver.execute_script(js)

def josh_email(movinghelp, bebp, file_loc):

    """
    Author: Noe Navarro
    Date: 11/17/2020
    param1: NA
    return: Forwards an email to designated recepients with the attached
    flash report spreadsheet
    Email parameters
    -----------------
    Email attachments:
        josh_loc csv | complete
    Body:
        BE/BP from MIQ_bot | complete NN
        Moving Help: David Lopresti | complete NN
    """
    outlook = win32com.client.Dispatch('Outlook.Application')
    email = outlook.CreateItem(0)
    email.To = '<email>'
    # email.cc = ''
    email.Subject = 'Flash Report'
    email.HtmlBody = dedent("""
        Good morning, </p>
        BE/BP: {} <br>
        Moving Help: {} </p>
        Regards, <br>
        Financial Analysis
        """).format(bebp, movinghelp)

    email.Attachments.Add(Source=file_loc)
    email.Send()

def download_rpt():

    # import flashrpt
    flashrpt_downloader.main()
    user = getuser()
    flashrpt_path = fr'C:\Users\{user}\Desktop\Flashrpt.csv'
    wb = xw.Book(flashrpt_path)
    data = wb.sheets[0].range('A2:E97').value

    # archive
    # calcualte fiscal year
    if (datetime.date.today().month >= 4) & (datetime.date.today().month <= 12):
        fiscal_year = datetime.date.today().year + 1
    else:
        fiscal_year = datetime.date.today().year

    # flash week
    week_list = []
    root_dir = rf'\\adfs01.uhi.amerco\departments\mia\group\MIA\Flash Report\FY{fiscal_year}'
    for filename in glob.iglob(str(root_dir) + '\*.xlsm'):
        z = re.findall('[aA-zZ]{2}[0-9]{1,2}', filename)[3]
        wk_num = int(''.join(re.findall('[0-9]', z)))
        week_list.append(wk_num)

    try:
        max_wk = max(week_list)
    except:
        max_wk = 1

    wb.save(root_dir + '\\Details\\' + f'Flashrpt - FY{fiscal_year} - Wk{max_wk}.csv')
    wb.close()

    return data

"""
Compile Excel spreadsheet
"""
def excel_flash_comp():

    # extract fiscal year
    if (datetime.date.today().month >= 4) & (datetime.date.today().month <= 12):
        fiscal_year = datetime.date.today().year + 1
    else:
        fiscal_year = datetime.date.today().year

    # extract flash week
    week_list = []
    root_dir = rf'\\adfs01.uhi.amerco\departments\mia\group\MIA\Flash Report\FY{fiscal_year}'
    for filename in glob.iglob(str(root_dir) + '\*.xlsm'):
        z = re.findall('[aA-zZ]{2}[0-9]{1,2}', filename)[3]
        wk_num = int(''.join(re.findall('[0-9]', z)))
        week_list.append(wk_num)

    try:
        max_wk = max(week_list)
    except:
        max_wk = 1

    # calcualte appropriate week for the new filepath
    new_wk = max_wk + 1
    new_path = rf'\\adfs01.uhi.amerco\departments\mia\group\MIA\Flash Report\FY{fiscal_year}\Flash Report - FY{fiscal_year} - Wk{new_wk}.xlsm'

    # Open max workbook and update each sheet
    wb = xw.Book(rf'\\adfs01.uhi.amerco\departments\mia\group\MIA\Flash Report\FY{fiscal_year}\Flash Report - FY{fiscal_year} - Wk{max_wk}.xlsm')

    # revenue and key expenses
    rev_key_exp = wb.sheets['Revenue and Key Expense']
    rev_key_exp.range('BC2').value = fiscal_year
    rev_key_exp.range('BC3').value = new_wk

    # data entry
    data_entry = wb.sheets['Data Entry']
    data_entry.range('C4').value = new_wk
    data_entry.range('F4:H4').api.ClearContents()
    data_entry.range('B8:M8').api.ClearContents()
    data_entry.range('B12:F107').api.ClearContents()

    # update values
    data_entry.range('F4').value = num2 / 1000
    data_entry.range('G4').value = num1 / 1000

    # import flash umove data
    umv_wb = xw.Book(filepath)
    umv_sheet = umv_wb.sheets[0]
    data_entry.range('B8:M8').value = umv_sheet.range('A2:L2').value
    umv_wb.close()

    # download and import flashrpt
    rpt_data = download_rpt()
    data_entry.range('B12:F107').value = rpt_data

    # update flash numbers macro
    wb.macro('UpdateFlashData')() #update flash data macro
    wb.save(new_path)

"""
Email Compilation: Completion
"""
def flash_live_email():

    """
    Author: Noe Navarro
    Date: 12/2/2020
    param1: NA
    return: Forwards email to appropriate parties, lets them know the flash is live.
    """

    outlook = win32com.client.Dispatch('Outlook.Application')
    email = outlook.CreateItem(0)
    email.To = '<email>'
    email.cc = '<email>'

    # calculate fiscal year
    if (datetime.date.today().month >= 4) & (datetime.date.today().month <= 12):
        fiscal_year = datetime.date.today().year + 1
    else:
        fiscal_year = datetime.date.today().year

    # flash week
    week_list = []
    root_dir = rf'\\adfs01.uhi.amerco\departments\mia\group\MIA\Flash Report\FY{fiscal_year}'
    for filename in glob.iglob(str(root_dir) + '\*.xlsm'):
        z = re.findall('[aA-zZ]{2}[0-9]{1,2}', filename)[3]
        wk_num = int(''.join(re.findall('[0-9]', z)))
        week_list.append(wk_num)

    try:
        max_wk = max(week_list)
    except:
        max_wk = 1

    email.Subject = f'Flash Report Week {max_wk}'

    email.HtmlBody = dedent("""
        Good afternoon, </p>
        Flash week {} is now active. Click the link to access the report on company.net:  <a href='<email>'>Flash Report</a> </p>
        Regards, <br>
        Financial Analysis
    """).format(max_wk)

    email.Send()

def flash_pdf_email():

    """
    Author: Noe Navarro
    Date: 12/2/2020
    param1: NA
    return: Forwards email with relevant attachments to be forwarded to appropriate parties
    ** Important **
    With regards to the auction pdf sent by Brandon Crim on Wednesdays, we are able to extract
    the max value file modify date because it should be saved within the filepath before running
    the flash aggregation.  As of 12/3/2020, there is no catch to ensure the most recent week's file is
    pulled in. Be sure to run the pdf aggregation after the file has been saved in the desired location.
    """

    outlook = win32com.client.Dispatch('Outlook.Application')
    email = outlook.CreateItem(0)
    email.To = '<email>'
    email.cc = '<email>'

    # calculate fiscal year
    if (datetime.date.today().month >= 4) & (datetime.date.today().month <= 12):
        fiscal_year = datetime.date.today().year + 1
    else:
        fiscal_year = datetime.date.today().year

    # flash week
    week_list = []
    root_dir = rf'\\adfs01.uhi.amerco\departments\mia\group\MIA\Flash Report\FY{fiscal_year}'
    for filename in glob.iglob(str(root_dir) + '\*.xlsm'):
        z = re.findall('[aA-zZ]{2}[0-9]{1,2}', filename)[3]
        wk_num = int(''.join(re.findall('[0-9]', z)))
        week_list.append(wk_num)

    try:
        max_wk = max(week_list)
    except:
        max_wk = 1

    # calculate week start and finish;
    today = datetime.date.today()
    week_start = today + datetime.timedelta(-9)
    week_end = today + datetime.timedelta(-3)

    email.Subject = f'Flash Report Week {max_wk} PDF'

    email.HtmlBody = dedent("""
        Alyssa, </p>
        See attached for this week's Flash files (Wk{}, {} -{}) </p>
        Regards, <br>
        Financial Analysis
    """).format(max_wk, week_start.strftime(format='%m/%d'), week_end.strftime(format='%m/%d/%Y'))

    rev_key_file = rf'\\adfs01.uhi.amerco\departments\mia\group\MIA\Flash Report\FY{fiscal_year}\PDF\Revenue & Key Expenses - FY{fiscal_year} - Wk{max_wk}.pdf'
    truck_file = rf'\\adfs01.uhi.amerco\departments\mia\group\MIA\Flash Report\FY{fiscal_year}\PDF\Truck Performance - FY{fiscal_year} - Wk{max_wk}.pdf'
    trailer_file = rf'\\adfs01.uhi.amerco\departments\mia\group\MIA\Flash Report\FY{fiscal_year}\PDF\Trailer Performance - FY{fiscal_year} - Wk{max_wk}.pdf'
    sri_file = rf'\\adfs01.uhi.amerco\departments\mia\group\MIA\Flash Report\FY{fiscal_year}\PDF\SRI Performance - FY{fiscal_year} - Wk{max_wk}.pdf'

    # llist containers
    file_list, auction_list, modify_time = [], [], []

    # extract most recent acution file
    for root, dirs, files in os.walk(rf'\\adfs01.uhi.amerco\departments\mia\group\MIA\Flash Report\FY{fiscal_year}\Other'):
        file_list.append(files)

    # if pdf, extract append to a new list
    for file in file_list[0]:
        if '.pdf' in file:
            auction_list.append(file)

            # extract file modify time
            modify_time.append(datetime.datetime.fromtimestamp(os.path.getctime(rf'\\adfs01.uhi.amerco\departments\mia\group\MIA\Flash Report\FY{fiscal_year}\Other' +'\\'+ file)))
        else:
            pass

    # create dictionary & extract filename using value
    auct_dict = dict(zip(auction_list, modify_time))
    auction_filename = ''
    for key, val in auct_dict.items():
        if val == max(auct_dict.values()):
            auction_filename = key
    auction_filepath = rf'\\adfs01.uhi.amerco\departments\mia\group\MIA\Flash Report\FY{fiscal_year}\Other\{auction_filename}'

    # loop through all relevant filepaths and append to email
    pdf_attachments = [rev_key_file, truck_file, trailer_file, sri_file, auction_filepath]
    for pdf in pdf_attachments:
        email.Attachments.Add(Source=pdf)
    email.Send()

def flash_uhi_email(desktop_path):
    """
    Author: Noe Navarro
    Date: 12/3/2020
    param1: NA
    return: Forwards the flash UHI email to relevant parties
    """

    outlook = win32com.client.Dispatch('Outlook.Application')
    email = outlook.CreateItem(0)
    email.To = '<email>'
    email.cc = '<email>'

    # calcualte fiscal year
    if (datetime.date.today().month >= 4) & (datetime.date.today().month <= 12):
        fiscal_year = datetime.date.today().year + 1
    else:
        fiscal_year = datetime.date.today().year

    # flash week
    week_list = []
    root_dir = rf'\\adfs01.uhi.amerco\departments\mia\group\MIA\Flash Report\{fiscal_year}'
    for filename in glob.iglob(str(root_dir) + '\*.xlsm'):
        z = re.findall('[aA-zZ]{2}[0-9]{1,2}', filename)[3]
        wk_num = int(''.join(re.findall('[0-9]', z)))
        week_list.append(wk_num)
    try:
        max_wk = max(week_list)
    except:
        max_wk = 1

    # calculate week start and finish;
    today = datetime.date.today()
    week_start = today + datetime.timedelta(-9)
    week_end = today + datetime.timedelta(-3)

    # attach flash rpt and uhi
    email.Subject = 'Flash Report'
    email.HtmlBody = dedent("""
        David, </p>
        See attached for this week's Flash files (Wk{}, {} -{}) </p>
        Regards, <br>
        Financial Analysis
    """).format(max_wk, week_start.strftime(format='%m/%d'), week_end.strftime(format='%m/%d/%Y'))

    # attach flah rpt and uhi
    rpt_path = rf'{desktop_path}\Flashrpt.csv'
    uhi_path = rf'{desktop_path}\Flashuhi.csv'

    file_list = [rpt_path, uhi_path]
    for file in file_list:
        email.Attachments.Add(Source=file)

    email.Send()

"""
Flash aggregation function
"""
def main():

    """
    IMPORTANT:
        Update email variable to the preparer's email.
        Update fiscal_year and max_wk variables if needed for transition to new fiscal year
    """
    email = '<email>''

    # data compilation
    filepath = pzt_email(user_email=email, desktop_path=r'C:\Users\1217543\Desktop') # step1: process pzt friday umove email
    num1 = moving_help(user_email=email) * 1000 # step2: moving help
    num2 = MIQ_bot().be_bp()
    josh_email(movinghelp=num1, bebp=num2, file_loc=filepath)

    # extracts the auction pdf sent from Brandon Crim and saves to archive
    brandons_auction_pdf(user_email=email)

    # compile Excel spreadsheet; it will also download the rpt/uhi files to desktop & archive
    excel_flash_comp()

    """
    Prelim Flash email
    -------------------
    Once the preliminary flash data is uploaded and correct, we can send the prelim flash email.
    Be sure to give an hour window for people to respond. No response is a good response.
    Activate new week
    -----------------
    With the hour gone by, we can activate the preliminary view to be published. Once that is complete
    we can download the PDF files and aggrregate the final emails.
    """

    # compile Flash PDFs
    Flash_bot().rev_key_exp()
    # html for space: <p>&nbsp</p><p>&nbsp</p>
    Flash_bot().truck_performance()
    Flash_bot().trailer_performance()
    Flash_bot().SRI_performance()

    """
    Important: Only run these after the PDF files have been saved.
    """
    flash_live_email()
    flash_pdf_email()
    flash_uhi_email(desktop_path = r'C:\Users\1309919\Desktop')
