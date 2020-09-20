#!/usr/bin/env python
# coding: utf-8

from selenium import webdriver
import pandas as pd
import xlwings as xw
import time
from selenium.webdriver.support.select import Select
from selenium.common.exceptions import TimeoutException,NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options


def offer_form_automate():
    # In[2]:


    # Paths & Urls Variable
    HR_LINKS_PORTAL_URL = r"http://grdhrdbwhq00.northamerica.cerner.net/ReportServer?/HR+Links+Portal/HR_Links_Portal&rs:Command=Render&rc:Toolbar=False"
    CHROME_DRIVER_PATH = "C:/Users/MS064294/Downloads/chromedriver_win32/chromedriver"
    #CHROME_DRIVER_PATH = "C:/Program Files (x86)/chromedriver_win32/chromedriver"
    FILE_PATH = r'C:/Users/MS064294/Desktop/OfferFormAutomation.xlsm'
    ARROW_DOWN = u'\ue015'

    # Visible Text to Value under Select Tag
    Hiring_Type = {"Hire":"HIR" , "Rehire":"REI" , "Transfer":"TRA"}
    Shift_Type = {"Normal":"NORM" , "AMS Shift":"AMS" , "AUS Shift":"AUS"}
    Virtual_Position = {"No":"N" , "Yes":"Y"}
    CPP ={"None":"1" , "TBL": "2" , "TCA": "3"}
    Pay_Group = {"Hourly" :"Overtime Eligible" , "Salary": "Salary", "Salary Overtime Eligible": "Salary Overtime Eligible"}
    Advance = {"1":"None" , "2": "Quick Start", "3": "Recoverable Only","4" : "Guaranteed","5":"Quick Start and Fully Recoverable (Combo)"}

    wb = xw.Book(FILE_PATH)
    sht = wb.sheets['Sheet1']
    print(sht.range('C2').value)

    # In[4]:

    options = Options()
    options.add_argument('--start-maximized')
    browser = webdriver.Chrome(CHROME_DRIVER_PATH, options=options)
        
    browser.implicitly_wait(30) #wait 30 seconds when doing a find_element before carrying on

    browser.get(HR_LINKS_PORTAL_URL)
        
    try:
        window_before = browser.window_handles[0]
        print(browser.title)
            
        kenexa_link = browser.find_element_by_link_text("Kenexa")
        kenexa_link.click()
            
            
        #get the window handle after a new window has opened
        window_after = browser.window_handles[1]
            
        browser.switch_to.window(window_after)
        browser.implicitly_wait(30)
        print(browser.title)
        time.sleep(5)
        browser.find_element_by_id('searchlink').click()
        time.sleep(5)
        #browser.find_element_by_css_selector('a span.reqsimg').click()
        #time.sleep(2)
        browser.execute_script('''
            document.querySelector('a span.reqsimg').click();
        ''')
        time.sleep(3)
        
        '''
        browser.find_element_by_xpath('//*[@id="menuLink"]/span').click() #Clicking Menu Button
        time.sleep(3)
        browser.find_element_by_xpath('//*[@id="pm_1"]').click() #Clicking Candidates Dropdown 
        time.sleep(3)
        browser.find_element_by_xpath('//*[@id="cm_1_0"]').click() #Clicking My Candidates Dropdown
        time.sleep(3)
        browser.find_element_by_xpath('//*[@id="scm_1_0_0"]').click() #Clicking All Candidates option
        time.sleep(3)
        '''
        
        data = pd.read_excel(FILE_PATH)
        data = data.dropna(how='all')
        print(data)
        for row_index, row in data.iterrows():
            if row['Req ID'] and isinstance(row['Req ID'], str) and len(str(row['Req ID']).strip()) > 0 and str(row['FORM-STATUS']).lower().strip() != 'completed':            
                status_row = 'C' + str(row_index+2)
                print(row['Req ID'])
                print(type(row['Req ID']))
                sht.range(status_row).value = 'IN PROGRESS'
                wb.save()
            
                try:
                    candidate_name = row['Candidate Names']
                    req_id = row['Req ID']
                    time.sleep(2)
                    
                    search_field = browser.find_element_by_xpath('//*[@id="quicksearch"]')  #Quick Search text Field
                    search_field.clear()
                    #search_field.send_keys(candidate_name) # Sending Candidate name for Quick Search
                    search_field.send_keys(req_id) # Send REQ-ID to Quick-search field
                    time.sleep(2)
                    #browser.find_element_by_xpath('//*[@id="pageContents"]/form/div[1]/div/div[2]/div/span').click() #clicking Search Icon
                    browser.execute_script('''
                        document.querySelector('div.globalSearchInput span.searchimg').click();
                    ''') # click on search button, it did not work NON-JS way, so did it by JS
                    time.sleep(2)
                    req_link_found = browser.find_element_by_link_text(req_id)
                    time.sleep(2)
                    
                    browser.find_element_by_css_selector('a.candcnt_total').click() # Click on Total Candidates
                    time.sleep(3)
                    
                    #
                    #
                    #browser.find_element_by_css_selector('a.trlink.candname') # Wait for table to load candidates
                    #search_cand_input = browser.find_element_by_id('quicksearchcand') # Filter box on left panel
                    #search_cand_input.send_keys(candidate_name) # Send Candidate Name to the filter box
                    #browser.execute_script('''
                    #    document.querySelector('div.facetblock span.searchimg').click();
                    #''') # Click on search icon next to filter box
                    #time.sleep(3)
                    #browser.find_element_by_link_text(candidate_name)
                    #
                    #'''
                    
                
                    time.sleep(3)
                
                    try:
                        hr_status_dropdown = browser.find_element_by_xpath("//span[@class='ng-binding'][contains(text(),'HR Status')]")
                        hr_status_dropdown.click()
                    except NoSuchElementException as nosee:
                        filters_button = browser.find_element_by_css_selector("input[ng-click='toggleFilterSideBar()']")
                        filters_button.click()
                        time.sleep(1)
                        hr_status_dropdown = browser.find_element_by_xpath("//span[@class='ng-binding'][contains(text(),'HR Status')]")
                        hr_status_dropdown.click()  
                
                    time.sleep(6)
                    browser.implicitly_wait(10)
                
                    try:
                        assign_cand_input = browser.find_element_by_xpath("//input[@id='hrstatus-870557']") # Assign Candidate to Hiring Req           
                    except NoSuchElementException:
                        try:
                            assign_cand_input = browser.find_element_by_xpath("//input[@id='hrstatus-862943']") # Interview
                        except NoSuchElementException:
                            assign_cand_input = browser.find_element_by_xpath("//input[@id='hrstatus-13704']") # Prepare Offer (in case my script has changed the status previously , but not submitted the form)
                            
                    time.sleep(2)        
                            
                    browser.implicitly_wait(30)
                    
                    browser.execute_script('''
                        var assigncandinput = arguments[0];
                        assigncandinput.click();
                    ''', assign_cand_input)
                
                    time.sleep(10)
                
                    #cand_name_link = browser.find_element_by_css_selector('a.candname')
                    #cand_name_link.click()
                
                    #time.sleep(3)
                
                    #browser.switch_to.window(browser.window_handles[2])
                    
                    """
                    prepare_offer_flag = browser.execute_script('''
                        var cells = document.querySelectorAll('.ui-grid-cell-contents');
                        var a_hr_status;
                        for(var i = 0; i < cells.length; i++){
                            if(cells[i].innerText.includes('HR Status') && cells[i].innerHTML.includes('Update HR Status for')){
                                //var title="Update HR Status for Test, Vaishali  - 56901BR";
                                //var title = 'Update HR Status for ' + arguments[0] + '  - ' + arguments[1];
                                //var title = 'Update HR Status for';
                                var update_hr_status = 'a[title^="Update HR Status for"]';

                                a_hr_status = cells[i].querySelector(update_hr_status);
                                
                                if(!(a_hr_status.innerText.includes('Prepare Offer'))){
                                    a_hr_status.click();
                                    i = cells.length;
                                }
                            }
                        }
                        return a_hr_status;

                    ''', candidate_name, req_id)
                    """
                    
                    prepare_offer_flag = browser.execute_script('''
                        var a_hr_status = document.querySelector('.ui-grid-cell-contents a[title^="Update HR Status for"]');
                        if(!(a_hr_status.innerText == 'Prepare Offer')){
                            a_hr_status.click();
                        }
                        return a_hr_status;
                    ''')
                    
                    #print(prepare_offer_flag.get_attribute('innerText'))
                    print('prepare_offer_flag')
                    print(prepare_offer_flag)
                    prepare_offer_status_found = False
                    
                    
                    #if prepare_offer_flag and prepare_offer_flag.get_attribute('innerText') != 'Prepare Offer':  
                    #    #browser.find_element_by_css_selector('div.hraction')
                    #    mySelectStatus = Select(browser.find_element_by_id("hrstatusoptions"))
                    #    status_list = browser.find_element_by_xpath('//*[@id="hrstatusoptions-menu"]')
                    #    
                    #    #print('mySelectStatus')
                    #    #print(mySelectStatus)
                    #    #print('status_list')
                    #    #print(status_list)
                    #    
                    #    for o in mySelectStatus.options:
                    #        print(o.get_attribute('value'))
                    #        print(o.get_attribute('innerText'))
                    #        if o.get_attribute('value')=="13704" and o.get_attribute('innerText').strip().lower()=="Prepare Offer".lower(): # '13704' is the code for PREPARE-OFFER
                    #            status_list.send_keys(Keys.ENTER)
                    #            prepare_offer_status_found = True
                    #            prepare_offer_flag = o
                    #            time.sleep(1)
                    #            break
                    #        else:
                    #            #status_list.send_keys(Keys.UP)
                    #            status_list.send_keys(ARROW_DOWN)
                    #            time.sleep(1)
                    #    
                    #    # above logic fails sometimes to change the status to pre-offer
                    #    # final trial to be done with new logic : Select this option : <option value="-1">Advanced Options</option>
                    #    # it will open new window : where this select box is there :
                    #    
                    #    #<select name="hrstatus" class="hrstatus hasModel ng-pristine ng-untouched ng-valid" id="hrstatus" ng-model="selectedstatus" select-menu="" style="width: 30%; display: none;" ng-class="{'ts-error':!selectedstatus &amp;&amp; formSubmit &amp;&amp; !undomode}" aria-labelledby="hrstatusScreenReaderTxt" tabindex="0" aria-invalid="false">
                    #    #<option value="" aria-label="None"></option>
                    #    #<option value="13704" folderresumeid="34035" popupformrequired="false">Prepare Offer</option>
                    #    #<option value="321733" folderresumeid="345" popupformrequired="false">Reject-Event Manager</option>
                    #    #<option value="321737" folderresumeid="340535" popupformrequired="false">Candidate Withdraw-Event Manager</option>
                    #    #</select>
                    #    
                    #    # select the right option here n submit n come back to old-window
                    #    # then click forms icon n proceed
                    #    
                    #    """
                    #    browser.execute_script('''
                    #        statusList = arguments[0];
                    #        statusList.click();                            
                    #    ''', status_list)
                    #    
                    #    time.sleep(2)
                    #    
                    #    browser.execute_script('''
                    #        statusesListElems = document.querySelectorAll('#hrstatusoptions-menu li.ui-menu-item');
                    #            for(var k = 0; k < statusesListElems.length; k++){
                    #                if(statusesListElems[k].innerText.includes('Prepare Offer')){
                    #                    statusesListElems[k].click();
                    #                }
                    #            }
                    #    ''')
                    #    """
                    #    
                    #    #time.sleep(20000)
                    #    
                    #    try:
                    #        browser.find_element_by_css_selector('button[ng-click="HRStatusUpdate()"]').click()
                    #        print('hr-status-update-clicked')
                    #    except Exception as ex:
                    #        pass
                    
                    print(len(browser.window_handles))
                    
                    #if prepare_offer_flag.get_attribute('innerText') != 'Prepare Offer' and not prepare_offer_status_found: 
                    #    raise Exception
                    
                    
                    time.sleep(2)
                    
                    if len(browser.window_handles) < 3:
                        print("WINDOW-HANDLES >>>> LESS THAN 3")
                        browser.execute_script('''
                            document.querySelector('a.formsicon').click();
                        ''', candidate_name, req_id)
                        
                        #getting the window handle of new window that has opened
                        browser.switch_to.window(browser.window_handles[2])
                        time.sleep(2)
                        
                        Form_List = browser.find_element_by_xpath('//*[@id="ddlFormsList-input"]') #Dropdown for Form List         
                        browser.find_element_by_xpath('//*[@id="formlistDiv"]/div[1]/span[1]/a').click() #Clicking FormList Dropdown
                        formsSelect = Select(browser.find_element_by_id("ddlFormsList")) #Select Option List of Dropdown
                        
                        for o in formsSelect.options:
                            if o.get_attribute('value')=="181" and o.get_attribute('innerText').strip().lower()=="Offer Form".lower():  # checking If the Current Value is Equal to OFFER-FORM Option Value "181"
                                Form_List.send_keys(Keys.ENTER)
                                time.sleep(0.5)
                                break
                            else:
                                Form_List.send_keys(ARROW_DOWN)
                                time.sleep(0.5)
                                
                        time.sleep(1)                
                        browser.find_element_by_xpath('//*[@id="formlistDiv"]/div[1]/div[1]/button[1]').click() #Clicking Add Button
                        time.sleep(1)
                        
                        #getting the window handle of new window that has opened
                        #window_after = browser.window_handles[3]
                        browser.switch_to.window(browser.window_handles[3])
                        #time.sleep(3)
                        
                        #try:
                        #    continue_button = browser.find_element_by_id('continue')
                        #    continue_button.click() #Clicking Continue Button 
                        #except Exception as ex:
                            #Checking if the form Already Exist
                        #    error = browser.find_element_by_css_selector('div.ui-collapsible-content')
                        #    message = error.text
                        #    if(message == "This form has already been completed for this candidate for this requisition."):
                        #        browser.switch_to.window(browser.window_handles[3])
                        #        browser.close()
                        #        browser.switch_to.window(browser.window_handles[2])
                        #        browser.close()
                        #        #window_after = browser.window_handles[1]
                        #        browser.switch_to.window(browser.window_handles[1])
                        #        continue

                        time.sleep(2)
                    
                    #time.sleep(8)
                    
                    # forms_link = browser.find_element_by_css_selector('a[alt="Forms"]')
                    
                    #time.sleep(5)
                    
                    #getting the window handle of new window that has opened
                    #window_after = browser.window_handles[2]
                    #browser.switch_to.window(browser.window_handles[2])
                    #print(browser.title)
                    #time.sleep(2)
                    
                    #Form_List = browser.find_element_by_xpath('//*[@id="ddlFormsList-input"]') #Dropdown for Form List         
                    #browser.find_element_by_xpath('//*[@id="formlistDiv"]/div[1]/span[1]/a').click() #Clicking FormList Dropdown
                    #mySelect = Select(browser.find_element_by_id("ddlFormsList")) #Select Option List of Dropdown
                    
                    #for o in mySelect.options:
                    #    if o.get_attribute('value')=="181":  # checking If the Current Value is Equal to Offer form Option Value
                    #        Form_List.send_keys(Keys.ENTER)
                    #        break
                    #    else:
                    #        Form_List.send_keys(ARROW_DOWN)
                    #time.sleep(1)                
                    #browser.find_element_by_xpath('//*[@id="formlistDiv"]/div[1]/div[1]/button[1]').click() #Clicking Add Button
                    #time.sleep(3)
                    
                    #getting the window handle of new window that has opened
                    #window_after = browser.window_handles[3]
                    #browser.switch_to.window(window_after)
                    #time.sleep(3)
                    
                    #continue_button = browser.find_element_by_id('continue')
                    #continue_button.click() #Clicking Continue Button 
                    #time.sleep(2)

                    #Checking if the form Already Exist
                    try:
                        error = browser.find_element_by_css_selector('div.ui-collapsible-content')
                        message = error.text
                        if message == "This form has already been completed for this candidate for this requisition.":
                            browser.switch_to.window(browser.window_handles[3])
                            browser.close()
                            browser.switch_to.window(browser.window_handles[2])
                            browser.close()
                            #window_after = browser.window_handles[1]
                            browser.switch_to.window(browser.window_handles[1])
                            continue
                    except:
                        pass
                    
                    browser.switch_to.window(browser.window_handles[3])
                    time.sleep(3)
                    
                    #Job Profile
                    ui_content_elm = browser.find_element_by_css_selector('div.ui-content') 
                    if str(row["Job Profile:"]).strip().lower() != 'from req':
                        job_profile_input = browser.find_element_by_xpath('//*[@id="slt_3553_fname_slt_0_0_0_0-input"]')
                        job_profile_input.clear()
                        job_profile_input.send_keys(str(row['Job Profile:']))
                        time.sleep(1)
                        job_profile_input.send_keys(ARROW_DOWN)
                        job_profile_input.send_keys(Keys.ENTER)
                    
                    time.sleep(2)
                    
                    #Title
                    if str(row["Title:"]).strip().lower() != "from req":
                        title_input = browser.find_element_by_id("txt_3549_fname_txt_0_0_0_0")
                        title_input.clear()
                        title_input.send_keys(row['Title:'])
                        
                    time.sleep(2)
                    
                    #Offer Country/Group Dropdown
                    #ui_content_elm = browser.find_element_by_css_selector('div.ui-content')        
                    mySelect = Select(browser.find_element_by_id("slt_4849_fname_slt_0_0_0_0"))
                    country_list = browser.find_element_by_xpath('//*[@id="slt_4849_fname_slt_0_0_0_0-button"]')
                    country_str = str(row['Offer Country/ Group:']).strip().lower() # country/group value from sheet
                    if not country_str : 
	                    country_str = 'Cerner India'.strip().lower() # if whatever got from sheet is empty-string , make it 'cerner india'
                    for o in mySelect.options:
                        if o.get_attribute('innerText').strip().lower()==country_str:
                            country_list.send_keys(Keys.ENTER)
                            time.sleep(0.5)
                            break
                        else:
                            country_list.send_keys(ARROW_DOWN)
                            time.sleep(0.5)
                    
                    time.sleep(2)
                    
                    #Workspace Dropdown
                    workspace = browser.find_element_by_xpath('//*[@id="slt_49714_fname_slt_0_0_0_0-input"]')
                    workspace.send_keys(row['Workspace:'])
                    time.sleep(1)
                    workspace.send_keys(ARROW_DOWN)
                    workspace.send_keys(Keys.ENTER)

                    time.sleep(2)
                    
                    #Type of Transfer/Hire Dropdown
                    mySelect = Select(browser.find_element_by_id("slt_3562_fname_slt_0_0_0_0"))
                    transfer_type = browser.find_element_by_id("slt_3562_fname_slt_0_0_0_0-button")
                    for o in mySelect.options:
                        if o.get_attribute('value').lower().strip()==Hiring_Type[row['Type of transfer/hire:']].lower().strip():
                            transfer_type.send_keys(Keys.ENTER)
                            time.sleep(0.5)
                            break
                        else:
                            transfer_type.send_keys(ARROW_DOWN)
                            time.sleep(0.5)

                    time.sleep(2)
                    
                    #Recruiter
                    if str(row["Recruiter:"]).strip().lower() != "from req":
                        recruiterSelect = Select(browser.find_element_by_id("slt_6932_fname_slt_0_0_0_0"))
                        recruiter_btn = browser.find_element_by_id("slt_6932_fname_slt_0_0_0_0-button")
                        for o in recruiterSelect.options:
                            #print(o.get_attribute('innerText'))
                            if str(row['Recruiter:']).lower() in o.get_attribute('innerText').lower():
                                #recruiter_btn.send_keys(Keys.ENTER)
                                browser.execute_script("""
                                    //alert(arguments[0].value);
                                    document.getElementById("slt_6932_fname_slt_0_0_0_0").value = arguments[0].value;
                                    document.getElementById("slt_6932_fname_slt_0_0_0_0-button_text").innerText = arguments[0].innerText;
                                """, o)
                                time.sleep(0.5)
                                break
                            #else:
                            #    recruiter_btn.send_keys(ARROW_DOWN)
                            #    time.sleep(0.5)

                    time.sleep(2)
                            
                    #Working Hours Dropdown        
                    mySelect = Select(browser.find_element_by_id("slt_14860_fname_slt_0_0_0_0"))
                    shift_type = browser.find_element_by_id("slt_14860_fname_slt_0_0_0_0-button")
                    for o in mySelect.options:
                        if o.get_attribute('value').lower()==Shift_Type[row['Please indicate the working hours of the associate::']].lower():
                            shift_type.send_keys(Keys.ENTER)
                            time.sleep(0.5)
                            break
                        else:
                            shift_type.send_keys(ARROW_DOWN)
                            time.sleep(0.5)

                    time.sleep(2)
                    
                    #Virtual Position Dropdown
                    mySelect = Select(browser.find_element_by_id("slt_2926_fname_slt_0_0_0_0"))
                    virtual_position = browser.find_element_by_id("slt_2926_fname_slt_0_0_0_0-button")
                    for o in mySelect.options:
                        if o.get_attribute('value').lower()==Virtual_Position[row['Virtual Position:']].lower():
                            virtual_position.send_keys(Keys.ENTER)
                            time.sleep(0.5)
                            break
                        else:
                            virtual_position.send_keys(ARROW_DOWN)
                            time.sleep(0.5)

                    time.sleep(2)
                    
                    #Filling Offer Expiration Date and Target Start Date        
                    browser.find_element_by_xpath('//*[@id="dat_2810_fname_dat_0_0_0_0_dattxt"]').clear()
                    time.sleep(2)
                    browser.find_element_by_xpath('//*[@id="dat_2810_fname_dat_0_0_0_0_dattxt"]').send_keys(row['Offer Date'].strftime("%m/%d/%Y"))
                    
                    time.sleep(2)
                    
                    browser.find_element_by_xpath('//*[@id="dat_2811_fname_dat_0_0_0_0_dattxt"]').send_keys(row['Offer Expiration date'].strftime("%m/%d/%Y"))

                    time.sleep(2)

                    browser.find_element_by_xpath('//*[@id="dat_2812_fname_dat_0_0_0_0_dattxt"]').send_keys(row['Targeted Start date'].strftime("%m/%d/%Y"))

                    time.sleep(2)
                    
                    #CPP Dropdown
                    CPP_select = Select(browser.find_element_by_id("slt_2831_fname_slt_0_0_0_0"))
                    cpp = browser.find_element_by_id("slt_2831_fname_slt_0_0_0_0-button")
                    for o in CPP_select.options:
                        if o.get_attribute('value').lower()==CPP[row['CPP']].lower():
                            cpp.send_keys(Keys.ENTER)
                            time.sleep(0.5)
                            break
                        else:
                            cpp.send_keys(ARROW_DOWN)
                            time.sleep(0.5)

                    time.sleep(2)
                    
                    #Pay Group Dropdown        
                    pay_group_select = Select(browser.find_element_by_id("slt_2816_fname_slt_0_0_0_0"))
                    pay_group = browser.find_element_by_id("slt_2816_fname_slt_0_0_0_0-button")
                    for o in pay_group_select.options:
                        if pay_group.text.lower()==Pay_Group[row['Pay Group']].lower():
                            pay_group.send_keys(Keys.ENTER)
                            time.sleep(0.5)
                            break
                        else:
                            pay_group.send_keys(ARROW_DOWN)
                            time.sleep(0.5)

                    time.sleep(2)

                    #Currency Type Dropdown        
                    currency_select = Select(browser.find_element_by_id("slt_14886_fname_slt_0_0_0_0"))
                    currency = browser.find_element_by_id("slt_14886_fname_slt_0_0_0_0-button")
                    for o in currency_select.options:
                        if currency.text.lower()==row['Currency2'].lower():
                            currency.send_keys(Keys.ENTER)
                            time.sleep(0.5)
                            break
                        else:
                            currency.send_keys(ARROW_DOWN)
                            time.sleep(0.5)

                    time.sleep(2)
                    
                    #Currency Dropdown 2
                    currency2_select = Select(browser.find_element_by_id("slt_3451_fname_slt_0_0_0_0"))
                    currency2 = browser.find_element_by_id("slt_3451_fname_slt_0_0_0_0-button")
                    for o in currency2_select.options:
                        if currency.text.lower()==row['Currency'].lower():
                            currency2.send_keys(Keys.ENTER)
                            time.sleep(0.5)
                            break
                        else:
                            currency2.send_keys(ARROW_DOWN)
                            time.sleep(0.5)

                    time.sleep(2)
                    
                    #Salary Frequency
                    frequency_select = Select(browser.find_element_by_id("slt_2814_fname_slt_0_0_0_0"))
                    frequency = browser.find_element_by_id("slt_2814_fname_slt_0_0_0_0-button")
                    for o in frequency_select.options:
                        if frequency.text.lower()==row['Frequency'].lower():
                            frequency.send_keys(Keys.ENTER)
                            time.sleep(0.5)
                            break
                        else:
                            frequency.send_keys(ARROW_DOWN)
                            time.sleep(0.5)

                    time.sleep(2)
                            
                    #Offer Over 10% Plan Dropdown        
                    offer_over_select = Select(browser.find_element_by_id("slt_2934_fname_slt_0_0_0_0"))
                    offer_over = browser.find_element_by_id("slt_2934_fname_slt_0_0_0_0-button")
                    for o in offer_over_select.options:
                        if offer_over.text.lower()==row['Is offer over 10% of Plan Dollars:'].lower():
                            offer_over.send_keys(Keys.ENTER)
                            time.sleep(0.5)
                            break
                        else:
                            offer_over.send_keys(ARROW_DOWN)
                            time.sleep(0.5)

                    time.sleep(2)
                    
                    #Filling Total AGC Cash and CTC Amount
                    browser.find_element_by_xpath('//*[@id="txt_2813_fname_txt_0_0_0_0"]').send_keys(str(row['*Total Cash (AGC for India):']))
                    time.sleep(2)
                    browser.find_element_by_xpath('//*[@id="txt_14885_fname_txt_0_0_0_0"]').send_keys(str(row['*CTC Amount:']))
                    time.sleep(2)

                    
                    #Allowance Plan Dropdown
                    allowance_plan_select = Select(browser.find_element_by_id("slt_49725_fname_slt_0_0_0_0"))
                    allowance_plan = browser.find_element_by_id("slt_49725_fname_slt_0_0_0_0-button")
                    browser.find_element_by_id('slt_49725_fname_slt_0_0_0_0-button').click()
                    for o in allowance_plan_select.options:
                        a = o.get_attribute('value')[0:-3]
                        a = a.split("_")
                        val = " ".join(a)
                        if val.lower()==row['*Allowance Plan:'].lower():
                            allowance_plan.send_keys(Keys.ENTER)
                            time.sleep(0.5)
                            break
                        else:
                            allowance_plan.send_keys(ARROW_DOWN)
                            time.sleep(0.5)

                    time.sleep(2)
                    
                    #Associate Shift
                    associate_shift_select = Select(browser.find_element_by_id("slt_14865_fname_slt_0_0_0_0"))
                    associate_shift = browser.find_element_by_id("slt_14865_fname_slt_0_0_0_0-button")
                    for o in associate_shift_select.options:
                        if associate_shift.text.lower()==row['*Please indicate the shift of the associate:'].lower():
                            associate_shift.send_keys(Keys.ENTER)
                            time.sleep(0.5)
                            break
                        else:
                            associate_shift.send_keys(ARROW_DOWN)
                            time.sleep(0.5)

                    time.sleep(2)
                    
                    #Advance Payment
                    pay_advances_select = Select(browser.find_element_by_id("slt_3271_fname_slt_0_0_0_0"))
                    pay_advances = browser.find_element_by_id("slt_3271_fname_slt_0_0_0_0-button")
                    browser.find_element_by_id('slt_3271_fname_slt_0_0_0_0-button').click()
                    for o in pay_advances_select.options:
                        a = str(o.get_attribute('value')[0]) if o.get_attribute('value') else None
                        print(a)
                        if a and Advance[a].lower()==row['Type of Advance'].lower():
                            pay_advances.send_keys(Keys.ENTER)
                            time.sleep(0.5)
                            break
                        else:
                            pay_advances.send_keys(ARROW_DOWN)
                            time.sleep(0.5)

                    time.sleep(2)
                            
                    #Stock Option Dropdown
                    stock_select = Select(browser.find_element_by_id("slt_3270_fname_slt_0_0_0_0"))
                    stock = browser.find_element_by_id("slt_3270_fname_slt_0_0_0_0-button")
                    for o in stock_select.options:
                        if stock.text.lower()==row['*Stock Options:'].lower():
                            stock.send_keys(Keys.ENTER)
                            time.sleep(0.5)
                            break
                        else:
                            stock.send_keys(ARROW_DOWN)
                            time.sleep(0.5)

                    time.sleep(2)
                            
                    
                    if row['Type of transfer/hire:'].lower().strip() != 'transfer':
                        #Employment Bonus Dropdown       
                        employment_bonus_select = Select(browser.find_element_by_id("slt_3264_fname_slt_0_0_0_0"))
                        employment_bonus = browser.find_element_by_id("slt_3264_fname_slt_0_0_0_0-button")
                        emp_bonus = str(row['Type of Employment Bonus:']).strip().lower()
                        if not emp_bonus:
                            emp_bonus = 'none'
                        for o in employment_bonus_select.options:
                            if employment_bonus.text.strip().lower()==emp_bonus:
                                employment_bonus.send_keys(Keys.ENTER)
                                time.sleep(0.5)
                                break
                            else:
                                employment_bonus.send_keys(ARROW_DOWN)
                                time.sleep(0.5)
                                
                        time.sleep(2)

                    time.sleep(2)
                    
                    #Relocation Assistance Dropdown
                    relocation_select = Select(browser.find_element_by_id("slt_14411_fname_slt_0_0_0_0"))
                    relocation = browser.find_element_by_id("slt_14411_fname_slt_0_0_0_0-button")
                    for o in relocation_select.options:
                        if relocation.text.lower()==row['*Are they eligible for relo?:'].lower():
                            relocation.send_keys(Keys.ENTER)
                            time.sleep(0.5)
                            break
                        else:
                            relocation.send_keys(ARROW_DOWN)
                            time.sleep(0.5)

                    time.sleep(2)
                    
                    #Type of Benefits
                    associate_benefits_select = Select(browser.find_element_by_id("slt_3268_fname_slt_0_0_0_0"))
                    associate_benefits = browser.find_element_by_id("slt_3268_fname_slt_0_0_0_0-button")
                    for o in associate_benefits_select.options:
                        if associate_benefits.text.lower()==row['*Type of Benefits:'].lower():
                            associate_benefits.send_keys(Keys.ENTER)
                            time.sleep(0.5)
                            break
                        else:
                            associate_benefits.send_keys(ARROW_DOWN)
                            time.sleep(0.5)

                    time.sleep(2)
                    
                    #Filling No. of PTO
                    browser.find_element_by_xpath('//*[@id="txt_3427_fname_txt_0_0_0_0"]').send_keys(row['No. Of Days'])

                    time.sleep(2)
                    
                    #Car Eligigible or not
                    car_eligible_select = Select(browser.find_element_by_id("slt_3428_fname_slt_0_0_0_0"))
                    car_eligible = browser.find_element_by_id("slt_3428_fname_slt_0_0_0_0-button")
                    for o in car_eligible_select.options:
                        if car_eligible.text.lower()==row['*Are they eligible for a car?:'].lower():
                            car_eligible.send_keys(Keys.ENTER)
                            time.sleep(0.5)
                            break
                        else:
                            car_eligible.send_keys(ARROW_DOWN)
                            time.sleep(0.5)

                    time.sleep(2)
                            
                    #Special Employment Agreement
                    special_emp_select = Select(browser.find_element_by_id("slt_2833_fname_slt_0_0_0_0"))
                    special_emp = browser.find_element_by_id("slt_2833_fname_slt_0_0_0_0-button")
                    for o in special_emp_select.options:
                        if special_emp.text.lower()==row['*Special employment agreement:'].lower():
                            special_emp.send_keys(Keys.ENTER)
                            time.sleep(0.5)
                            break
                        else:
                            special_emp.send_keys(ARROW_DOWN)
                            time.sleep(0.5)

                    time.sleep(2)
                    
                    # Versant Test Type & Scores 
                    versant_test_type_select = Select(browser.find_element_by_id("slt_51079_fname_slt_0_0_0_0"))
                    versant_test_type_button = browser.find_element_by_id("slt_51079_fname_slt_0_0_0_0-button")

                    time.sleep(2)
                    
                    browser.find_element_by_xpath('//*[@id="txt_44964_fname_txt_0_0_0_0"]').send_keys(str(int(row['VERSANT_SCORE'])))
                    
                    
                    if 'Mobile' in row['Versant Test Type:'] and '17' in row['Versant Test Type:']:
                        for o in versant_test_type_select.options:
                            print(o.get_attribute('value'))
                            if o.get_attribute('value').lower()=='mobile':
                                versant_test_type_button.send_keys(Keys.ENTER)
                                time.sleep(0.5)
                                break
                            else:
                                versant_test_type_button.send_keys(ARROW_DOWN)
                                time.sleep(0.5)

                        browser.find_element_by_css_selector("input[dbfieldname='MASTERY_SCORE']").send_keys(str(int(row['MASTERY_SCORE'])))
                        time.sleep(0.5)
                        browser.find_element_by_css_selector("input[dbfieldname='VOCAB_SCORE']").send_keys(str(int(row['VOCAB_SCORE'])))
                        time.sleep(0.5)            
                        browser.find_element_by_css_selector("input[dbfieldname='FLUENCY_SCORE']").send_keys(str(int(row['FLUENCY_SCORE'])))
                        time.sleep(0.5)            
                        browser.find_element_by_css_selector("input[dbfieldname='PRONUNCIATION_SCORE']").send_keys(str(int(row['PRONUNCIATION_SCORE'])))
                        time.sleep(0.5)
                        

                    elif 'Full' in row['Versant Test Type:'] and '50' in row['Versant Test Type:']:
                        for o in versant_test_type_select.options:
                            print(o.get_attribute('value'))
                            if o.get_attribute('value').lower()=='full':
                                versant_test_type_button.send_keys(Keys.ENTER)
                                time.sleep(0.5)                    
                                break
                            else:
                                versant_test_type_button.send_keys(ARROW_DOWN)  
                                time.sleep(0.5)

                        browser.find_element_by_css_selector("input[dbfieldname='SPEAKING_SCORE']").send_keys(str(int(row['SPEAKING_SCORE'])))
                        time.sleep(0.5)            
                        browser.find_element_by_css_selector("input[dbfieldname='LISTENING_SCORE']").send_keys(str(int(row['LISTENING_SCORE'])))
                        time.sleep(0.5)            
                        browser.find_element_by_css_selector("input[dbfieldname='WRITING_SCORE']").send_keys(str(int(row['WRITING_SCORE'])))
                        time.sleep(0.5)            
                        browser.find_element_by_css_selector("input[dbfieldname='READING_SCORE']").send_keys(str(int(row['READING_SCORE'])))
                        time.sleep(0.5)
                                       
                    time.sleep(2)
                    
                    # Decision Makers
                    browser.find_element_by_xpath('//*[@id="txt_3133_fname_txt_0_0_0_0"]').send_keys(row['Decision Maker(s):'])
                    time.sleep(2)
                      
                    if row['Type of transfer/hire:'].lower().strip() != 'transfer':                      
                        #Is Person old Cerner Client
                        is_cerner_client_select = Select(browser.find_element_by_id("slt_4534_fname_slt_0_0_0_0"))
                        is_cerner_client = browser.find_element_by_id("slt_4534_fname_slt_0_0_0_0-button")
                        for o in is_cerner_client_select.options:
                            if is_cerner_client.text.lower()==row['*Is the person from a Cerner client?:'].lower():
                                is_cerner_client.send_keys(Keys.ENTER)
                                time.sleep(0.5)                
                                break
                            else:
                                is_cerner_client.send_keys(ARROW_DOWN)
                                time.sleep(0.5)
                            
                        time.sleep(2)            
                    
                        #Is Minimum gpa required
                        min_gpa_elem = browser.find_element_by_id("slt_14427_fname_slt_0_0_0_0")
                        if min_gpa_elem:			
                            minimum_GPA_select = Select(browser.find_element_by_id("slt_14427_fname_slt_0_0_0_0"))
                            minimum_GPA = browser.find_element_by_id("slt_14427_fname_slt_0_0_0_0-button")
                            for o in minimum_GPA_select.options:
                                if minimum_GPA.text.lower()==row['*Does this offer depend on minimum GPA requirements?:'].lower():
                                    minimum_GPA.send_keys(Keys.ENTER)
                                    time.sleep(0.5)                
                                    break
                                else:
                                    minimum_GPA.send_keys(ARROW_DOWN)
                                    time.sleep(0.5)
                            time.sleep(2)    
                        time.sleep(2)
                            
                        #Is this a Contract Conversion        
                        contract_conversion_select = Select(browser.find_element_by_id("slt_22003_fname_slt_0_0_0_0"))
                        contract_conversion = browser.find_element_by_id("slt_22003_fname_slt_0_0_0_0-button")
                        for o in contract_conversion_select.options:
                            if contract_conversion.text.lower()==row['*Is this a contractor conversion?:'].lower():
                                contract_conversion.send_keys(Keys.ENTER)
                                time.sleep(0.5)                
                                break
                            else:
                                contract_conversion.send_keys(ARROW_DOWN)
                                time.sleep(0.5)                
                    
                        time.sleep(2)
                    
                    time.sleep(2)
                    
                    #collapse_arrow = browser.find_element_by_css_selector('#div-16941 button.collapseIconArrow')
                    recruit_lead_selection = browser.find_element_by_css_selector('label[for="chk_248_approvaltitles_bypass"]')
                    browser.execute_script("arguments[0].scrollIntoView();", recruit_lead_selection)

                    #Selection 3 byPass Checkbox
                    browser.execute_script('''
                        //document.querySelector('label[for="chk_28_approvaltitles_bypass"]').click();
                        //document.querySelector('label[for="chk_30_approvaltitles_bypass"]').click();
                        //document.querySelector('label[for="chk_37_approvaltitles_bypass"]').click();
                        document.querySelector('label[for="chk_248_approvaltitles_bypass"]').click();
                    ''', 'None')

                    time.sleep(2)

                    #Position Management Approval
                    position_management_select = Select(browser.find_element_by_id("slt_29_approvaltitles_txt_0_1"))
                    position_management = browser.find_element_by_id("slt_29_approvaltitles_txt_0_1-button")
                    for o in position_management_select.options:
                        if position_management.text.lower()==row['*Position Management Approval'].lower():
                            position_management.send_keys(Keys.ENTER)
                            time.sleep(0.5)                
                            break
                        else:
                            position_management.send_keys(ARROW_DOWN)
                            time.sleep(0.5)                

                    time.sleep(2)
                     
                    #System User
                    System_user = browser.find_element_by_xpath('//*[@id="ssd_textareausers-input"]')
                    System_user.send_keys(str(int(row['Specific system users'])))
                    time.sleep(3)
                    System_user.send_keys(Keys.ARROW_UP)
                    System_user.send_keys(Keys.ENTER)

                    time.sleep(2)
                    '''
                    browser.find_element_by_id("saveroute").click()
                    
                    time.sleep(2)
                    
                    browser.find_element_by_css_selector('.ts-btn-primary').click()
                    
                    time.sleep(3)
                    
                    alert_success_found = browser.find_element_by_css_selector('.ts-alert-success')
                    
                    if alert_success_found:
                        sht.range(status_row).value = 'COMPLETED'
                        wb.save()
                        time.sleep(2)
                        browser.find_element_by_css_selector('.ts-btn-primary').click()
                        time.sleep(2)
                    '''

                    save_submit_final = browser.find_element_by_id("saveroute")
                    save_submit_final.click()
                    
                    time.sleep(3)
                    
                    #print(browser.find_element_by_css_selector('div.modal-body').get_attribute('innerHTML'))
                    
                    browser.find_element_by_css_selector('div.modal-body')
                    
                    final_modal_approve = browser.find_element_by_css_selector('button.ts-btn-primary')
                    print(final_modal_approve.get_attribute('innerHTML'))
                    #final_modal_approve.click()
                    
                    browser.execute_script('''
                        document.querySelector('div.modal-body button.ts-btn-primary').click();
                    ''')
                    
                    time.sleep(3)
                    
                    sht.range(status_row).value = 'COMPLETED'
                    wb.save()
                    time.sleep(2)
                    
                    for i in range(len(browser.window_handles)-1, 1, -1):  
                        #print(i)
                        browser.switch_to.window(browser.window_handles[i])
                        browser.close()            

                    time.sleep(2)

                    browser.switch_to.window(browser.window_handles[1])
                
                except Exception as e:
                    print(str(e))
                    print(row)
                    sht.range(status_row).value = 'FAILED'
                    wb.save()

                    for i in range(len(browser.window_handles)-1, 1, -1):  
                        #print(i)
                        browser.switch_to.window(browser.window_handles[i])
                        browser.close()            

                    time.sleep(2)

                    browser.switch_to.window(browser.window_handles[1])            
    #     browser.switch_to.window(browser.window_handles[2])
    #     browser.close()
                                           
    except TimeoutException:
        print("Timed out waiting for page to load")


    # In[ ]:


if __name__ == "__main__":
    offer_form_automate()
