import webbrowser
import time
import xlwings as xw
import pandas as pd
from mailmerge import MailMerge
from datetime import date
# from robobrowser import RoboBrowser
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains


DEPARTURE_PATH = r"\\Cernfs05\functional\HR\HRSC\Departure"
#DEPARTURE_PATH = r"Z:\\Departure"
#DEPARTURE_PATH = r"Z:\HR\HRSC\Departure"
FTE_TEMPLATE = DEPARTURE_PATH + "\Depart Letter Template for Automation\FTE_Experience_Template.docx"
INTERN_TEMPLATE = DEPARTURE_PATH + "\Depart Letter Template for Automation\INTERN_Experience_Template.docx"
EXCEL_SHEET_TO_READ = DEPARTURE_PATH + "\DepartLetters.xlsm"
FOLDER_TO_ADD_LETTERS = DEPARTURE_PATH + "\TestLetters"
CHROME_DRIVER_PATH = "C:\\Program Files (x86)\\chromedriver_win32\\chromedriver"
ASK_HR_URL = "https://ask.cerner.com"


@xw.sub
def generate_experience_letters():
    """Generates experience letters in docx format for all rows"""
    #wb = xw.Book.caller()
    #cell_range = wb.app.selection
    #print('cell range:')
    #print(cell_range)

    df = pd.read_excel(EXCEL_SHEET_TO_READ)

    #selected_df = cell_range.options(pd.DataFrame).value
    #row_1_selected = list(selected_df)
    #rows_all_selected = [list(selected_row) for selected_i, selected_row in selected_df.iterrows()]
    #rows_all_selected.insert(0, row_1_selected)
    #print(rows_all_selected)

    #selected_df_final = pd.DataFrame(rows_all_selected, columns=list(df)[1:])
    #print(selected_df_final)
    selected_df_final = df[['Assoc ID', 'Operator ID', 'Associate Name', 'Departure Date', 'Title', 'Last Hire Date']]
    print(selected_df_final)
    empty_df = pd.DataFrame()

    #try:
    for row_index, row in selected_df_final.iterrows():
        # print(row)
        if 'Intern'.lower() in row['Title'].lower():
            document = MailMerge(INTERN_TEMPLATE)
            document.merge(
                #todays_date='{:%B %d, %Y}'.format(date.today()),
                operator_id=str(row['Operator ID']),
                associate_name=str(row['Associate Name']),
                employed_from_date='{:%B %d, %Y}'.format(row['Last Hire Date']),
                employed_to_date='{:%B %d, %Y}'.format(row['Departure Date'])
            )
            # have all these fields as MergeFields in INTERN_TEMPLATE

        else:
            document = MailMerge(FTE_TEMPLATE)
            document.merge(
                #todays_date='{:%B %d, %Y}'.format(date.today()),
                # associate_id=str(row['Assoc ID']),
                associate_id=str(row['Operator ID'][2:]),
                associate_name=str(row['Associate Name']),
                job_title=str(row['Title']),
                last_joining_date='{:%B %d, %Y}'.format(row['Last Hire Date']),
                departure_date='{:%B %d, %Y}'.format(row['Departure Date'])
            )
            # have all these fields as MergeFields in FTE_TEMPLATE

        document.write(FOLDER_TO_ADD_LETTERS + '\{}.docx'.format(row['Operator ID'][2:]))
    #empty_df.to_excel(EXCEL_SHEET_TO_READ, 'Sheet1', columns=list(df), index=False)
    #except Exception as e:
    #    return "There was some problem while performing the operation."+str(e)
    return "Generated Experience letters successfully !"


@xw.sub
def submit_departure_sr():
    """Submit Departure SRs for all rows"""
    browser = webdriver.Chrome(CHROME_DRIVER_PATH)
    browser.get(ASK_HR_URL)
    timeout = 30

    df = pd.read_excel(EXCEL_SHEET_TO_READ)
    df.dropna(how='all', inplace=True)
    print(df[['Assoc ID', 'Associate Name', 'Operator ID', 'Departure Date']])

    try:
        catalog_present = EC.presence_of_all_elements_located((By.CLASS_NAME, 'catalog-item'))
        WebDriverWait(browser, timeout).until(catalog_present)
        link_elems = browser.find_elements_by_css_selector('.catalog-item')

        for row_index, row in df.iterrows():
            try:
                clear_search_button = browser.find_element_by_css_selector('.search-field-box__clear-btn')
                clear_search_button.click()
                time.sleep(2)
            except Exception as e:
                print(e)
                pass
                
            search_box_present = EC.presence_of_element_located((By.CSS_SELECTOR, '.catalog-search__field'))
            WebDriverWait(browser, timeout).until(search_box_present)
            search_box = browser.find_element_by_css_selector('.catalog-search__field')
            search_box.send_keys('india - depart')            
            
            if row['Operator ID'] :
                india_depart_action_presence = EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'div[title = "India - Depart"]'))
                WebDriverWait(browser, timeout).until(india_depart_action_presence)
                india_depart_action = browser.find_element_by_css_selector('div[title = "India - Depart"]')
                india_depart_action.click()
                time.sleep(1)

                form_present = EC.presence_of_element_located((By.NAME, 'createSRForm'))
                WebDriverWait(browser, timeout).until(form_present)
                time.sleep(1)
                
                form_sr_create_modal = browser.find_element_by_css_selector('.modal')
                form_sr_create_modal.send_keys(Keys.DOWN)
                form_sr_create_modal.send_keys(Keys.DOWN)
                time.sleep(1)
                
                details_textarea_presence = EC.presence_of_element_located((By.CSS_SELECTOR, 'textarea[name="QSGCNOOBDLHLSANZRQ0XU3I4PSD5HM"]'))
                WebDriverWait(browser, timeout).until(details_textarea_presence)
                details_textarea = browser.find_element_by_css_selector('textarea[name="QSGCNOOBDLHLSANZRQ0XU3I4PSD5HM"]')
                details_textarea.send_keys(str(row['Operator ID'])[2:] + "\n\n"+'{:%d-%b-%Y}'.format(row['Departure Date']))
                
                form_sr_create_modal.send_keys(Keys.DOWN)
                form_sr_create_modal.send_keys(Keys.DOWN)            
                time.sleep(1)
                
                id_radio_presence = EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="radio"][value="ID"]'))
                WebDriverWait(browser, timeout).until(id_radio_presence)
                id_radio_button = browser.find_element_by_css_selector('input[type="radio"][value="ID"]')
                # print(id_radio_button.text())
                id_radio_button.click()
                
                form_sr_create_modal.send_keys(Keys.DOWN)
                time.sleep(1)
                
                id_filter_input_presence = EC.presence_of_element_located((By.NAME, 'QSGIW2WRIDVLKANZRQ3QHOJZAAE0KO'))
                WebDriverWait(browser, timeout).until(id_filter_input_presence)
                id_filter_input = browser.find_element_by_name('QSGIW2WRIDVLKANZRQ3QHOJZAAE0KO')
                id_filter_text = str(row['Operator ID'])[2:]
                id_filter_input.send_keys(id_filter_text)
                       
                form_sr_create_modal.send_keys(Keys.PAGE_DOWN)
                form_sr_create_modal.send_keys(Keys.UP)
                
                time.sleep(3)
                
                associate_id_select_presence = EC.presence_of_element_located((By.NAME, 'QSGIW2WRIDVLKANZRQ3SHQ4CHEE0Q9'))
                WebDriverWait(browser, timeout).until(associate_id_select_presence)
                time.sleep(1)
                associate_id_select = browser.find_element_by_name('QSGIW2WRIDVLKANZRQ3SHQ4CHEE0Q9')
                # associate_id_select.send_keys(str(row['Operator ID']))
                associate_id_select.click()
                
                time.sleep(1)
                
                assoc_id_options_presence = EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'select[name="QSGIW2WRIDVLKANZRQ3SHQ4CHEE0Q9"] > option'))
                WebDriverWait(browser, timeout).until(assoc_id_options_presence)
                time.sleep(1)
                options_assoc_id = browser.find_elements_by_css_selector('select[name="QSGIW2WRIDVLKANZRQ3SHQ4CHEE0Q9"] > option')
                options_assoc_id[1].click()
                
                time.sleep(2)            
                
                associate_lastname_select_presence = EC.presence_of_element_located((By.NAME, 'QSGIW2WRIDVLKANZRQ3UHSBNJ2E0VZ'))
                WebDriverWait(browser, timeout).until(associate_lastname_select_presence)
                time.sleep(1)
                associate_lastname_select = browser.find_element_by_name('QSGIW2WRIDVLKANZRQ3UHSBNJ2E0VZ')
                associate_lastname_select.click()
                
                time.sleep(2)
                
                # lastname_options_presence = EC.presence_of_element_located(
                #     (By.CSS_SELECTOR, 'select[name="QSGIW2WRIDVLKANZRQ3UHSBNJ2E0VZ"] > option'))
                # WebDriverWait(browser, timeout).until(lastname_options_presence)
                # time.sleep(1)
                options_lastname = browser.find_elements_by_css_selector('select[name="QSGIW2WRIDVLKANZRQ3UHSBNJ2E0VZ"] > option')
                options_lastname[1].click()
                
                time.sleep(2)
                
                # associate_firstname_select_presence = EC.presence_of_element_located((By.NAME, 'QSGIW2WRIDVLKANZRQ3YHW2TO6E1HU'))
                # WebDriverWait(browser, timeout).until(associate_firstname_select_presence)
                # time.sleep(1)
                associate_firstname_select = browser.find_element_by_name('QSGIW2WRIDVLKANZRQ3YHW2TO6E1HU')
                associate_firstname_select.click()
                
                time.sleep(2)
                
                # firstname_options_presence = EC.presence_of_element_located(
                #     (By.CSS_SELECTOR, 'select[name="QSGIW2WRIDVLKANZRQ3YHW2TO6E1HU"] > option'))
                # WebDriverWait(browser, timeout).until(firstname_options_presence)
                # time.sleep(1)
                options_firstname = browser.find_elements_by_css_selector('select[name="QSGIW2WRIDVLKANZRQ3YHW2TO6E1HU"] > option')
                options_firstname[1].click()

                time.sleep(2)
                
                # associate_site_select_presence = EC.presence_of_element_located((By.NAME, 'QSGCNOOBDLHLSANZRXTXOAYIOANOE8'))
                # WebDriverWait(browser, timeout).until(associate_firstname_select_presence)
                # time.sleep(2)
                associate_site_select = browser.find_element_by_name('QSGCNOOBDLHLSANZRXTXOAYIOANOE8')
                associate_site_select.click()
                
                time.sleep(2)
                
                # site_options_presence = EC.presence_of_element_located(
                #     (By.CSS_SELECTOR, 'select[name="QSGCNOOBDLHLSANZRXTXOAYIOANOE8"] > option'))
                # WebDriverWait(browser, timeout).until(site_options_presence)
                # time.sleep(1)
                options_site = browser.find_elements_by_css_selector('select[name="QSGCNOOBDLHLSANZRXTXOAYIOANOE8"] > option')
                options_site[1].click()

                time.sleep(2)
                
                site_span = browser.find_element_by_css_selector('button[name="QSGCNOOBDLHLSANZRXTXOAYIOANOE8"] > span')
                site_location = site_span.text
                # print(site_span.get_attribute('innerHtml'))
                
                if 'KOL' in site_location :
                    # kolkata_radio_presence = EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="radio"][value="KLKT"]'))
                    # WebDriverWait(browser, timeout).until(kolkata_radio_presence)
                    kolkata_radio_button = browser.find_element_by_css_selector('input[type="radio"][value="KLKT"]')
                    kolkata_radio_button.click()
                elif 'MTP' in site_location :
                    # manyata_radio_presence = EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="radio"][value="CHS"]'))
                    # WebDriverWait(browser, timeout).until(manyata_radio_presence)
                    manyata_radio_button = browser.find_element_by_css_selector('input[type="radio"][value="CHS"]')
                    manyata_radio_button.click()
                elif 'GTP' in site_location :
                    # global_radio_presence = EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="radio"][value="BNGL"]'))
                    # WebDriverWait(browser, timeout).until(global_radio_presence)
                    global_radio_button = browser.find_element_by_css_selector('input[type="radio"][value="BNGL"]')
                    global_radio_button.click()
                    
                # depart_label_field = browser.find_element_by_css_selector('div.srd-question-label-container__left')
                depart_input_field = browser.find_element_by_css_selector('div.srd-question-datetimepicker__input-container > input[type="text"]')
                # depart_input_field.click()
                # depart_input_field.send_keys('{:%b %d, %Y}'.format(row['Departure Date'])).perform()
                depart_date_str = '{:%b %d, %Y}'.format(row['Departure Date'])
                browser.execute_script('document.querySelector("div.srd-question-datetimepicker__input-container > input[type=\'text\']").removeAttribute("readonly");')
                # document.querySelector("div.srd-question-datetimepicker__input-container > input[type=\'text\']").value = \'{}\''.format(depart_date_str))
                depart_input_field_enabled = browser.find_element_by_css_selector('div.srd-question-datetimepicker__input-container > input[type="text"]')
                # depart_input_field_enabled.send_keys(depart_date_str)
                depart_input_field.send_keys(depart_date_str)
                # ActionChains(browser).move_to_element(depart_input_field).click().send_keys().perform()

                                            
                # let the user see that you have entered these values
                time.sleep(2)
                browser.execute_script('document.querySelector("div.srd-question-datetimepicker__input-container > input[type=\'text\']").setAttribute("readonly", true);')
                
                # depart_label_field.click()

                # then close the modal
                # form_sr_create_modal.send_keys(Keys.HOME)
                # time.sleep(2)
                # close_button = form_sr_create_modal.find_element_by_css_selector('glyph-icon.modal__close')
                # close_button.click()
                
                form_sr_create_modal.send_keys(Keys.DOWN)
                modal_footer = browser.find_element_by_css_selector('.modal-footer')
                modal_footer.click()
                time.sleep(2)
                
                submit_button = form_sr_create_modal.find_element_by_css_selector('button.btn-primary[ng-click="createSR($uibModal)"]')
                submit_button.click()

                time.sleep(15)
                

        time.sleep(2)

    except TimeoutException:
        print("Timed out waiting for page to load")


if __name__ == "__main__":
    # xw.Book('DepartLetters.xlsm').set_mock_caller()
    #generate_experience_letters()
    submit_departure_sr()

