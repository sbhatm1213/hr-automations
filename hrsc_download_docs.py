import os
import webbrowser
import time
import glob
import json
import pdfkit
import datetime as dt
from shutil import copy
from shutil import move
import xlwings as xw
import pandas as pd
from mailmerge import MailMerge
from PyPDF2 import PdfFileMerger
# from robobrowser import RoboBrowser
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities



HIRING_QUEUE_PATH = r"\\Cernfs05\functional\HR\HRSC\HIRING QUEUE\automation-resume-download\Python file"
#HIRING_QUEUE_PATH = r"Z:\\HIRING QUEUE\automation-resume-download\Python file"
INDIA_HIRING_COMPLETED_SHEET = HIRING_QUEUE_PATH + r"\\automation-resume-download.xlsm"
INDIA_HIRING_COMPLETED_SHEET = HIRING_QUEUE_PATH + r"\\final-automation.xlsm"
CHROME_DRIVER_PATH = "C:\\Program Files (x86)\\chromedriver_win32\\chromedriver"
WKHTML_PATH = "C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe"
HR_LINKS_PORTAL_URL = r"http://grdhrdbwhq00.northamerica.cerner.net/ReportServer?/HR+Links+Portal/HR_Links_Portal&rs:Command=Render&rc:Toolbar=False"
#DOWNLOADS_DIR = r"C:\\Users\\KY047036\\Downloads"
#DOWNLOADS_DIR = r"C:\\Users\\DS062683\\Downloads"
DOWNLOADS_DIR = os.path.expanduser("~/Downloads")
#DOWNLOADS_COPY_TO_PATH = r"Z:\\HIRING QUEUE\automation-resume-download"
RESUMES_DOWNLOADS_COPY_TO = HIRING_QUEUE_PATH + r"\automation_resumes"
OFFERS_DOWNLOADS_COPY_TO = HIRING_QUEUE_PATH + r"\automation_offers"


def download_str(pdf_raw_url):
    download_js_str = r"""
        setTimeout(function() {
            url = """ + pdf_raw_url + r""";
               downloadFile(url);                       
            }, 2000);
            
        window.downloadFile = function (sUrl) {

            //iOS devices do not support downloading. We have to inform user about this.
            if (/(iP)/g.test(navigator.userAgent)) {
               //alert('Your device does not support files downloading. Please try again in desktop browser.');
               window.open(sUrl, '_blank');
               return false;
            }

            //If in Chrome or Safari - download via virtual link click
            if (window.downloadFile.isChrome || window.downloadFile.isSafari) {
                //Creating new link node.
                var link = document.createElement('a');
                link.href = sUrl;
                link.setAttribute('target','_blank');

                if (link.download !== undefined) {
                    //Set HTML5 download attribute. This will prevent file from opening if supported.
                    var fileName = sUrl.substring(sUrl.lastIndexOf('/') + 1, sUrl.length);
                    //link.download = fileName;
                    link.download = 'UserFile';
                }

                //Dispatching click event.
                if (document.createEvent) {
                    var e = document.createEvent('MouseEvents');
                    e.initEvent('click', true, true);
                    link.dispatchEvent(e);
                    return true;
                }
            }

            // Force file download (whether supported by server).
            /*
            if (sUrl.indexOf('?') === -1) {
                sUrl += '?download';
            }
            */
            window.open(sUrl, '_blank');
            setTimeout(function(){
                //if (sUrl.indexOf('?') === -1) {
                //    sUrl += '?download';
                //}
                document.location += '?download'
            }, 20000)

            return true;
        }

        window.downloadFile.isChrome = navigator.userAgent.toLowerCase().indexOf('chrome') > -1;
        window.downloadFile.isSafari = navigator.userAgent.toLowerCase().indexOf('safari') > -1;                
                            
    """
    return download_js_str

    
def resume_download(browser):
                
    resume_tab_link = browser.find_element_by_xpath('//a[@href="#resumeTab"]')
    resume_tab_link.click()
    
    time.sleep(3)
    
    resume_open_link = browser.find_element_by_css_selector('a.pdf-icon')
    resume_open_link.click()
                    
    time.sleep(3)
    
    browser.switch_to.window(browser.window_handles[3])
    
    time.sleep(3)
    
    pdf_embed_tag = browser.find_element_by_css_selector('embed[type="application/pdf"]')
    time.sleep(12)

    pdf_href = pdf_embed_tag.get_attribute('src')
    print(pdf_href)
    pdf_href = browser.execute_script('''
        //document.querySelector('viewer-pdf-toolbar').shadowRoot.querySelector('cr-icon-button#download').click()

        
        return window.location.href;
    ''')
    
    print(str(pdf_href))
    
    print("=============================Download Resume Logic Starts===================================")

                            
    browser.switch_to.window(browser.window_handles[1])
    
    time.sleep(4)
                
    #searching_radio = browser.find_element_by_id('SearchReqs')
    #searching_radio.send_keys(Keys.CONTROL + "t")
    
    pdf_raw_url = "%r"%pdf_href
    
    #browser.get(pdf_href)
    
    #time.sleep(500)
    
    download_script_str = download_str(pdf_raw_url)
                
    #print(download_script_str)
    
    browser.execute_script(download_script_str)
    
    # logic 6 end

    print("==========================================Download Resume Logic Ends===================================================")
        
    
def offer_letter_download(browser, row):
    #browser.execute_script("""document.querySelector('img[alt="View candidates"]').click()""")
    
    expand_activity_log = browser.find_element_by_css_selector('a#expandCollapseActLog')
    expand_activity_log.click()
    
    time.sleep(5)
    
    browser.execute_script("""
       /*
       var grid_headers = document.querySelectorAll('div[role="columnheader"]');
       for(var i=0; i < grid_headers.length; i++){
            if(grid_headers[i].innerText.includes('Date')){
                var col_menu_down = grid_headers[i].querySelector('div[aria-label="Column Menu"] i.ui-grid-icon-angle-down');
                col_menu_down.click();
                var sort_desc_btn = document.querySelector('i.ui-grid-icon-sort-alt-down');
                sort_desc_btn.click();
            }
       }
       */
    """)
    
    time.sleep(0.5)                      
                            
    browser.execute_script("""
        //var actions_table = document.querySelectorAll('div#actionlogGrid ui-grid-row ');
        /*
        var objDiv = document.querySelector("div#actionlogGrid");
        objDiv.scrollTop = objDiv.scrollHeight;
        */
    """)
    
    time.sleep(0.5)
    
    browser.execute_script("""                
        /*
        var trs = document.querySelectorAll('div#actionlogGrid .ui-grid-row ');
        for(var i = 0; i < trs.length; i++){
            var tds = trs[i].querySelectorAll('.ui-grid-cell a.viewattachment');                    
            
            for(var j = 0; j < tds.length; j++){
                if(tds[j].innerText.includes('Offer Letter')){
                    var offer_link = trs[i].querySelector('.ui-grid-cell a.viewattachment');
                    var row_cells = trs[i].querySelectorAll('.ui-grid-cell-contents');
                    var view_folder_req = false;
                    for(var k = 0; k < row_cells.length; k++){
                        if(!view_folder_req){
                            var view_folder_link = row_cells[k].querySelector('a.viewfolder');
                            if(view_folder_link != undefined && view_folder_link != null){
                                view_folder_req = view_folder_link.innerText.includes(arguments[2]);
                            }
                        }
                    }
                    console.log(view_folder_link);
                    console.log(view_folder_req);
                    
                    if(view_folder_req 
                        && offer_link.innerText.includes('Offer Letter') 
                        && (offer_link.innerText.includes(arguments[0]) 
                        || offer_link.innerText.includes(arguments[1]))
                        && offer_link.innerText.includes('pdf')){
                            
                            console.log(trs[i]);
                            offer_link.click();
                            //j = tds.length;
                            i = trs.length;
                            //break;
                    }
                }
            }
        }
        */                                       
    """, row['Primary_First_Name'], row['Primary_Last_Name'], row['Job_Opening'])
    
    time.sleep(0.5)
    
    attachments_tab = browser.find_element_by_css_selector("a[href='#attachmentsTab']")
    attachments_tab.click()
    
    time.sleep(2)
    
    #ctc_link = browser.find_element_by_css_selector('')
    
    browser.execute_script("""
        // GET THE OFFER LETTER LINK
        var trs = document.querySelectorAll('div#attachmentsGrid .ui-grid-row ');
        for(var i = 0; i < trs.length; i++){
            var tds = trs[i].querySelectorAll('.ui-grid-cell a.viewattachment');
            for(var j = 0; j < tds.length; j++){
                if(tds[j].innerText.includes('Offer Letter')){
                    var offer_link = trs[i].querySelector('.ui-grid-cell a.viewattachment');
                    var row_cells = trs[i].querySelectorAll('.ui-grid-cell-contents');
                    var view_folder_req = false;
                    for(var k = 0; k < row_cells.length; k++){
                        if(!view_folder_req){
                            var view_folder_link = row_cells[k].querySelector('a.viewfolder');
                            if(view_folder_link != undefined && view_folder_link != null){
                                view_folder_req = view_folder_link.innerText.includes(arguments[2]);
                            }
                        }
                    }
                    console.log(view_folder_link);
                    console.log(view_folder_req);
                    
                    if(view_folder_req 
                        && offer_link.innerText.includes('Offer Letter') 
                        && (offer_link.innerText.includes(arguments[0]) 
                        || offer_link.innerText.includes(arguments[1]))
                        && offer_link.innerText.includes('pdf')){
                            
                            console.log(trs[i]);
                            offer_link.click();
                            //j = tds.length;
                            i = trs.length;
                            //break;
                    }
                }
            }
        }
        
    """, row['Primary_First_Name'], row['Primary_Last_Name'], row['Job_Opening'])
                           
    time.sleep(12)                
    
    browser.switch_to.window(browser.window_handles[3])
    print(browser.title)
    
    time.sleep(12)
    
    print(browser.find_element_by_tag_name('html').get_attribute('innerHTML'))
    offer_pdf_embed_tag = browser.find_element_by_css_selector('embed[type="application/pdf"]')
    offer_pdf_href = offer_pdf_embed_tag.get_attribute('src')
    
    print(offer_pdf_href)

    url_current_page = browser.execute_script('''                
        return window.location.href;
    ''')
    print(url_current_page)
    
    
    url_current_page_2 = browser.execute_script('''                
        return document.querySelector('embed[type="application/pdf"]').getAttribute('src');
    ''')
    print(url_current_page_2)
    
    offer_pdf_href = url_current_page

    #options = webdriver.ChromeOptions()

    #profile = {"plugins.plugins_list": [{"enabled": False, "name": "Chrome PDF Viewer"}], # Disable Chrome's PDF Viewer
    #           "download.default_directory": DOWNLOADS_DIR , "download.extensions_to_open": "applications/pdf"}
    #options.add_experimental_option("prefs", profile)
    #driver_2 = webdriver.Chrome(CHROME_DRIVER_PATH, chrome_options=options)  # Optional argument, if not specified will search path.

    #driver_2.get(url_current_page_2)
    
    print("=============================Download OfferLetter Logic Starts===================================")

    browser.switch_to.window(browser.window_handles[1])
    print(browser.title)
    time.sleep(4)           
    offer_pdf_raw_url = "%r"%offer_pdf_href
    
    offer_download_script_str = download_str(offer_pdf_raw_url)

    browser.execute_script(offer_download_script_str)

    print("==========================================Download OfferLetter Logic Ends===================================================")
    

def offer_accept_download(browser, row):
    expand_activity_log = browser.find_element_by_css_selector('a#expandCollapseActLog')
    expand_activity_log.click()
    
    time.sleep(5)
    
    attachments_tab = browser.find_element_by_css_selector("a[href='#attachmentsTab']")
    attachments_tab.click()

    time.sleep(2)
    

    browser.execute_script("""
       /*
       var grid_headers = document.querySelectorAll('div[role="columnheader"]');
       for(var i=0; i < grid_headers.length; i++){
            console.log(grid_headers[i]);
            if(grid_headers[i].innerText.includes('Date')){
                var col_menu_down = grid_headers[i].querySelector('div[aria-label="Column Menu"] i.ui-grid-icon-angle-down');
                console.log(col_menu_down);
                col_menu_down.click();
                var sort_desc_btn = document.querySelector('i.ui-grid-icon-sort-alt-down');
                console.log(sort_desc_btn);
                sort_desc_btn.click();
            }
       }
       */
    """)
    
    time.sleep(0.5)
    
    browser.execute_script("""                
        /*
        var trs = document.querySelectorAll('div#actionlogGrid .ui-grid-row ');
        for(var i = 0; i < trs.length; i++){
            var tds = trs[i].querySelectorAll('.ui-grid-cell a.viewattachment');
            for(var j = 0; j < tds.length; j++){
                if(tds[j].innerText.includes('Offer Letter')){
                    var offer_link = trs[i].querySelector('.ui-grid-cell a.viewattachment');
                    var offer_accept_link = trs[i].querySelector('.ui-grid-cell a.viewdocsubform');
                    var row_cells = trs[i].querySelectorAll('.ui-grid-cell-contents');
                    var view_folder_req = false;
                    for(var k = 0; k < row_cells.length; k++){
                        if(!view_folder_req){
                            var view_folder_link = row_cells[k].querySelector('a.viewfolder');
                            if(view_folder_link != undefined && view_folder_link != null){
                                view_folder_req = view_folder_link.innerText.includes(arguments[2]);
                            }
                        }
                    }
                    console.log(view_folder_link);
                    console.log(view_folder_req);
                    
                    if(view_folder_req && offer_link != null && offer_accept_link != null
                        && offer_link.innerText.includes('Offer Letter') 
                        && (offer_link.innerText.includes(arguments[0]) 
                        || offer_link.innerText.includes(arguments[1]))
                        && offer_link.innerText.includes('pdf')
                        && offer_accept_link.innerText.includes('View')){
                            
                            //console.log(offer_accept_link.innerText);
                            //alert(offer_accept_link.innerText);
                            offer_accept_link.click();
                            //j = tds.length;
                            i = trs.length;
                            //break;
                    }
                }
            }
        }
        */
    """, row['Primary_First_Name'], row['Primary_Last_Name'], row['Job_Opening'])
    
    time.sleep(0.5)
    
    #attachments_tab = browser.find_element_by_css_selector("a[href='#attachmentsTab']")
    #attachments_tab.click()
    
    time.sleep(2)
    
    #ctc_link = browser.find_element_by_css_selector('')
    
    browser.execute_script("""
        // GET THE OFFER ACCEPTANCE LINK
        var trs = document.querySelectorAll('div#attachmentsGrid .ui-grid-row ');
        for(var i = 0; i < trs.length; i++){
            var tds = trs[i].querySelectorAll('.ui-grid-cell a.viewattachment');
            for(var j = 0; j < tds.length; j++){
                if(tds[j].innerText.includes('Offer Letter')){
                    var offer_link = trs[i].querySelector('.ui-grid-cell a.viewattachment');
                    var offer_accept_link = trs[i].querySelector('.ui-grid-cell a.viewdocsubform');
                    var row_cells = trs[i].querySelectorAll('.ui-grid-cell-contents');
                    var view_folder_req = false;
                    for(var k = 0; k < row_cells.length; k++){
                        if(!view_folder_req){
                            var view_folder_link = row_cells[k].querySelector('a.viewfolder');
                            if(view_folder_link != undefined && view_folder_link != null){
                                view_folder_req = view_folder_link.innerText.includes(arguments[2]);
                            }
                        }
                    }
                    console.log(view_folder_link);
                    console.log(view_folder_req);
                    
                    if(view_folder_req && offer_link != null && offer_accept_link != null
                        && offer_link.innerText.includes('Offer Letter') 
                        && (offer_link.innerText.includes(arguments[0]) 
                        || offer_link.innerText.includes(arguments[1]))
                        && offer_link.innerText.includes('pdf')
                        && offer_accept_link.innerText.includes('View')){
                            
                            //console.log(offer_accept_link.innerText);
                            //alert(offer_accept_link.innerText);
                            offer_accept_link.click();
                            //j = tds.length;
                            i = trs.length;
                            //break;
                    }
                }
            }
        }
        
    """, row['Primary_First_Name'], row['Primary_Last_Name'], row['Job_Opening'])
                           
    time.sleep(12)                
    
    offer_letter_window = browser.window_handles[3]
    browser.switch_to.window(offer_letter_window)
    print(browser.title)
    
    time.sleep(2)
    
    view_pdf_button_elem = browser.find_element_by_id('btnViewPDFForDocSubForm')
    print(browser.window_handles)
    
    #view_pdf_button_elem.click()
    
    #print_pdf_button_elm = browser.find_element_by_id('btnPrint')
    #print_pdf_button_elm.click()
    
    
    #browser.switch_to.window(browser.window_handles[4])
    #print(browser.title)
    #time.sleep(20) 
                              

    accept_pdf_href = browser.execute_script("""              

            return window.location.href;
            //return document.documentElement.innerHTML;

            //setInterval(function(){return document.querySelector('div#content').getAttribute('outerHTML');}, 3000);
    """)
    print(str(accept_pdf_href))
    
    
    time.sleep(20)
    
    print("=============================Download AcceptanceLetter Logic Starts===================================")

    browser.switch_to.window(browser.window_handles[1])
    print(browser.title)            
    time.sleep(1)           
    accept_pdf_raw_url = "%r"%accept_pdf_href
    
    accept_download_script_str = download_str(accept_pdf_raw_url)
    
    browser.execute_script(accept_download_script_str)
    
    
    #browser.execute_script('''
    #    //document.querySelector('body').hover();
    #    //document.querySelector('#toolbar').shadowRoot.querySelector('#download').click();

    #    frag = document.querySelector("print-preview-app").shadowRoot;
    #    sidebar = frag.querySelector("#sidebar").shadowRoot;
    #    dest = sidebar.querySelector("#destinationSettings").shadowRoot;
    #    preview_set = dest.querySelector("print-preview-settings-section");
    #    dest_select = preview_set.querySelector("print-preview-destination-select").shadowRoot;
    #    //dest_select.querySelector("select").value;


    #    dest_select.querySelector("select").value = "Save as PDF/local/";
    #''')
    

    print("==========================================Download AcceptanceLetter Logic Ends===================================================")
    

def ctc_download(browser, row):

    expand_activity_log = browser.find_element_by_css_selector('a#expandCollapseActLog')
    expand_activity_log.click()
    
    time.sleep(5)
    
    attachments_tab = browser.find_element_by_css_selector("a[href='#attachmentsTab']")
    attachments_tab.click()

    time.sleep(2)
    

    #attachments_tab = browser.find_element_by_css_selector("a[href='#attachmentsTab']")
    #attachments_tab.click()
    
    time.sleep(2)
    
    #ctc_link = browser.find_element_by_css_selector('')
    
    browser.execute_script("""
        /*
        var attachment_links = document.querySelectorAll('a.viewattachment');
        for(var i = 0; i < attachment_links.length; i++){
            var a_txt = attachment_links[i].innerText;
            if(
                (a_txt.toLowerCase().includes('ctc') || a_txt.toLowerCase().includes('breakup') || a_txt.toLowerCase().includes('grid')) && 
                (a_txt.includes('pdf') || a_txt.includes('PDF'))){
                    attachment_links[i].click();
                    i = attachment_links.length;
            }
        }
        */
        
        // GET THE CTC LETTER LINK
        var trs = document.querySelectorAll('div#attachmentsGrid .ui-grid-row ');
        for(var i = 0; i < trs.length; i++){
            var tds = trs[i].querySelectorAll('.ui-grid-cell a.viewattachment');
            for(var j = 0; j < tds.length; j++){
                if((
                    tds[j].innerText.toLowerCase().includes('ctc') 
                    || tds[j].innerText.toLowerCase().includes('breakup') 
                    || tds[j].innerText.toLowerCase().includes('grid')
                    || tds[j].innerText.toLowerCase().includes(arguments[0].toLowerCase())
                    || tds[j].innerText.toLowerCase().includes(arguments[2].toLowerCase())
                    ) && !tds[j].innerText.toLowerCase().includes('Offer Letter:')){
                    var ctc_link = trs[i].querySelector('.ui-grid-cell a.viewattachment');
                    var row_cells = trs[i].querySelectorAll('.ui-grid-cell-contents');
                    var view_folder_req = false;
                    for(var k = 0; k < row_cells.length; k++){
                        if(!view_folder_req){
                            var view_folder_link = row_cells[k].querySelector('a.viewfolder');
                            if(view_folder_link != undefined && view_folder_link != null){
                                view_folder_req = view_folder_link.innerText.includes(arguments[1]);
                            }
                        }
                    }
                    console.log(view_folder_link);
                    console.log(view_folder_req);
                    var a_txt = ctc_link.innerText;                            
                    if(view_folder_req 
                        && (a_txt.toLowerCase().includes('ctc') || a_txt.toLowerCase().includes('breakup') 
                            || a_txt.toLowerCase().includes('grid') || a_txt.toLowerCase().includes(arguments[0].toLowerCase()) 
                            || a_txt.toLowerCase().includes(arguments[2].toLowerCase())) 
                        && (a_txt.includes('pdf') || a_txt.includes('PDF'))
                        && !a_txt.includes('Offer Letter:')){
                            
                            //console.log(offer_accept_link.innerText);
                            //alert(offer_accept_link.innerText);
                            ctc_link.click();
                            //j = tds.length;
                            i = trs.length;
                            //break;
                    }
                }
            }
        }
        
    """, row['Primary_First_Name'], row['Job_Opening'], row['Primary_Last_Name'])
                           
    time.sleep(12)

    browser.switch_to.window(browser.window_handles[3])
    print(browser.title)            

    ctc_pdf_href = browser.execute_script("""              
            //return document.querySelector('embed[type="application/pdf"]').getAttribute('src');
            return window.location.href;
    """)
    
    print("=============================Download CTCBreakupLetter Logic Starts===================================")

    browser.switch_to.window(browser.window_handles[1])
    print(browser.title)            
    time.sleep(1)           
    ctc_pdf_raw_url = "%r"%ctc_pdf_href
    
    ctc_download_script_str = download_str(ctc_pdf_raw_url)

    browser.execute_script(ctc_download_script_str)

    print("==========================================Download CTCBreakupLetter Logic Ends===================================================")


def rename_newest_download(new_file_name, copy_to_dir='resumes_dir'):
    now = dt.datetime.now()
    ago = now-dt.timedelta(minutes=1)

    for root, dirs, files in os.walk(DOWNLOADS_DIR):  
        for fname in files:
            path = os.path.join(root, fname)
            st = os.stat(path)    
            mtime = dt.datetime.fromtimestamp(st.st_mtime)
            if mtime > ago and 'UserFile' in fname and (fname.split('.')[-1] == 'pdf' or fname.split('.')[-1] == 'html'): 
                print('FILEPATH >> %s ================ MODIFIED TIME >> %s'%(path, mtime))
                latest_downloaded_file = path
                #os.rename(latest_downloaded_file, resume_name)
                move(latest_downloaded_file, new_file_name)
                if copy_to_dir == 'resumes_dir':
                    # copy(new_file_name, RESUMES_DOWNLOADS_COPY_TO)
                    move(new_file_name, RESUMES_DOWNLOADS_COPY_TO)
                elif copy_to_dir == 'offers_dir':
                    # copy(new_file_name, OFFERS_DOWNLOADS_COPY_TO)
                    move(new_file_name, OFFERS_DOWNLOADS_COPY_TO)
                
    
@xw.sub
def download_all_docs():
    #browser = webdriver.Chrome(CHROME_DRIVER_PATH)
    
    # settings_download = {
        # "appState": {
            # "recentDestinations": [{
                # "id": "Save as PDF",
                # "origin": "local"
            # }],
            # "selectedDestinationId": "Save as PDF",
            # "version": 2
        # }  
    # }
    os.environ["webdriver.chrome.driver"] = CHROME_DRIVER_PATH
    chrome_options = Options()
    prefs = {
        'profile.default_content_setting_values.automatic_downloads': 1,
        'download.prompt_for_download': False,
        # 'printing.print_preview_sticky_settings': json.dumps(settings_download),
        # 'plugins.plugins_disabled': 'Chrome PDF Viewer',
        # 'plugins.always_open_externally': True
        }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument('start-maximized')

    browser = webdriver.Chrome(CHROME_DRIVER_PATH, options=chrome_options)

    # profs_dir = os.path.expanduser("~/AppData/Roaming/Mozilla/Firefox/Profiles")
    # prof_found = max([os.path.join(profs_dir,d) for d in os.listdir(profs_dir)], key=os.path.getmtime)
    # print(prof_found)
    
    # profile = webdriver.FirefoxProfile(prof_found)
    # profile.set_preference('browser.download.folderList', 2) # custom location

    # profile.set_preference('browser.download.manager.showWhenStarting', False)
    # profile.set_preference('browser.download.dir', DOWNLOADS_DIR)
    # profile.set_preference('browser.helperApps.neverAsk.saveToDisk', "application/pdf")
    # caps = DesiredCapabilities.FIREFOX.copy()
    # caps['marionette'] = False
    # binary = FirefoxBinary(r'C:\\Program Files (x86)\\geckodriver-v0.26.0-win64\\geckodriver.exe')
    # browser = webdriver.Firefox(firefox_binary=binary, caps)

    
    
    # capabilities = webdriver.DesiredCapabilities().FIREFOX
    # capabilities["marionette"] = True
    # binary = FirefoxBinary('C:/Program Files/Mozilla Firefox/firefox.exe')
    # browser = webdriver.Firefox(firefox_profile=profile, 
                                # firefox_binary=binary, 
                                # capabilities=capabilities, 
                                # executable_path=r'C:\\Program Files (x86)\\geckodriver-v0.26.0-win64\\geckodriver.exe')
    #driver.get("http://www.google.com")    
    
    
    browser.implicitly_wait(30) #wait 30 seconds when doing a find_element before carrying on
    #browser.delete_all_cookies()
    browser.get(HR_LINKS_PORTAL_URL)
    browser.maximize_window()
    #timeout = 30
    
    try:
        # browser.implicitly_wait(30) #wait 30 seconds when doing a find_element before carrying on

        #links_present = EC.presence_of_all_elements_located((By.CLASS_NAME, 'a65a'))
        #WebDriverWait(browser, timeout).until(links_present)
        
        #get the window handles using window_handles( ) method
        window_before = browser.window_handles[0]
        print(browser.title)
        
        kenexa_link = browser.find_element_by_link_text("Kenexa")
        kenexa_link.click()
                  
        time.sleep(5)
        
        browser.switch_to.window(browser.window_handles[1])
        #browser.implicitly_wait(30)
        print(browser.title)
        
        completed_df = pd.read_excel(INDIA_HIRING_COMPLETED_SHEET, "Completed")
        completed_df_useful = completed_df[['Associate ID', 'Applicant_ID', 'Job_Opening', 
                           'Primary_First_Name', 'Primary_Last_Name', 'Type_Hire', 'Currency']]
                           
        #print(completed_df_useful)
                           
        for row_index, row in completed_df_useful.iterrows():
            print(type(row['Job_Opening']))
            if isinstance(row['Job_Opening'], str):            
                print(row['Job_Opening'])
                time.sleep(5)
                
                search_link = browser.find_element_by_id('searchlink')
                search_link.click()

                search_reqs_link = browser.find_element_by_link_text('Reqs')
                search_reqs_link.click()
                
                search_req_text = browser.find_element_by_id('quicksearch')
                search_req_text.clear()
                search_req_text.send_keys(row['Job_Opening'])
                search_req_text.send_keys(Keys.ENTER)
                
                time.sleep(3)
                
                cand_count_link = browser.find_element_by_css_selector('a.candidateCount')
                cand_count_link.click()
                
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
                
                time.sleep(2)
                
                try:
                    hired_input = browser.find_element_by_xpath("//input[@id='hrstatus-13709']")            
                except NoSuchElementException:
                    try:
                        hired_input = browser.find_element_by_xpath("//input[@id='hrstatus-683513']")
                    except NoSuchElementException:
                        hired_input = browser.find_element_by_xpath("//input[@id='hrstatus-180171']")
                    
                browser.execute_script('''
                    var hiredInput = arguments[0];
                    hiredInput.click();
                ''', hired_input)
                
                time.sleep(10)
                
                cand_name_link = browser.find_element_by_css_selector('a.candname')
                cand_name_link.click()
                
                time.sleep(3)
                
                browser.switch_to.window(browser.window_handles[2])
                
                
                time.sleep(3)
                
                if 'internal' not in str(row['Type_Hire']).lower():
                    print("========================================== NOW START DOWNLOADING RESUME ===========================================")
                    resume_download(browser)                
                      
                    print("==========================================Rename Resume File Logic Starts===============================================")


                    time.sleep(20)
                    resume_name = DOWNLOADS_DIR+"\\0"+str(row['Associate ID'])+"_"+row['Primary_Last_Name'].capitalize()+","+row['Primary_First_Name'].capitalize()+"_Resume.pdf"
                    rename_newest_download(resume_name, 'resumes_dir')
     
                                
                    print("==========================================Rename Resume File Logic Ends===============================================")
                    
                    #print(browser.window_handles)
                    
                    time.sleep(2)
                    browser.switch_to.window(browser.window_handles[3])
                    print(browser.title)
                    browser.close()
                    browser.switch_to.window(browser.window_handles[2])
                    print(browser.title)

                                
                    print("==========================================Resume Download-Rename Completed===================================")

                time.sleep(2)
                print("========================================== NOW START DOWNLOADING CTC-LETTER ===========================================")
                offer_letter_download(browser, row)
          
                print("==========================================Rename OfferLetter File Logic Starts===============================================")
                
                time.sleep(20)
                offer_letter_name = DOWNLOADS_DIR+"\\"+str(row['Associate ID'])+"_"+row['Primary_Last_Name'].capitalize()+","+row['Primary_First_Name'].capitalize()+"_OfferLetter.pdf"
                rename_newest_download(offer_letter_name, 'offers_dir')                
                
                print("==========================================Rename OfferLetter File Logic Ends===============================================")
                
                print(browser.window_handles)
                
                time.sleep(2)
                browser.switch_to.window(browser.window_handles[3])
                print(browser.title)
                browser.close()
                browser.switch_to.window(browser.window_handles[2])
                print(browser.title)
                            
                print("==========================================OfferLetter Download-Rename Completed===================================")
                
                time.sleep(2)
                print("========================================== NOW START DOWNLOADING ACCEPTANCE-LETTER ===========================================")
                offer_accept_download(browser, row)
                
                print("==========================================Rename AcceptanceLetter File Logic Starts===============================================")
                
                time.sleep(20)
                accept_letter_name_html = DOWNLOADS_DIR+"\\"+str(row['Associate ID'])+"_"+row['Primary_Last_Name'].capitalize()+","+row['Primary_First_Name'].capitalize()+"_AcceptanceLetter.html"
                rename_newest_download(accept_letter_name_html, 'offers_dir')
                
                # SINCE, ACCEPTANCE WILL BE HTML, CONVERT TO PDF HERE

                config = pdfkit.configuration(wkhtmltopdf=WKHTML_PATH)
                accept_letter_name = OFFERS_DOWNLOADS_COPY_TO+"\\"+str(row['Associate ID'])+"_"+row['Primary_Last_Name'].capitalize()+","+row['Primary_First_Name'].capitalize()+"_AcceptanceLetter.pdf"
                css = 'acceptance_styles.css'
                
                pdfkit.from_file(accept_letter_name_html, accept_letter_name, configuration=config, css=css)
                                
                
                print("==========================================Rename AcceptanceLetter File Logic Ends===============================================")
                
                time.sleep(2)
                
                # browser.switch_to.window(browser.window_handles[4])
                # browser.close()
                browser.switch_to.window(browser.window_handles[3])
                browser.close()
                browser.switch_to.window(browser.window_handles[2])

                print("==========================================AcceptanceLetter Download-Rename Completed===================================")
                
                    
                time.sleep(2)
                if 'inr' in str(row['Currency']).lower():
                    print("========================================== NOW START DOWNLOADING CTC-LETTER ===========================================")
                    ctc_download(browser, row)
                    
                    print("==========================================Rename CTCBreakupLetter File Logic Starts===============================================")
                    
                    time.sleep(20)
                    ctc_letter_name = DOWNLOADS_DIR+"\\"+str(row['Associate ID'])+"_"+row['Primary_Last_Name'].capitalize()+","+row['Primary_First_Name'].capitalize()+"_CTCLetter.pdf"
                    rename_newest_download(ctc_letter_name, 'offers_dir')
                    
                    print("==========================================Rename CTCBreakupLetter File Logic Ends===============================================")
                    print("==========================================CTCBreakupLetter Download-Rename Completed===================================")
                            
                print("==========================================Merge Offer & Acceptance & CTC Logic Starts=================================================")
                
                offer_letter_name = OFFERS_DOWNLOADS_COPY_TO+"\\"+str(row['Associate ID'])+"_"+row['Primary_Last_Name'].capitalize()+","+row['Primary_First_Name'].capitalize()+"_OfferLetter.pdf"
                accept_letter_name = OFFERS_DOWNLOADS_COPY_TO+"\\"+str(row['Associate ID'])+"_"+row['Primary_Last_Name'].capitalize()+","+row['Primary_First_Name'].capitalize()+"_AcceptanceLetter.pdf"
                ctc_letter_name = OFFERS_DOWNLOADS_COPY_TO+"\\"+str(row['Associate ID'])+"_"+row['Primary_Last_Name'].capitalize()+","+row['Primary_First_Name'].capitalize()+"_CTCLetter.pdf"
                result_letter_name = OFFERS_DOWNLOADS_COPY_TO+"\\0"+str(row['Associate ID'])+"_"+row['Primary_Last_Name'].capitalize()+","+row['Primary_First_Name'].capitalize()+"_Signed_OL.pdf"
                pdfs = []
                if 'inr' in str(row['Currency']).lower():                
                    pdfs = [offer_letter_name, accept_letter_name, ctc_letter_name]
                else:
                    pdfs = [offer_letter_name, accept_letter_name]
                    
                merger = PdfFileMerger()

                for pdf in pdfs:
                    merger.append(pdf)

                merger.write(result_letter_name)
                merger.close()
                
                print("======================JUST DELETE THE OfferLetter & AcceptanceLetter & CTCBreakupLetter since finished mergeing============")
                
                if os.path.isfile(result_letter_name):
                    [os.remove(pdf_name) for pdf_name in pdfs]                                       
                    
                print("=================================DELETE COMPLETE===========================================")
                
                print("==========================================Merge Offer & Acceptance & CTC Logic Ends===================================================")
                
                for i in range(len(browser.window_handles)-1, 1, -1):  
                    browser.switch_to.window(browser.window_handles[i])
                    browser.close()
                    
                browser.switch_to.window(browser.window_handles[1])

                time.sleep(5)
            
            
        time.sleep(5)
        
        browser.switch_to.window(browser.window_handles[0])
        browser.close()
        
    except TimeoutException:
        print("Timed out waiting for page to load")    
        
                        
                        
if __name__ == "__main__":
    download_all_docs()
