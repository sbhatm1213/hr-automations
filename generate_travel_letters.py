import webbrowser
import time
import xlwings as xw
import pandas as pd
from mailmerge import MailMerge
from datetime import date


VISA_AUTOMATE_PATH = r"\\Cernfs05\functional\HR\HRSC\Visa\automate-visa"
#VISA_AUTOMATE_PATH = r"Z:\\Departure"
#VISA_AUTOMATE_PATH = r"Z:\HR\HRSC\Visa\automate-visa"
ENTRY_TEMPLATES_DIR = VISA_AUTOMATE_PATH + r"\Entry"
INTERVIEW_TEMPLATES_DIR = VISA_AUTOMATE_PATH + r"\Interview"
EXCEL_SHEET_TO_READ = VISA_AUTOMATE_PATH + r"\travel-employee-data.xlsm"
FOLDER_TO_ADD_LETTERS = VISA_AUTOMATE_PATH + r"\test-letters"


@xw.sub
def generate_travel_letters():
    """Generates travel letters in docx format for all rows"""

    df = pd.read_excel(EXCEL_SHEET_TO_READ)
    df.dropna(thresh=10, inplace=True) # drop if atleast 10 columns have non-NaN values.
    #df.dropna() # drop if 10 columns are not NaN

    #selected_df = cell_range.options(pd.DataFrame).value
    #row_1_selected = list(selected_df)
    #rows_all_selected = [list(selected_row) for selected_i, selected_row in selected_df.iterrows()]
    #rows_all_selected.insert(0, row_1_selected)
    #print(rows_all_selected)

    #selected_df_final = pd.DataFrame(rows_all_selected, columns=list(df)[1:])
    #print(selected_df_final)
    #selected_df_final = df[['FIRST NAME', 'MIDDLE NAME', 'LAST NAME', 'GENDER', 'TRAVEL TO', 'LETTER TYPE', 'TRAVEL FROM DATE', 'TRAVEL TO DATE', 'JOB TITLE']]
    
    selected_df_final = df
    print(selected_df_final)
    empty_df = pd.DataFrame()

    #try:
    for row_index, row in selected_df_final.iterrows():
        #print(df.iloc[row_index])
        mr_or_ms, her_or_his_small, his_or_her_caps, him_or_her_small = '', '', '', ''
        if 'Female'.lower() in row['GENDER'].lower():
            mr_or_ms = 'Ms.'
            her_or_his_small = 'her'
            him_or_her_small = 'her'
            his_or_her_caps = 'Her'
        elif 'Male'.lower() in row['GENDER'].lower():
            mr_or_ms = 'Mr.'
            her_or_his_small = 'his'
            him_or_her_small = 'him'
            his_or_her_caps = 'His'
        if 'Entry'.lower() in row['LETTER TYPE'].lower():
            if 'KC'.lower() in row['TRAVEL TO'].lower():
                document = MailMerge(ENTRY_TEMPLATES_DIR+r'\Kansas City - Entry Letter.docx')
            elif 'Malvern'.lower() in row['TRAVEL TO'].lower():
                document = MailMerge(ENTRY_TEMPLATES_DIR+r'\Malvern - Entry Letter.docx')
            if document:
                print(her_or_his_small)
                document.merge(
                    #todays_date='{:%B %d, %Y}'.format(date.today()),
                    mr_or_ms=mr_or_ms,
                    his_or_her_caps=his_or_her_caps,
                    small_his_or_her=her_or_his_small,
                    first_name=str(row['FIRST NAME']),
                    #middle_name=str(row['MIDDLE NAME']),
                    last_name=str(row['LAST NAME']),
                    job_title=str(row['JOB TITLE']),
                    from_date='{:%B %d, %Y}'.format(row['TRAVEL FROM DATE']),
                    to_date='{:%B %d, %Y}'.format(row['TRAVEL TO DATE']),
                    stay_address=str(row['STAY ADDRESS'])
                )
                # have all these fields as MergeFields in ENTRY TEMPLATE 
                
                document.write(FOLDER_TO_ADD_LETTERS + '\{}.docx'.format(str(row['OPERATOR ID'])))

        elif 'Interview'.lower() in row['LETTER TYPE'].lower():
            if 'KC'.lower() in row['TRAVEL TO'].lower():
                document = MailMerge(INTERVIEW_TEMPLATES_DIR+r'\KC- Invitation Letter Chennai Consulate.docx')
            elif 'Malvern'.lower() in row['TRAVEL TO'].lower():
                document = MailMerge(INTERVIEW_TEMPLATES_DIR+r'\Malvern - Invitation Letter Chennai Consulate.docx')
            if document:
                print(her_or_his_small)
                document.merge(
                    #todays_date='{:%B %d, %Y}'.format(date.today()),
                    mr_or_ms=mr_or_ms,
                    his_or_her_caps=his_or_her_caps,
                    small_his_or_her=her_or_his_small,
                    him_or_her_small=him_or_her_small,
                    first_name=str(row['FIRST NAME']),
                    #middle_name=str(row['MIDDLE NAME']),
                    last_name=str(row['LAST NAME']),
                    job_title=str(row['JOB TITLE']),
                    job_location=str(row['JOB LOCATION']),
                    employed_from_date='{:%B %d, %Y}'.format(row['EMPLOYED FROM DATE']),
                    from_date='{:%B %d, %Y}'.format(row['TRAVEL FROM DATE']),
                    to_date='{:%B %d, %Y}'.format(row['TRAVEL TO DATE']),
                    stay_address=str(row['STAY ADDRESS']),
                    passport_number=str(row['PASSPORT NUMBER'])

                )

            # have all these fields as MergeFields in FTE_TEMPLATE

                #document.write(FOLDER_TO_ADD_LETTERS + '\{}.docx'.format(row['FIRST NAME']+' '+row['MIDDLE NAME']+' '+row['LAST NAME']+'-'+row['LETTER TYPE']+'-'+row['TRAVEL TO']))
                document.write(FOLDER_TO_ADD_LETTERS + '\{}.docx'.format(str(row['OPERATOR ID'])))
    #empty_df.to_excel(EXCEL_SHEET_TO_READ, 'Sheet1', columns=list(df), index=False)
    #except Exception as e:
    #    return "There was some problem while performing the operation."+str(e)
    return "Generated travel letters successfully !"
    
    
if __name__ == '__main__':
    generate_travel_letters()
