from selenium.webdriver.support.ui import Select
import time
import openpyxl

def login(browser, username, password):
    institutionField = browser.find_element_by_id('txtClientName')
    institutionField.send_keys('Icahn School of Medicine at Mount Sinai - ISMMS')
    usernameField = browser.find_element_by_id('txtUsername')
    usernameField.send_keys(username)
    passwordField = browser.find_element_by_id('txtPassword')
    passwordField.send_keys(password)
    browser.find_element_by_id('btnLoginNew').click()
    return

def log_procedure(browser, data):
    browser.get('https://www.new-innov.com/Logger/LOGGER_Host.aspx?Control=PLogView')
    browser.find_element_by_id('ctl04_lbnAddLog').click()

    browser.find_element_by_id('txtID').send_keys(data[0])
    # print(data)

    gender=Select(browser.find_element_by_id('ctl04_PatientFieldsContainer1_ddlGender'))
    if data[1].lower()[0] == 'm':
        gender.select_by_index(1)
    elif data[1].lower()[0] == 'f':
        gender.select_by_index(2)
    else:
        gender.select_by_index(0)

    browser.find_element_by_id('ctl04_ProcedureDiagnosisCustomFieldsContainerManager1_ctl00_ProcedureFieldsContainer1_txtDatePerformed_dateInput').send_keys(data[2])

    location=Select(browser.find_element_by_id('ctl04_ProcedureDiagnosisCustomFieldsContainerManager1_ctl00_ProcedureFieldsContainer1_ddlLocations'))
    locationOptions = location.options
    for i in range(len(locationOptions)):
        if data[3].lower() == locationOptions[i].text.lower():
            location.select_by_index(i)
            break

    supervisor = Select(browser.find_element_by_id(
        'ctl04_ProcedureDiagnosisCustomFieldsContainerManager1_ctl00_ProcedureFieldsContainer1_ddlSupervisors'))
    supervisorOptions = supervisor.options
    for i in range(len(supervisorOptions)):
        if data[5].lower() == supervisorOptions[i].text.lower():
            # print(supervisorOptions[i].text.lower())
            supervisor.select_by_index(i)
            break

    role = Select(browser.find_element_by_id(
        'ctl04_ProcedureDiagnosisCustomFieldsContainerManager1_ctl00_ProcedureFieldsContainer1_ddlRoles'))
    roleOptions = role.options
    for i in range(len(roleOptions)):
        if data[6].lower() == roleOptions[i].text.lower():
            role.select_by_index(i)
            break

    diagnosis_test = browser.find_element_by_id(
        'ctl04_ProcedureDiagnosisCustomFieldsContainerManager1_ctl00_DiagnosisFieldsContainerManager1_ctl00_txtDiagnosisText')
    diagnosis_test.send_keys(data[7])

    procedure = browser.find_element_by_id('ctl04_ProcedureDiagnosisCustomFieldsContainerManager1_ctl00_ProcedureFieldsContainer1_ddlProcedures_Input')
    # procedureOptions = procedure.options
    # for i in range(len(procedureOptions)):
    #     if data[4].lower() == procedureOptions[i].text.lower():
    #         location.select_by_index(i)
    procedure.clear()
    procedure.send_keys(data[4].lower())
    browser.find_element_by_id('txtComplication').click()

    browser.find_element_by_id('ctl04_btnSave').click()
    time.sleep(2)

    return

def get_log_data(excel_file):
    wb = openpyxl.load_workbook(excel_file)
    sheet = wb['Log']
    data = []
    for rows in sheet.rows:
        temp = []
        for columns in rows:
            temp.append(str(columns.value))
        data.append(temp)
    wb.close()
    return data