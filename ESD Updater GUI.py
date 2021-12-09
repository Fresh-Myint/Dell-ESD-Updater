import pyodbc
import os
import pandas as pd
import PySimpleGUI as sg
from PySimpleGUI.PySimpleGUI import BUTTON_TYPE_BROWSE_FILE

cevaLogo = r'C:\Users\myintch\Desktop\ESD_UPDATE_PY\CEVA_LOGO.ico'
dataFrame = ''
ODBC_Conn = ()


def unloadUserData():  # Passed
    window['-USER-'].update('')
    window['-PASSWORD-'].update('')


def ODBCConnection(user, password):
    try:
        connection = pyodbc.connect(f"Driver=iSeries Access ODBC Driver;System=PKMS.US.LOGISTICS.CORP;Uid={user};Pwd={password};")
    except pyodbc.Error as err:
        if err.args[0] == '28000':
            sg.popup('General ODBC Error:\nPlease verify login info.', icon=cevaLogo)
            return  # nothing
        else:
            sg.popup('General ODBC Error:\nTry again or contact admin if problem persists.', icon=cevaLogo)
            return  # nothing

    cursor = connection.cursor()
    print('PKMS Login Succesful!')
    return connection, cursor


def getDataFromExcel(excelFile):
    fileName = os.path.basename(excelFile)

    if fileName != 'ESD_Modifier.xlsx':
        sg.popup('Invalid file selected. \n'
                 'Please select a new file and try again.',
                 title='Invalid File',
                 icon=cevaLogo)
        return  # nothing
    else:
        df = pd.read_excel(excelFile, engine='openpyxl')

    return df


def executeSQLUpdate(connection, cursor, df):
    numberOfRows = len(df.index)
    errorControlNum = []
    msgBox = sg.popup_yes_no(f'You are about to update {numberOfRows} rows.\n'
                             'Do you wish to continue?', title='Confirm', icon=cevaLogo)

    if msgBox == 'Yes':
        for row in df.itertuples():
            controlNum = str(row[1])
            expectedShipDte = str(row[2])

            if len(controlNum) == 10:
                #cursor.execute(f"UPDATE CAPM01.WM0272PRDD.PHPICK00 SET PHCMDT = '{expectedShipDte}' WHERE PHPCTL = '{controlNum}'")
                print(f"UPDATE CAPM01.WM0272PRDD.PHPICK00 SET PHCMDT = '{expectedShipDte}' WHERE PHPCTL = '{controlNum}'")
                #connection.commit()
            else:
                errorControlNum.append(controlNum)

        if errorControlNum:
            print(f'The below order(s) were invalid and did not get update:\n{errorControlNum}')
        else:
            print('No errors found within records.')

        numebroferrors = numberOfRows - len(errorControlNum)
        print(f'{numebroferrors} of {numberOfRows} orders were updated in PKMS!')
    else:
        sg.popup('SQL update canceled by user.')


sg.theme('Default1')

layout = [
    [sg.Text('PKMS Username:', pad=5), sg.InputText(size=20, pad=5, key='-USER-'), sg.Button('Connect', size=(8, 1), key='-LOGIN-')],
    [sg.Text('PKMS Password:', pad=5), sg.InputText(size=20, pad=5, password_char='*', key='-PASSWORD-'),
     sg.Checkbox('Show', enable_events=True, key='-SHOWPWD-')],
   #[sg.Button('Cancel SQL Update', key='-CANCELSQL-')],
    [sg.FileBrowse(key='-FILE-'), sg.Button('Run SQL Update', key='-RUN-'), sg.Exit()],
    [sg.Output(size=(122,30),key='-OPUTBOX-')]
]

window = sg.Window('CEVA Logistics - ESD Updater', layout, size=(900, 500), icon=cevaLogo, resizable=True)

while True:
    event, values = window.read()
    # closes GUI
    if event in (sg.WIN_CLOSED, 'Exit'):  # Passed
        break
    # hide or show user password
    if event == '-SHOWPWD-':  # Passed
        if values['-SHOWPWD-']:
            window['-PASSWORD-'].update(password_char='')  # Passed
        else:
            window['-PASSWORD-'].update(password_char='*')  # Passed
    # Connects to SQL
    if event == '-LOGIN-':  # Passed
        if values['-USER-'] != '' and values['-PASSWORD-'] != '':
            ODBC_Conn = ODBCConnection(values['-USER-'], values['-PASSWORD-'])
            unloadUserData()  # passed
        else:
            sg.popup('Username or Password cannot be blank.', title='Invalid input', icon=cevaLogo)  # passed
            unloadUserData()  # passed
    # Runs SQL update after getting file info and PkMS login
    if event == '-RUN-':
        if os.path.basename(values['-FILE-']) == 'ESD_Modifier.xlsx':
            if ODBC_Conn:  # Variable not initiated.
                dataFrame = getDataFromExcel(values['-FILE-'])

                executeSQLUpdate(ODBC_Conn[0], ODBC_Conn[1], dataFrame)
                window['-FILE-'].update('')  # unloads file - so user does not inadvertently update again.
            else:
                print('User not connected to PKMS.\n'
                      'Please enter login info above then click connect.')
        else:
            sg.popup('Invalid file selected.\n'
                     'Please select a new file and try again.', title='Invalid File', icon=cevaLogo)


window.close()
