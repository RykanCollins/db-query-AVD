#! python3.11
#Query app to review/edit data in dbo.Laserfiche_emails database
#programmer: Rykan Collins @MesaColdStorage
#Date: 04/05/2023

import PySimpleGUI as sg
import pyodbc, sys, re, pyperclip, pyinputplus as pyip

sg.theme('LightBlue1')
def main():
    #connecting to database
    conn = pyodbc.connect ("Driver=ODBC Driver 17 for SQL Server;"
                "Server=mcs16-sql;"
                "Database=dc00mes;"
                "UID=DC00MES;"
                "PWD=DC00MES;")

    cursor = conn.cursor()



    font = ('Aial',14)
    layout = [[sg.Text('Laserfiche Email Application', size=(30,1), font=font)],
              #[sg.Text('Client number: '), sg.Text(size=(15,1), key='-OUTPUT-')],
              #[sg.Text('DocType: '),sg.Text(size=(15,1), key='-DOCOUT-')],
              #[sg.Text('Client name: '), sg.Text(size=(15,1), key='-NAMEOUT-')],

              [sg.Text('*Client Number: '),sg.InputText(key=('cl_number'))],
              [sg.Text('Client Name: '), sg.InputText(key='-INCLNAME-')],
              [sg.Text('Email (sperate with ;): '), sg.InputText(key='cl_email', size=(30,4))],
              [sg.Text('*DocType: '), sg.OptionMenu(values=('INV','BOL','REC'), size=(20 ,1), key='-DOCTYPE-')], #Choice of doctypes in db
              [sg.Text('Email invoice pdf: Y/N'), sg.InputText(key='email_add')],
              [sg.Button('View'), sg.Button('Add'), sg.Button('Delete'), sg.Button('Quit')]
              
              ]

    window = sg.Window('Look Ma', layout)

    while True: #Event Loop
        event, values = window.read()
        print(event, values)
        if event in (sg.WIN_CLOSED, 'Quit'):
            sg.popup('!!Double Middle Finger Emoj!!')
            break
        
        elif event == 'View':
            #doc == '-DOCTYPE-', clNum == int('-IN-')
            clNum = int(values['cl_number'])
            docty = values['-DOCTYPE-']
            viewStmt = cursor.execute('''SELECT client_nbr, client_name, email, doctype, emailinv
                            FROM dbo.laserfiche_emails where doctype = ? and client_nbr = ?''', docty, clNum)
                                    
            text = str(viewStmt.fetchall())
            #sg.popup(text)

            pyperclip.copy(text)

            # Creates email address regex.
            emailRegex = re.compile(r'''(
                [a-zA-Z0-9._%+-]+      # username
                @                      # @ symbol
                [a-zA-Z0-9.-]+         # domain name
                (\.[a-zA-Z]{2,4})       # dot-something
                )''', re.VERBOSE)

            matches = []

            for groups in emailRegex.findall(text):
                matches.append(groups[0])
                                    
            # Copy results to the clipboard and display results on screen
            if len(matches) > 0:
                pyperclip.copy('\n'.join(matches))
                #print('Copied to clipboard: ')
                sg.Print('\n'.join(matches),'\n')
                
            else:
                sg.popup_error('No records found.')

        elif event == 'Add':
            
            acctNbr = int(values['cl_number'])
            acctName = values['-INCLNAME-']
            acctEmail = values['cl_email']
            acctDocTy = values['-DOCTYPE-']
            emailAdd = values['email_add']
            insert_stmt = ("""Insert into dbo.Laserfiche_emails (client_nbr, client_name, email, doctype, emailinv)\
                            values (?, ?, ?, ?, ?)""")
            data = (acctNbr, acctName, acctEmail, acctDocTy, emailAdd)
            try:
               # Executing the SQL command
               
               cursor.execute(insert_stmt,data)
           
               # Commit your changes in the database
               conn.commit()
               sg.Print("Data entered: ", data)
               

            except OSError as err:
               # Rolling back in case of error
               conn.rollback()
               
               sg.popup_error("Data error:",err)

        elif event == 'Delete':
            acctNbr = values['cl_number']
            acctDocTy = values['-DOCTYPE-']
            
            delete_stmt = ("""Delete from dbo.Laserfiche_emails where client_nbr = ? and doctype = ?""")
            data = (acctNbr, acctDocTy)
            
            try:
               # Executing the SQL command
               cursor.execute(delete_stmt, data)
           
               # Commit your changes in the database
               conn.commit()
               sg.Print("Data delete for: ", data )

            except OSError as err:
               # Rolling back in case of error
               conn.rollback()
               sg.Print("Data error: ", err)
                
    conn.close()   #close the db connection
    window.close()

if __name__ == '__main__':

    main()


