from time import strftime, gmtime
import os.path
from twilio.rest import Client
import xlrd
import datetime
import openpyxl
import time


while True:         
    wb = xlrd.open_workbook(os.path.join('D:/Test case','Book2.xlsx'))
    book = openpyxl.load_workbook('D:/Test case/Book2.xlsx') #opening openpyxl workbook
    sheet = book.active
    wb.sheet_names()
    sh = wb.sheet_by_index(0)
    for i in range(1,(sh.nrows)):
        D1 =(str((sh.cell_value(i,3))))
        D2 =str(sh.cell_value(i,9))
        
        D2 = int(float(D2) * 86400)
        yn = str(sh.cell_value(i,13))
        b=str(datetime.datetime.now().time())[:8]
        
        a=str(datetime.datetime.now().date())
        loc=str(sh.cell_value(i,7))
        wi=str(sh.cell_value(i,2))
        def timetonum(time_str):
            hh,mm,ss=map(int,time_str.split(":"))
            return ss+60*(mm+60*hh)
       
            
        b=int(timetonum(b))
       
        c=b

        if D1==a and yn=="yes" and abs(D2-b)<=120  :
            
            account_sid = "ACd4f6641a1e7853611bd50ad0fa8973c0"
            auth_token = "85f1c41d7f3c84ce4df6a26e277b117b"
            client = Client(account_sid, auth_token)

            msgintro="You have an appointment with "+wi+" at "+loc+" at time "+str(b)
            msgout="You have an appointmnent at "+loc+" at time "+str(b)
            msg1 = client.messages.create(
                to="+918239644770",
                from_="+16264653123",
                    
                body=msgintro
            )
            #sheet['U'+str(i+1)] = 1  #writing to cell 1,20
            #book.save("D:/Test case/Book2.xlsx") #Saving the xlsx
            
            print(msg1.sid)
            msg2 = client.messages.create(
                to="+918239644770",
                from_="+16264653123",
                    
                body=msgout
                )

        
            print(msg2.sid)
    z=str(datetime.datetime.now().time())[:8]
    z=int(timetonum(z))
    y=int(z)-int(c)
    y=int(y)+120
    time.sleep((y))
