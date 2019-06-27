import win32com.client
import sys
import re
import pandas
from openpyxl import load_workbook
import datetime

class EmailReader(object):

    def __init__(self):
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        try:
            root_folder = self.outlook.Folders.Item(1)
            # Get all mails in Alerts folder
            self.messages = root_folder.Folders['Alerts'].items
        except Exception as e:
            print("Unable to open email " + str(e))
            sys.exit(-1)

    # clean up
    def close(self):
        del self.outlook

    # Return all messages
    def getMessages(self):
        return self.messages

    def processEmailRequest(self,datefrm,dateto):

        df = pandas.DataFrame()
        DATE_FORMAT = "%Y-%m-%d  %H:%M"
        from_date = datetime.datetime.strptime(datefrm, DATE_FORMAT)
        to_date = datetime.datetime.strptime(dateto, DATE_FORMAT)

        for mailItem in self.getMessages():
            recv = str(mailItem.ReceivedTime)
            new_dt = recv[:19]
            mailRecvTime = datetime.datetime.strptime(new_dt,'%Y-%m-%d %H:%M:%S')
            if mailRecvTime > from_date and mailRecvTime < to_date:
                subj = mailItem.subject.upper()
                lineSplit = subj.split(" ")
                for x in lineSplit[:]:
                    data = re.findall(r'MTS(.*)', x, re.M | re.I)
                    if data:
                        data = 'MTS' + ''.join(data)
                        data1 = pandas.DataFrame({"Case#":[data],"Date": [mailRecvTime]})
                        df = df.append(data1)

        if not(df.empty):
            try:
                book = load_workbook('C:\projects\python3.6\python36\log_panda.xlsx')
                writer = pandas.ExcelWriter('C:\projects\python3.6\python36\log_panda.xlsx', engine='openpyxl')
                writer.book = book
                df.to_excel(writer,"Main",index=False)
                writer.save()
                writer.close()
            except PermissionError:
                print("Could not open file for writing data! Please close Excel!")


if __name__ == "__main__":
    processor = EmailReader()
    try:
        inp1 = input("Enter from date for searching alerts: ")
        inp2 = input("Enter to date for seaching alerts: ")
        processor.processEmailRequest(inp1, inp2)
        processor.close()
    except ValueError as value:
        print("The error is {}. Please provide correct data format".format(value))


