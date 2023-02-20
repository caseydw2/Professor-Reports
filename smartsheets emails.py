import pandas as pd
from GUI.GUI_Utilities import get_file
import datetime as dt

def record_emails(sheet,smartsheet_records):
    pass

def fix_ID(id:str):
    id,_ = id.split(".")
    id = id.zfill(7)
    return id

def since_monday(df,date_column = "Date"):
    weekday_today = dt.datetime.today().weekday()
    this_monday = dt.datetime.today() - dt.timedelta(days=weekday_today)
    this_monday = this_monday.replace(hour=0,minute=0,second=0,microsecond=0)
    emails_to_report = df.loc[df[date_column] > this_monday]
    return emails_to_report


def clean_smartsheet_records(filepath = get_file("Writing Studio Smartsheet") ) -> pd.DataFrame:
    smartsheet_record = pd.read_excel(filepath)
    emails = smartsheet_record.loc[smartsheet_record["Type of Session"] == "Email"]
    emails_to_report = emails.loc[emails["Report Visit?"] == "Yes"]
    emails_to_report = since_monday(emails_to_report)
    emails_to_report = emails_to_report.astype({"Student's ID":"str"})
    emails_to_report["Student's ID"] = emails_to_report["Student's ID"].apply(lambda x: fix_ID(x))
    return emails_to_report

if __name__ == "__main__":
    x = clean_smartsheet_records()
    print(x.dtypes)
    print(x.head())