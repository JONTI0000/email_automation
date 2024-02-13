import pandas as pd
from datetime import datetime
from main import EmailSender

def load_data():
    #loading data
    df = pd.read_csv("shedule.csv")
    df["date"] = pd.to_datetime(df["date"],dayfirst=True)
    df['step'] = df['step'].astype(str).str.zfill(2)
    current_date = datetime.now()

    #getting the closest date
    closest_future_date = min(df["date"], key=lambda x: abs((x - current_date).days))
    closest_future_date.date()

    #loading the closest date to the current date
    records = df[df["date"] == closest_future_date]
    return records

def getting_input():
    records =load_data()
    instances = []
    for index,row in records.iterrows():
        instance= EmailSender(name=row["name"],batch_no=row["batch no"], email_address=row["email address"], subject=row["subject"], step=row["step"])
        instances.append(instance)
    return instances
