import pandas as pd
from datetime import datetime
from main import EmailSender

def finding_closest_date(df:pd.DataFrame):
    current_date = datetime.now()
    #getting the closest date
    closest_future_date = min(df["date"], key=lambda x: abs((x - current_date).days))
    return closest_future_date

def get_the_time(df:pd.DataFrame,date):
    records = df[df["date"] == date]
    time = records['time'].value_counts().idxmax()
    return time


def load_data():
    #loading data
    df = pd.read_csv("shedule.csv")
    df["date"] = pd.to_datetime(df["date"],dayfirst=True)
    df['step'] = df['step'].astype(str).str.zfill(2)
    return df
    
    
def getting_input():
    df = load_data()
    #getting the closest date
    closest_future_date = finding_closest_date(df)
    #loading the closest date to the current date
    records = df[df["date"] == closest_future_date]

    instances = []
    for index,row in records.iterrows():
        instance= EmailSender(name=row["name"],batch_no=row["batch no"], email_address=row["email address"], subject=row["subject"], step=row["step"])
        instances.append(instance)
    return instances
