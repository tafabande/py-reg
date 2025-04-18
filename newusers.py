import pandas as pd
import random
import string
from datetime import datetime

def read():
    try:
        return pd.read_excel("names.xlsx")
    except FileNotFoundError:
        return pd.DataFrame(Columns=["Name","surname","Age","Points","color","Gender"])

def write(df,filename="names.xlsx"):
        df.to_excel("names.xlsx", index = False)
        print("new information added")

def idgen(exists):
        while True:
            three = f"{random.randint(0, 999):03}"
            lets = ''.join(random.choices(string.ascii_lowercase, k=2))
            l2 = f"{random.randint(0, 999):02}"
            newid = three + lets + l2
            if newid not in exists:
                exists.add(newid)
                break
        return newid

def info(newid):
    name1 = input("Please enter your Name: ")
    surname1 = input("please enter your surname: ")
    point1 = int(input("your allocated ponts: "))
    age1 = int(input("Please enter your age: "))
    color1 = input("Please enter your prefered colour: ")
    gender1 = input("Please enter your gender: ")
    time1 = datetime.now().strftime("%H:%M")
    date1 = datetime.now().strftime("%y-%m-%d")
    return {
        "Name": name1,
        "Surname": surname1,
        "Age": age1,
        "Points": point1,
        "Color": color1,
        "Gender": gender1,
        "Date of Entry": date1,
        "Time of Entry": time1,
        "ID": newid
    }

def main():
    print("hie")
    filename = "names.xlsx"
    df =read()
    td = input("Do oyu want to add users (Y/N): ")
    if td.upper() != "Y":
        print("cancelled ")
        return
    try:
        num = int(input("How many?: "))
    except ValueError:
        print("Invalid Number: ")
        return
    for i in range(num):
        exists = set(df["ID"].astype(str)) if not df.empty else set()
        newid= idgen(exists)
        new = info(newid)
        df = pd.concat([df, pd.DataFrame([new])], ignore_index=True)
        write(df)

if __name__ == "__main__":
    main()