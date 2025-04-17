import pandas as pd
import random
import string
from datetime import datetime
df =pd.read_excel("names.xlsx")
print(df)


td = input("add users(Y/N): ")
num = int(input("If so how many: "))







if td.capitalize() == ("Y") and num > 0:
    while num > 0:
        exists = set(df["ID"].astype(str)) if not df.empty else set()
        while True:
            three = f"{random.randint(0, 999):03}"
            lets = ''.join(random.choices(string.ascii_lowercase, k=2))
            l2 = f"{random.randint(0, 999):02}"
            newid = three + lets + l2
            if newid not in exists:
                exists.add(newid)
                break
        num-=1
        name1 = input("Please enter your Name: ")
        surname1 = input("please enter your surname: ")
        point1 = int(input("your allocated ponts: "))
        age1 = int(input("Please enter your age: "))
        color1 = input("Please enter your prefered colour: ")
        gender1 = input("Please enter your gender: ")
        time1 = datetime.now().strftime("%H:%M")
        date1 = datetime.now().strftime("%y-%m-%d")
        new = {
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
        df = pd.concat([df, pd.DataFrame([new])], ignore_index=True)
        df.to_excel("names.xlsx", index = False)
        print("new user added")
