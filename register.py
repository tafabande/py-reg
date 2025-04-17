import sys

import pandas as pd
import os

print("current directory: ",os.getcwd())
df =pd.read_excel("names.xlsx")
df.columns =df.columns.str.strip()
if "Events" not in df.columns:
    df["Events"] = 0
num = int(input("how many people attended: "))
unreg = 0
if num==0:
    sys.exit()
for i in range(num):
    name1 = input("Enter Name: ")
    if name1 in df["Name"].values:
        match = df[df["Name"] == name1]
        index = df[df["Name"] == name1].index[0]
        current = df.at[index, "Events"]
        if pd.isna(current):
            df.at[index, "Events"] = 1
        else:
            df.at[index, "Events"] = int(current)+1

        print("Already a member \n Attendence marked")
        id = match["ID"].values[0]
        print(id)
        df.to_excel("names.xlsx", index = False)
    else:
        print("not yet a member")
        unreg +=1
    print(f"There are {unreg} people that attended but are unreg\n Please register them")




