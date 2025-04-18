from datetime import datetime
import time
import pandas as pd
import string
import sys
import random



def read(file_path="names.xlsx"):
    try:
        df = pd.read_excel(file_path, dtype={'Contact': str})
    except FileNotFoundError:
        df = pd.DataFrame(columns=["Name", "Surname", "Age", "Points", "Color", "Gender", "Events"])

    if 'Events' in df.columns:
        df['Events'] = df['Events'].fillna(0).astype(int)


    df.columns = df.columns.str.strip()
    if "Events" not in df.columns:
        df["Events"] = 0
    return df


def write(df, file_path="names.xlsx"):
    try:
        df.to_excel(file_path, index=False)
        print("New information saved.")
    except PermissionError:
        print("Permission Denied: The file is currently open or in use.")
        response = input("Please close the file and press Enter to try again, or type 'exit' to cancel: ").strip().lower()
        if response != 'exit':
            time.sleep(2)
            write(df, file_path)
        else:
            print("Save operation cancelled.")
def vam():
    df = read()
    if df.empty:
        print("No members found.")
    else:
        print("\n--- All Registered Members ---\n".center(60, "-"))
        print(df.to_string(index=False, justify='center'))
        print("-" * 50)

def srch(df,names,surnames):
    df = pd.read_excel("Names.xlsx")
    print(df[(df["Name"] == names) & (df["Surname"] == surnames)])


def mad():
    df = read()
    try:
        num = int(input("How many?: "))
    except ValueError:
        print("Invalid Number: ")
        return
    for i in range(num):
        exists = set(df["ID"].astype(str)) if not df.empty else set()
        newid = idgen(exists)
        new = info(newid)
        df = pd.concat([df, pd.DataFrame([new])], ignore_index=True)
        write(df)


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
    name1 = input("Please enter Name: ")
    surname1 = input("please enter surname: ")
    point1 = int(input("Allocated ponts: "))
    age1 = int(input("Please enter age: "))
    color1 = input("Please enter prefered colour: ")
    gender1 = input("Please enter gender: ")
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
def nan():
    try:
        return int(input("How many people attended: "))
    except ValueError:
        print("Please enter a valid number: ")
        sys.exit()

def nmc(df, names, surnames):
    if names in df["Name"].values:
        match = df[(df["Name"] == names) & (df["Surname"] == surnames)]
        index = match.index[0]
        current = df.at[index, "Events"]
        if pd.isna(current):
            df.at[index, "Events"] = 1
        else:
            df.at[index, "Events"] = int(current)+1

        print("Is a member \n Attendence marked")
        id = match["ID"].values[0]
        print(f"{surnames} {names} is  a member with ID [{id}]")
        return True
    else:
        print("Not yet a member")
        return False

def MaT():
    file_path = "names.xlsx"
    df = read(file_path)

    num = nan()
    if num== 0:
        print("No one attended.")
        sys.exit()

    unreg = 0
    for i in range(num):

        name1 = input("Please enter the name: ")
        surname1 = input("Please enter the Surname: ")
        if not nmc(df,name1,surname1):
            unreg += 1
        write(df, file_path)
        print(f"There are {unreg} unregistered attendees.\nPlease register them.")

def main():
    print("hie")
    filename = "names.xlsx"
    df = read()
    print("What do you want to do?")
    print("1. Add Members")
    print("2. Search for a Members")
    print("3. Mark Attendence")
    print("4. View All Members")
    td = int(input("Please select what you want to do (1-4): "))
    if td not in (1, 2, 3, 4):
        sys.exit("cancelled ")
    elif td == 1:
        mad()
    elif td == 2:
        name = input("Enter first name: ")
        surname = input("Enter surname: ")
        srch(df, name, surname)

    elif td == 3:
        MaT()
    elif td == 4:
        vam()


if __name__ == "__main__":
    main()