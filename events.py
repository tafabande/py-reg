import pandas as pd
import random
import string
import os

file_path = "names.xlsx"
esname = "Events"
EVENT_COLUMNS = ["Name", "Theme", "Date", "Speaker", "Duration", "Attendance", "Rating", "ID"]

def get_new_event_info():
    try:
        return {
            "Name": input("Event Name: "),
            "Theme": input("Event Theme: "),
            "Date": input("Event Date (Y:M:D): "),
            "Speaker": input("Event Speaker: "),
            "Duration": input("Event Duration: "),
            "Attendance": int(input("Amount of People that attended: ")),
            "Rating": int(input("Audience overall rating: "))
        }
    except ValueError:
        print("Attendance and Rating must be numbers.")
        return None

def generate_unique_id(existing_ids):
    while True:
        new_id = ''.join(random.choices(string.ascii_uppercase, k=2)) + \
                 f"{random.randint(0, 999):03}" + \
                 ''.join(random.choices(string.ascii_uppercase, k=2)) + \
                 f"{random.randint(0, 999):03}"
        if new_id not in existing_ids:
            return new_id

def load_or_create_events_sheet(path, sheet, columns):
    if not os.path.exists(path):
        df = pd.DataFrame(columns=columns)
        with pd.ExcelWriter(path, engine="openpyxl", mode="w") as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)
        return df
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except Exception:
        return pd.DataFrame(columns=columns)

def view_previous_events(path, sheet):
    try:
        df = pd.read_excel(path, sheet_name=sheet)
        if not df.empty:
            print(f"\n--- All Previous Event Info from '{sheet}' ---\n".center(60, "-"))
            print(df.to_string(index=False))
        else:
            print(f"No event data found in the '{sheet}' sheet.")
        return df
    except Exception as e:
        print(f"Error reading events: {e}")
        return pd.DataFrame(columns=EVENT_COLUMNS)

def add_new_event(path, sheet, columns):
    event = get_new_event_info()
    if not event:
        return
    df = load_or_create_events_sheet(path, sheet, columns)
    existing_ids = set(df["ID"].dropna().astype(str)) if not df.empty else set()
    event["ID"] = generate_unique_id(existing_ids)
    new_df = pd.DataFrame([event], columns=columns)
    combined = pd.concat([df, new_df], ignore_index=True)

    with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        combined.to_excel(writer, sheet_name=sheet, index=False)

    print("\nEvent information saved successfully.\n")
    print(new_df.to_string(index=False))

def view_calendar():
    print("In Progress: (Aesthetic calendar feature coming soon.)")

def main():
    events_data = None
    while True:
        print("Events Management".center(60, "-"))
        print("1. Log New Event Information")
        print("2. View Previous Event Details")
        print("3. View Event Calendar")
        print("4. Exit")

        try:
            choice = int(input("Choose (1-4): "))
            if choice == 1:
                add_new_event(file_path, esname, EVENT_COLUMNS)
                events_data = view_previous_events(file_path, esname)
            elif choice == 2:
                events_data = view_previous_events(file_path, esname)
            elif choice == 3:
                view_calendar()
            elif choice == 4:
                print("\n--- All Event Data on Exit ---\n".center(60, "-"))
                if events_data is not None and not events_data.empty:
                    print(events_data.to_string(index=False))
                else:
                    print("No event data to display.")
                print("-" * 60)
                print("Goodbye, event overlord.")
                break
            else:
                print("Invalid choice. Pick a number between 1 and 4.")
        except ValueError:
            print("Enter numbers only.")
        except Exception as e:
            print(f"Chaos has struck: {e}")


if __name__ == "__main__":
    main()