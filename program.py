import pandas as pd
from datetime import datetime, timedelta
import random


def read_employee_from_excel(file_name):
    try:
        df = pd.read_excel(file_name)
        print(df)
        print("Read employee list from excel succesful")
        return df
    except FileNotFoundError:
        print("Error: File not found.")
        return None


def add_employee(employee_list):
    name = input("Enter the name of the new employee: ")
    team = input("Enter the team of this employee (Development or Operation): ")
    print(f"New employee: Name: {name}, Team: {team}")
    confirm = input("Confirm addition of this employee? (yes/no): ")
    if confirm.lower() == "yes":
        new_employee = pd.DataFrame(
            {"No.": [len(employee_list) + 1], "Name": [name], "Team": [team]}
        )
        employee_list = pd.concat([employee_list, new_employee], ignore_index=True)
        print("Employee added successfully.")
        print(employee_list)
    else:
        print("Addition canceled.")
    return employee_list


def remove_employee(employee_list):
    name = input("Enter the name of the employee to remove: ")
    print(f"Removing employee: {name}")
    confirm = input("Confirm removal of this employee? (yes/no): ")
    if confirm.lower() == "yes":
        employee_list = employee_list[employee_list["Name"] != name]
        print("Employee removed successfully.")
        print(employee_list)
    else:
        print("Removal canceled.")
    return employee_list


def save_employee_list(employee_list, file_name):
    employee_list.to_excel(file_name, index=False)
    print(f"Employee list saved successfully to {file_name} .")



def select_shift_schedule(employee_list):
    try:
        start_date = input(
            "Enter the starting date for the shift schedule (YYYY-MM-DD): "
        )
        start_date = datetime.strptime(start_date, "%Y-%m-%d")
        if start_date < datetime.now():
            print("Starting date is before the current date. Please try again")
            return
        end_date = input("Enter the ending date for the shift schedule (YYYY-MM-DD): ")
        end_date = datetime.strptime(end_date, "%Y-%m-%d")
        num_days = (end_date - start_date).days
        if num_days < 30:
            print("Shift schedule should be at least 30 days.")
            return
        end_date = start_date + timedelta(days=29)  # Assuming a 30-day shift schedule

        development_team = employee_list.loc[
            employee_list["Team"] == "development", "Name"
        ].tolist()
        operation_team = employee_list.loc[
            employee_list["Team"] == "operation", "Name"
        ].tolist()

        if not development_team or not operation_team:
            print("One or both teams do not have any members.")
            return

        dates = pd.date_range(start=start_date, end=end_date)
        shift_schedule = []

        for date in dates:
            dev_employee = random.choice(development_team)
            operation_employee = random.choice(operation_team)
            shift_schedule.append((date, dev_employee, operation_employee))

        schedule_df = pd.DataFrame(
            shift_schedule, columns=["Date", "Development Team", "Operation Team"]
        )
        schedule_df.to_excel("completed_request_3_2.xlsx", index=False)
        print("Shift schedule saved successfully.")
    except ValueError:
        print("Invalid date format. Please enter the date in YYYY-MM-DD format.")


def main():
    file_name = "Member_teams.xlsx"
    employee_list = read_employee_from_excel(file_name)
    
    if employee_list is not None:
        while True:
            print("\nMenu (1, 2, 3, 4):")
            print("1. Add Employee")
            print("2. Remove Employee")
            print("3. Create Shift Schedule")
            print("4. Exit")
            choice = input("Enter your choice: ")

            if choice == "1":
                employee_list = add_employee(employee_list)
                save_employee_list(employee_list, file_name)
                save_employee_list(
                    employee_list, "completed_request_2.xlsx"
                )  # complete request
            elif choice == "2":
                employee_list = remove_employee(employee_list)
                save_employee_list(employee_list, file_name)
                save_employee_list(
                    employee_list, "completed_request_2.xlsx"
                )  # complete request
            elif choice == "3":
                select_shift_schedule(employee_list=employee_list)
                
            elif choice == "4":
                print("Exit program.")
                break
            else:
                print("Invalid choice. Please enter 1, 2, 3, or 4.")


if __name__ == "__main__":
    main()
