import pandas as pd 
import csv
import os
# To managae paths incompatibility between Windows and Unix 
from pathlib import Path 
from datetime import datetime, timedelta

# Print entire dataframes  
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', None)

# Replace with the path to your Excel file
input_file_path = r"/home/inaig_seyer/Code/Python/Calendar_Automa/test-folder-for-calendar/Horarios semana 27-01-2025.xlsx"  
output_csv_path = r"../csvs"  

# Make the excel file a DataFrame
df = pd.read_excel(input_file_path) 


def cleaningvalues(df):

    # Rename df columns to weekdays
    df.rename(columns={'Unnamed: 0':'Names', 'Unnamed: 1':'Monday', 'Unnamed: 2':'Tuesday', 'Unnamed: 3':'Wendnesday', 'Unnamed: 4':'Thursday',
        'Unnamed: 5':'Friday', 'Unnamed: 6':'Saturday', 'Unnamed: 7':'Sunday'}, inplace=True)

    # # Select the row with my name
    df = df.loc[df['Names'] == 'Giani'] 

    # Change types to str
    df = df.astype(str)

    # Replace characters 
    df = df.map(lambda x: str(x).replace('.',':'))
    df = df.map(lambda x: str(x).replace('h',''))
    df = df.map(lambda x: str(x).replace(' ',''))
    df = df.map(lambda x: str(x).replace('a','-'))
    df = df.map(lambda x: str(x).replace('AC','- 01:30'))
    df = df.map(lambda x: str(x).replace('A','-'))

    # Make it jump the Name column
    df = df.loc[:,'Monday':'Sunday']

    return df



# Transform data
def format_shifts():
    df_copy = cleaningvalues(df)
    
    formated_shifts = []
    
    # if it's not a free day, iterate over it and extract start and end times
    start_n_end_times = df_copy.map(lambda x: x.split('-') if x.lower() != 'x' else 'free day')

    # Store start and end times in a list
    week_shifts = [values.iloc[0] for keys, values in start_n_end_times.items()]

    # Get the date from the name of the file (usually monday) 
    ref_date = input_file_path.split('/')[-1].split(' ')[-1].split('.')[0]

    # Turn string into a list and change the items type to integer  
    d = list(map(int, ref_date.split('-'))) 

    # Iterate through shifts
    for week_day, shift in enumerate(week_shifts):
        if shift == 'free day':
            continue
        
        # Iterate through times in each shift
        for time in shift:
            # In case times come like: "14". Turn it: "14:00"
            if len(time) != 4:
                time += ": 00"

            # split time to build datetime  
            t = time.split(':')
            
            # Get the day of the week of the reference date
            ref_week_day = datetime(d[2], d[1], d[0]).weekday()

            # Get diff between week_day and ref_week_day 
            day_diff = week_day - ref_week_day 
            
            # Build datetime
            formated_datetime = datetime(d[2], d[1], d[0], int(t[0]), int(t[1])) + timedelta(days=day_diff)

            formated_shifts.append(formated_datetime)

    return formated_shifts



# List-group the shifts by partners of two starting from index 0 (eg: [[0,1], [2,3], [3,4]])
def nested_pair_list():
    new_shifts = format_shifts()
    two_items_nested_list = []

    for i in new_shifts:       
        
        # If the loop it's working on the first item, jump to the next   
        if new_shifts.index(i) == 0:
            continue
        # If the index isn't an even number, append the item of the index and the previous one just like: [previous item, not-even index item] 
        if (new_shifts.index(i) % 2) != 0:
           nested_dates = [new_shifts[new_shifts.index(i)-1], new_shifts[new_shifts.index(i)]]
           two_items_nested_list.append(nested_dates)

    return two_items_nested_list


# Set dictionary schema for the csv 
def event_schema(start_time, end_time, start_date, end_date):

    return {
            "Subject": "Trabajo",
            "Start Date": str(start_date.strftime("%Y-%m-%d")),
            "Start Time": str(start_time.strftime("%H:%M:%S")),
            "End Date": str(end_date.strftime("%Y-%m-%d")),
            "End Time": str(end_time.strftime("%H:%M:%S")),
            "All Day Event": "FALSE",   
            "Description": "Jornada laboral",
            "Location": "", 
            "Private": "TRUE"
    }


# Populate date inside dictionaries
def createevents():
 
    dict_event = []
    # Populate event_schema() with the variables
    for shifts in nested_pair_list():
        if shifts == 'free day':
            continue

        start_date, end_date = shifts[0], shifts[1]
        start_time, end_time = shifts[0], shifts[1]
         
        # Save the difference between times in an list
        hours_diff = str(end_time - start_time).split(',')

        # If the end-time of the shift ends passed 00:00 set the the end date to next day  # PROBLEM: Date comparison doesn't work
        if '-1 day' in hours_diff:
            end_date = start_date + timedelta(days=1)
        
        # If there's double shift CREATE TWO SEPARATES EVENTS with [shifts[0], shifts[1]] and [shifts[2], shifts[3]] 
        if len(shifts) > 2:
            start_date, end_date = shifts[2].strftime("%Y-%m-%d"), shifts[3].strftime("%Y-%m-%d")

        events = event_schema(start_time, end_time, start_date, end_date)
        dict_event.append(events)
    
    return dict_event

print(createevents())

# SAVE SHIFTS.CSV IN A DIRECTORY OF NAME SHIFTS

# Create a directory with name "shifts" if don't exits already
try:
    # Path to the directory
    shift_dir_path = Path.cwd() / "shifts" 

    # Check if the directory exists
    if not Path(shift_dir_path).is_dir():
        Path.mkdir(shift_dir_path)

        # Check if the directory was created
        if Path(shift_dir_path).is_dir():
            print("Directory created successfully")

    elif Path(shift_dir_path).is_dir():
        print("Directory already exists")

except  Exception as e:
    print(e)


# Write the csv file with the shifts
with open(shift_dir_path / "shifts.csv", mode='w') as file:
    try:
        writer = csv.DictWriter(file, fieldnames=createevents()[0].keys())

        # Write the header row (createevents()[0].keys())
        writer.writeheader()

        # Write the data rows
        writer.writerows(createevents())

    except Exception as e:
        print(e)