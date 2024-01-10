import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta

global paq_gsp
global debt_collection
global education_goal
global funding_req
global acq_civ
global RPAs
global My_FSS
global Qual_revs


# Read the Excel file into a DataFrame
RPAs = pd.read_excel("FY24 Force Management Tracker.xlsx", sheet_name="RPAs")
debt_collection = pd.read_excel("FY24 Force Management Tracker.xlsx", sheet_name="PAQ Debt")
My_FSS = pd.read_excel("FY24 Force Management Tracker.xlsx", sheet_name="MyFSS Announcements")
acq_civ = pd.read_excel("FY24 Force Management Tracker.xlsx", sheet_name="ACQ CIV TA")
paq_gsp = pd.read_excel("FY24 Force Management Tracker.xlsx", sheet_name="GSPs Only")
funding_req = pd.read_excel("FY24 Force Management Tracker.xlsx", sheet_name="PAQ Tracking") # CTAP & PAQ tracking
Qual_revs = pd.read_excel("FY24 Force Management Tracker.xlsx", sheet_name="Qual Reviews")
education_goal = pd.read_excel("FY24 Force Management Tracker.xlsx", sheet_name="STEM+M") #aquiv Civ ta, PAQ tracking, CTAP





def generate_graph(df):
    # Assuming your date column is named 'Date'
    df['Date'] = pd.to_datetime(df['Date'])

    # Group by week and calculate some metric (e.g., sum, count) for each week
    weekly_data = df.groupby(df['Date'].dt.isocalendar().week).sum()

    # Plotting the graph
    plt.plot(weekly_data.index, weekly_data['Your_Column_To_Plot'], marker='o')
    plt.xlabel('Week')
    plt.ylabel('Your Y-Axis Label')
    plt.title('Weekly Data Graph')
    plt.show()


def generate_weekly_tables(df):
    # Assuming your date column is named 'Date'
    df['Received'] = pd.to_datetime(df['Received'], exact=False)
    data = {'Received': pd.date_range(start='2023-01-04', end='2024-12-27')}
    df = pd.DataFrame(data)
    # Create a new column for the week number
    df['Week'] = df['Received'].dt.isocalendar().week
    # Find the first Wednesday of the year
    start_date = df['Received'].loc[
        df['Received'].dt.day_name() == 'Wednesday'].min()

    # Find the last Wednesday of the year
    end_date = df['Received'].loc[
        df['Received'].dt.day_name() == 'Wednesday'].max()
    # Create a table for each week
    weekly_data = []
    curr_date = start_date
    # Print or manipulate the tables as needed
    while curr_date <= end_date:
        current_data = df.rolling('7D', on='Received').count

        # Process the data as needed
        # (Add your processing logic here)

        # Append the data to the list
        weekly_data.append(current_data)

        # Move to the next Wednesday
        curr_date += timedelta(days=7)

    with pd.ExcelWriter(
            "output.xlsx",
            mode="w",
            engine="openpyxl",
            date_format='YYYY-MM-DD',
            datetime_format='YYYY-MM-DD'
            # if_sheet_exists="overlay",
    ) as writer:
        weekly_data.to_excel(writer, sheet_name="MyFSS pivot").close()



if __name__ == "__main__":
    # Replace 'your_file.xlsx' with the path to your Excel file
    generate_weekly_tables(My_FSS)

