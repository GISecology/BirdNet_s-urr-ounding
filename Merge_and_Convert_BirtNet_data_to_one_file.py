# coding=utf8

import pandas
import os
import glob

# variable
## folders
folder_with_birdnet_csv = r"C:\Users\juergen.foerth\Documents\PythonProjekte\Aichtal"
folder_for_merged_xlsx = r"C:\Users\juergen.foerth\Documents\PythonProjekte"

## Local time
timezone = 'Europe/Berlin'

# get workspace
workspace_folder = os.getcwd()
print(os.getcwd())

# returns the file list
csv_file_list = glob.glob(os.path.join(folder_with_birdnet_csv, '*.csv'))
print(csv_file_list)
frames = []
counter = 0
for file in csv_file_list:
    print(file)
    df = pandas.read_csv(file)
    f_name = os.path.basename(file)
    print(f_name)

    # slice file name into project and datetime -strings
    datetime = f_name[f_name.find("_")+1:-20]
    project = f_name[:f_name.find("_")]
    date = datetime[:8]
    time = datetime[9:]
    print(project)
    print(datetime)
    print(date)
    print(time)

    date_formated = date[-4:-2] + "." + date[-2:] + "." + date[:4]
    time_formated = time[:2] + ":" + time[2:4] + ":" + time[4:6]
    df['date_time'] = pandas.to_datetime(date_formated + " " + time_formated).strftime("%d.%m.%Y %H:%M:%S")
    df['Date'] = pandas.to_datetime(date_formated)
    df['Start (s)'] = df['Start (s)'].astype('float64')

    # Spalte Aufnahmezeitpunkt mit Uhrzeit + Startzeit (Tracktime)
    startzeit = pandas.to_datetime(df["Start (s)"], unit='s').dt.strftime("%H:%M:%S")
    df['marker'] = startzeit
    df['start_time'] = pandas.to_datetime(df['date_time']) + pandas.to_timedelta(df['marker'])

    # convert time to local timezone (timezone definition on top of the script)
    df['local_time'] = pandas.to_datetime(df['start_time'], utc=True).map(lambda x: x.tz_convert(timezone))
    df['local_time'] = df['local_time'].astype('string')

    # Dateiname und Pfad
    df['filename'] = f_name
    df['path'] = file

    # export
    print(df.head(12))
    frames.append(df)


foldername = os.path.basename(folder_with_birdnet_csv)

# Export zu CSV
result = pandas.concat(frames)
print("Create CSV")
result.to_csv(folder_for_merged_xlsx + "\\" + foldername + ".csv")

# Export zu Excel
print("\nCreate xlsx...")
# Get the xlsxwriter workbook and worksheet objects.
# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pandas.ExcelWriter(folder_for_merged_xlsx + "\\" + foldername + ".xlsx", engine="xlsxwriter", date_format="dd.mm.yyyy", datetime_format="dd.mm.yyyy HH:MM:SS")

# Convert the dataframe to an XlsxWriter Excel object.
result.to_excel(writer, sheet_name="Sheet1")
workbook = writer.book
worksheet = writer.sheets["Sheet1"]

# Add some cell formats.
format1 = workbook.add_format({"num_format": "#,###0.000"})
format_datetime = workbook.add_format({"num_format": "dd.mm.yyyy HH:MM:SS"})
format_date = workbook.add_format({"num_format": "dd.mm.yyyy"})
format_time = workbook.add_format({"num_format": "HH:MM:SS"})

# Set the column width and format.
worksheet.set_column(5, 5, 10, format1)

# Set the format but not the column width.
worksheet.set_column(6, 6, 20, format_datetime)
worksheet.set_column(7, 7, 20, format_date)
worksheet.set_column(8, 8, 10, format_time)
worksheet.set_column(9, 9, 20, format_datetime)

# Close the Pandas Excel writer and output the Excel file.
writer.close()
print("\n...finished.")
print("writer closed")
