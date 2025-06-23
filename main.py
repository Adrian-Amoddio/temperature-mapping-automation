from pathlib import Path
import pyautogui
import os
import time
import cv2
import numpy as np
import shutil
import pandas as pd
import matplotlib.pyplot as plt
import requests
import win32com.client as win32


# Locates and clicks a button on the screen by referring to image files.
def click_button(btn_image, time_delay=0, adjust_x=0, cks=4, confidence=0.95):

    # btn_image - Path to the image file of the button.
    # time_delay - Optional delay (in seconds) after clicking.
    # adjust_x - Optional X-coordinate adjustment
    # cks - Number of clicks to perform.
    # confidence - Confidence threshold for image recognition.

    try:
        location = pyautogui.locateOnScreen(btn_image, confidence=confidence)
        if location:
            center_x, center_y = pyautogui.center(location)
            center_x += adjust_x  # optional x adjustment

            pyautogui.moveTo(center_x, center_y)
            pyautogui.click(clicks=cks)
            print(
                f"CLICK - {btn_image} was clicked at ({center_x}, {center_y})")
            time.sleep(time_delay)
        else:
            print(
                f"MISS - Button not found on screen: {btn_image} (confidence={confidence})")
    except Exception as e:
        print(f"ERROR - Error while trying to click {btn_image}: {e}")


# Stop the iButton sensor logging and save the data
def stop_save_data(directory):

    click_button("images/refresh.png", 6)
    click_button("images/stop_logger.png", 1.5)
    pyautogui.press("enter")
    time.sleep(0.5)
    click_button("images/autoload.png", 1.5, adjust_x=20)
    pyautogui.write(directory)
    pyautogui.press("enter")
    time.sleep(0.5)
    click_button("images/save.png", 1.5, adjust_x=20)
    pyautogui.press("enter")
    time.sleep(0.5)
    pyautogui.hotkey("alt", "f4")


# Once plugged into reader, starts the iButton sensor logging temperature data
def start_ibutton():

    click_button("images/refresh.png", 3)
    click_button("images/start_logger.png", 3)
    pyautogui.press("enter")
    time.sleep(0.5)
    wait_for_button("images/delay_start.png", 1)
    for x in range(11):
        pyautogui.press("tab")
    pyautogui.press("tab")
    time.sleep(0.5)
    pyautogui.press("enter")
    pyautogui.press("0")
    for x in range(8):
        pyautogui.press("tab")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(3)
    pyautogui.press("enter")


# Prints temperature data files and statistics to be reviewed and hand signed
def data_print(directory):

    dir = Path(f"{directory}/Printed")
    dir.mkdir(parents=True, exist_ok=True)

    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)

        if os.path.isfile(file_path):
            print(f"Processing file: {file_path}")

            try:
                # Open the file
                os.startfile(file_path)
                # Wait for file to open
                time.sleep(5)

                # Wait for the 'File' button to appear
                wait_for_button("images/file.png")

                # Go through the process of printing the file
                click_button("images/file.png", 1)
                click_button("images/inst_print.png", 1)
                time.sleep(0.75)
                click_button("images/landscape.png", 1)
                time.sleep(0.75)
                click_button("images/portrait.png", 1)
                time.sleep(0.75)
                click_button("images/print.png", 1)
                time.sleep(0.75)
                click_button("images/inst_print.png", 1)
                time.sleep(0.75)
                click_button("images/landscape.png", 1)
                time.sleep(0.75)
                click_button("images/portrait.png", 1)
                time.sleep(0.75)
                click_button("images/print.png", 1)
                time.sleep(0.75)

                # Wait for the 'analysis' button to appear
                wait_for_button("images/analysis.png")
                click_button("images/analysis.png", 2)
                time.sleep(0.75)
                click_button("images/statistics.png", 2)
                time.sleep(0.75)

                # Wait for the 'second print' button to appear
                wait_for_button("images/2nd_print.png")
                click_button("images/2nd_print.png", 1)
                time.sleep(0.75)
                click_button("images/landscape.png", 1)
                time.sleep(0.75)
                click_button("images/portrait.png", 1)
                time.sleep(0.75)
                click_button("images/print.png", 1)

                time.sleep(1.75)

                click_button("images/print.png")

                # Wait for 35 seconds to allow maximum amount of time for printing to complete (including delays)
                print("Waiting for 40 seconds while printing...")
                time.sleep(40)

                # Close the application
                pyautogui.hotkey("alt", "f4")
                time.sleep(3)
                pyautogui.hotkey("alt", "f4")

                # Log the movement of the file
                print(f"File printed and moved: {file_path}")
                # Move the file to the printed directory
                shutil.move(file_path, f"{directory}/Printed/{filename}")

                time.sleep(3)

            except Exception as e:
                # Catch any exceptions and log an error message
                print(f"Error processing {file_path}: {e}")


# Loops until a specific button appears on the screen or times out.
def wait_for_button(image_path, timeout=30):
    """
    Wait for a button (image) to appear on the screen.
    If the button does not appear within the timeout period, raise an exception.
    """
    start_time = time.time()
    while True:
        location = pyautogui.locateOnScreen(
            image_path, confidence=0.8)  # Adjust confidence if necessary
        if location:
            return True  # Button found
        if time.time() - start_time > timeout:
            raise TimeoutError(f"Button not found: {image_path}")
        time.sleep(1)  # Wait 1 sec before retrying


# Converts all .dta files in a directory to .csv files using T-Tech.
def dta_to_csv(directory):
    csv_directory = 'csv'
    os.makedirs(f"{directory}\\{csv_directory}", exist_ok=True)

    for filename in os.listdir(directory):
        if os.path.isfile(os.path.join(directory, filename)):
            print(filename)

            file_path = f"{directory}\\{filename}"
            os.startfile(file_path)
            time.sleep(2)
            click_button("images/file.png", 0.5)
            click_button("images/list_curve.png", 1.5)
            click_button("images/comma.png", 0.5)
            click_button("images/date.png", 0.5)
            click_button("images/csv_folder.png", 0.5)
            click_button("images/save.png", 0.5, cks=2)
            time.sleep(1)
            pyautogui.hotkey("alt", "f4")
            time.sleep(1)
            click_button("images/window_bar.png")
            time.sleep(1)
            pyautogui.hotkey("alt", "f4")
            time.sleep(2)


# Prepares and filters data from CSV files in a directory
def prepare_data(directory, data_length):
    os.makedirs(f"{directory}", exist_ok=True)

    for filename in os.listdir(directory):
        if os.path.isfile(os.path.join(directory, filename)):
            print(filename)

            file_path = f"{directory}\\{filename}"

            with open(f"{file_path}", "r") as file:
                lines = file.readlines()

            if lines[0] == "No., Temperature,G, Rel., Date, Time \n":
                print("Data has already been prepared.")
                break
            else:
                with open(f"{file_path}", "w") as file:
                    file.writelines(lines[16:])

                with open(f"{file_path}", "r+") as file:
                    lines = file.read()
                    file.seek(0)
                    file.write(
                        "No., Temperature,C, Rel., Date, Time\n" + lines)
                    file.truncate()

                with open(f"{file_path}", "r") as file:
                    lines = file.readlines()

                line_count = 0
                with open(f"{file_path}", "r") as file:
                    for line in file:
                        line_count += 1

                with open(f"{file_path}", "w") as file:
                    file.writelines(lines[:(-1*(line_count - data_length))+1])

                df = pd.read_csv(f"{file_path}", sep=",")
                print(df)


# Creates an excel table with minimum, maximum, and mean temperatures for each iButton
def min_max_mean_table(directory, data_register):
    os.makedirs(f"{directory}\\Statistics", exist_ok=True)
    ibutton_serial = {}
    ibutton_register = pd.read_csv(data_register)

    for index, row in ibutton_register.iterrows():
        ibutton_serial[row['i-button serial']] = row['i-button No.']

    columns = ['i-button', 'Min', 'Max', 'Mean']
    data = []

    for key in ibutton_serial:
        for file_name in os.listdir(directory):
            if file_name[:-4] == key:
                df = pd.read_csv(os.path.join(directory, file_name))
                temperature_min = round(df['Temperature.C'].min(), 2)
                temperature_max = round(df['Temperature.C'].max(), 2)
                temperature_mean = round(df['Temperature.C'].mean(), 2)
                ibutton = ibutton_serial[key]

                data.append([ibutton, temperature_min,
                            temperature_max, temperature_mean])
                print(
                    f"For file: {file_name} ----- Added Line: {ibutton}, {temperature_min}, {temperature_max}, {temperature_mean}")

    stats_results = pd.DataFrame(data, columns=columns)

    # Save as an Excel file
    # Specify your file path
    output_path = f"{directory}\\Statistics\\stats_results.xlsx"
    stats_results.to_excel(output_path, index=False, engine='openpyxl')

    print(f"Results saved as {output_path}")

    return stats_results


# Generates comparison graphs for temperature data (images to be copied into final report in results section)
def gen_graphs(directory, data_register, temp_limit):
    os.makedirs(f"{directory}\\Figures", exist_ok=True)

    # Read the iButton register CSV file
    df_register = pd.read_csv(data_register)

    # Create dictionary to group iButtons by shelf (using the serial number as the key)
    shelves = {
        'Bottom': df_register[df_register['Shelf'] == 'Bottom'],
        'Middle': df_register[df_register['Shelf'] == 'Middle'],
        'Top': df_register[df_register['Shelf'] == 'Top']
    }

    # Extracts temperature data from a given file (CSV)
    def extract_temperature_data(file_path):
        df = pd.read_csv(file_path)
        temperature_data = df['Temperature.C']
        return temperature_data

    # Plots temperature data for each shelf and save the figures
    def plot_temperature_for_shelves(directory, shelves):
        for shelf, shelf_data in shelves.items():
            plt.figure(figsize=(14, 6))
            stats_dict = {}

            for idx, row in shelf_data.iterrows():
                serial_number = row['i-button serial']

                for file_name in os.listdir(directory):
                    if serial_number in file_name:
                        file_path = os.path.join(directory, file_name)
                        temp_data = extract_temperature_data(file_path)

                        min_temp = temp_data.min()
                        max_temp = temp_data.max()
                        mean_temp = temp_data.mean()

                        stats_dict[row['i-button No.']] = {
                            'min': min_temp,
                            'max': max_temp,
                            'mean': mean_temp
                        }

                        label = f"{row['i-button No.']} ({serial_number})"
                        plt.plot(temp_data, label=label)
                        plt.yticks(np.arange(-20, 25 + 1, 1))

            plt.title(f"{shelf} Shelf Temperature Data", fontsize=16)
            plt.xlabel("Time (30 min intervals)", fontsize=12)
            plt.ylabel("Temperature (°C)", fontsize=12)
            plt.grid(True)

            stats_text = ""
            for i_button, stats in stats_dict.items():
                stats_text += f"{i_button}\n"
                stats_text += f"  Min: {stats['min']}°C\n"
                stats_text += f"  Max: {stats['max']}°C\n"
                stats_text += f"  Mean: {stats['mean']}°C\n"

            plt.gca().text(1.05, 0.5, stats_text, transform=plt.gca().transAxes, fontsize=9,
                           verticalalignment='center', horizontalalignment='left',
                           bbox=dict(facecolor='white', edgecolor='black', boxstyle='round,pad=1'))

            plt.subplots_adjust(right=1)
            plt.tight_layout()
            plt.savefig(f"{directory}\\Figures\\{shelf}.png")
            plt.close()

    # Call the plotting function
    plot_temperature_for_shelves(directory, shelves)


# Fetches weather data from Moorabbin Airport for a specified date range and saves it to an Excel file.
def fetch_bom_data(output_dir, start_date, end_date):  # date format: 'YYYY-MM-DD'
    os.makedirs(f"{output_dir}\\MoorabbinAirportData", exist_ok=True)

    # Moorabbin Airport coordinates
    lat = -37.9758
    lon = 145.1020

    url = (
        f"https://meteostat.p.rapidapi.com/point/daily?"
        f"lat={lat}&lon={lon}&start={start_date}&end={end_date}&alt=20"
    )

    headers = {
        "X-RapidAPI-Host": "meteostat.p.rapidapi.com",
        "X-RapidAPI-Key": "your-api-key"
    }

    response = requests.get(url, headers=headers)

    ma_data = []
    columns = ['Date', 'Max Temperature°C']

    if response.status_code == 200:
        data = response.json()
        for day in data['data']:
            ma_data.append([day['date'], day['tmax']])
    else:
        print(
            f"Failed to retrieve data: {response.status_code} - {response.text}")

    df = pd.DataFrame(ma_data, columns=columns)
    output_path = f"{output_dir}\\MoorabbinAirportData\\moorabbin_airport_temperatures.xlsx"

    df.to_excel(output_path, index=False, engine="openpyxl")

    print(f"Results saved as {output_path}")


def make_txt_boxes(directory, output_dir):
    # Load Excel data
    df = pd.read_excel(directory)

    # Start MS word
    word = win32.Dispatch("Word.Application")
    word.Visible = True  # Set to True to see it while debugging

    # Create a new word document
    doc = word.Documents.Add()

    # Layout and spacing settings
    boxes_per_row = 4
    box_width = 120
    box_height = 35
    horizontal_spacing = 20
    vertical_spacing = 30
    margin_left = 50
    margin_top = 50

    for i, row in df.iterrows():
        # Determine row and column
        row_num = i // boxes_per_row
        col_num = i % boxes_per_row

        # Calculate position
        left = margin_left + col_num * (box_width + horizontal_spacing)
        top = margin_top + row_num * (box_height + vertical_spacing)

        # This is the temperature data that will go in each textbox
        text = f"Min: {row['Min']}°C Max: {row['Max']}°C Mean: {row['Mean']}°C"

        # Add a textbox shape to the document
        shape = doc.Shapes.AddTextbox(
            Orientation=1,
            Left=left,
            Top=top,
            Width=box_width,
            Height=box_height
        )

        # Add text with the temperature data to the textbox
        text_range = shape.TextFrame.TextRange
        text_range.Text = text
        text_range.Font.Size = 12
        text_range.Font.Bold = True
        text_range.ParagraphFormat.LineSpacingRule = 0  # Single line spacing
        text_range.ParagraphFormat.SpaceBefore = 0
        text_range.ParagraphFormat.SpaceAfter = 0

        shape.Fill.Visible = False
        shape.Line.Visible = False

        print(f"Added box {i+1} at row {row_num}, column {col_num}")

        # Save the document
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, "stats.docx")
        doc.SaveAs(output_path)

        # Close word and log the output path
        doc.Close()
        word.Quit()
        print(f"Word document saved to: {output_path}")
