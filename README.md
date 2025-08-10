# Temperature Mapping Automation System

Automates the GMP temperature mapping workflow for pharmaceutical equipment and storage space. The tool uses iButton data loggers, GUI automation, data processing, API integration, and Word report generation to streamline reporting. I developed this during my time at Biological Therapies and proposed it to management. It was utilized and reduced hours of manual work down to minutes. The repo excludes private images and sensitive information.

# Results

- Reporting time reduced from 4 hours to under 15 minutes
- Improved worker satisfaction by reducing repetitive manual work
- Improved reporting consistency and eliminated the risk of human error
- Saved the company significant time and money

# Project Summary

Pharmaceutical companies that manufacture drug products are legally required to monitor the temperature of storage locations and manufacturing equipment. This is to guarantee that at no stage during manufacture or shipping the temperature of the products exceeded any tolerable limits which can impact drug performance. Each batch requires temperature data to be analysed and documented in a formal report. Previously, this process was manual and extremely time consuming. I proposed and implemented a script-based automation to streamline the process of data acquisition, analysis and reporting which reduced countless hours of work and saved the company significant cost.

This tool handles data collection, file conversion between .dta to .csv, statistical analysis of the data, and then formatting a report. The script automates the download of data from iButton loggers and integrates with T-Tech software to process the data. It uses Pandas to calculate min/max/mean statistics, matplotlib to generate graphs and visuals of the data, and assists with report preparation.

# Tech Stack

- Python
- pyautogui
- pandas
- matplotlib
- opencv-python
- requests
- pywin32
- openpyxl
- win32com for MS Word automation

# Usage

The different functions in the tool are used for:

- Starting and stopping iButton data loggers from recording via GUI automation
- Converting .dta files to .csv files
- Performing statistical calculations (min, max, mean)
- Generating formatted data layouts inside MS Word

# Notes

- Image recognition files are specific to original internal software GUI layout.
- Private credentials, API keys, and internal production tokens are not included.
- Originally developed and used internally at Biological Therapies.

# Setup

pip install -r requirements.txt
