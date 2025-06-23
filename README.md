# Temperature Mapping Automation System

This repository contains an automation system developed for pharmaceutical equipment and storage space GMP temperature mapping. The system uses iButton data loggers, GUI automation, data processing, API integration, and Word report generation to streamline reporting.

# Project Summary

This tool automates the temperature mapping process for pharmaceutical equipment and storage areas. It handles data collection, file conversion, statistical analysis, and report formatting. The script integrates with iButton data loggers and T-Tech software to process the data, calculate min/max/mean statistics, generate graphs, and assist with report preparation.

# Technologies

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

The tool is modular. The different functions handle:

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
