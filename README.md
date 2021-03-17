# Automated E-mail Reminder (AER)

## A sender bot based in Python

![gif](https://github.com/Meluiscruz/Automated_Email_Reminder/blob/master/Images/email_bot.gif)

## Project metrics

![Stars](https://img.shields.io/github/stars/Meluiscruz/Automated_Email_Reminder.svg)
![Forks](https://img.shields.io/github/forks/Meluiscruz/Automated_Email_Reminder.svg)
![Issues](https://img.shields.io/github/issues/Meluiscruz/Automated_Email_Reminder.svg)
![Tags](https://img.shields.io/github/tag/Meluiscruz/Automated_Email_Reminder.svg)

## Scope of the project

The main purpose of this project is to build messages and attachment files from some Excel files and send them to the clients as invoicing reminders.

This project is the first part of an RPA system developed in Python and uses an inner file manager for sorting pending files from submitted ones. The outputs of this script are the inputs of another automated process.

## Table of Contents

- [Python_Scripts](https://github.com/Meluiscruz/Automated_Email_Reminder/tree/master/Python_Scripts "Python_Scripts"):
  - [credentials.py](https://github.com/Meluiscruz/Automated_Email_Reminder/blob/master/Python_Scripts/credentials.py "credentials.py"): Edit this file with the sender data (email and password account).
  - [main.py](https://github.com/Meluiscruz/Automated_Email_Reminder/tree/master/Python_Scripts "main.py"): This is the head pice of code of the project.
- [input_files](https://github.com/Meluiscruz/Automated_Email_Reminder/tree/master/input_files "input_files"):
  - [providers](https://github.com/Meluiscruz/Automated_Email_Reminder/tree/master/input_files/providers "providers"): Where directories of clients (receivers information) are.
  - [pending_base_file](https://github.com/Meluiscruz/Automated_Email_Reminder/tree/master/input_files/pending_base_file "pending_base_file"): Where base files (patients, services and clinics information) are.
- [output_files](https://github.com/Meluiscruz/Automated_Email_Reminder/tree/master/output_files "output_files"):
  - [Reports](https://github.com/Meluiscruz/Automated_Email_Reminder/tree/master/output_files/Reports "Reports"): Where the summaries created by the bot are.
  - [Submitted](https://github.com/Meluiscruz/Automated_Email_Reminder/tree/master/input_files/pending_base_file "pending_base_file"):
    - [E1P1_base_file](https://github.com/Meluiscruz/Automated_Email_Reminder/tree/master/output_files/Submitted/E1P1_base_file "E1P1_base_file"): Where base files are moved after the process.
    - [E1P2_Submitted_files](https://github.com/Meluiscruz/Automated_Email_Reminder/tree/master/output_files/Submitted/E1P2_Submitted_files "E1P2_Submitted_files"): Where files created by the bot and used for the next process are.
  - [Tables](https://github.com/Meluiscruz/Automated_Email_Reminder/tree/master/output_files/Tables "Tables"): Where attachments (html and csv files created by the bot) are.
- [Images](https://github.com/Meluiscruz/Automated_Email_Reminder/tree/master/Images "Images"): Where files used in README.md are.

## What things are pending to develop?

- A script that installs all the required modules for the project. 
- A script that asks the user for sender credentials (e-mail address, password and port).
- A script that sets the absolute path of the project (enviroment variable DEFAULT_DIR in main.py, for Windows and Linux).
- Handle the most of excepetions in main.py.
- The main.py script should work as 'lazy'. So, files must not be created if all the messages are not sent.

## Technical information

I used VSCode as source-code editor. Also, I have employed the following modules listed in [main.py](https://github.com/Meluiscruz/Automated_Email_Reminder/tree/master/Python_Scripts "main.py") header:

- ```numpy==1.20.1```
- ```openpyxl==3.0.6```
- ```pandas==1.2.2```
- ```xlrd==2.0.1```
