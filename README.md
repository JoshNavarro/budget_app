# Budget Application

A simple application utilizing Google sheets to help budget everyday expenses

## Introduction
This budgeting application helps users to automatically input bank statements into a spreadsheet hosted on Google Sheets

## How it works
* Grabs recently downloaded csv files
* Differentiates between Debit and Credit bank statements
* Parses and saves useful information from statements such as: Transcation name, date, and cost
* Calls Google APIs to add information into existing spreadsheet
* Automatically sets category of transaction based on similar transactions
* Opens a Google Chrome tab with your budget spreadsheet

## Installation Steps
1. Make sure Git is installed
2. Make sure python3 is installed
3. Open terminal application
4. Create a directory in which you would like the script to live
5. From the CLI execute he following command:<br>
   `git clone git@github.com:JoshNavarro/budget_app.git`

## Using the application
1. Download pre-made budgeting spreadsheet <Google sheet links to be updated soon>
2. Fill in the cateogries you wish to budget for
3. Log in to online banking systems and download csv files of statements
4. Open terminal and navigate to the directory where this application lives: <br>
   `cd /Users/joshnavarro/Documents/budgeting/`
5. Execute the command: `python budget.py`
