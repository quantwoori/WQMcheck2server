# Check Data Transporter

## Python Dependency
* pandas
* openpyxl

## External Dependency
* Microsoft Power Automate


## How it's done?
<p>
    CheckExpert+ can be only loaded on to excel.
    It also uses custom function  module that evaporates after use
    so that even if we use win32com, 
    the function would not start.
</p>

<p>
    To counter that rather nasty security feature, we use Microsoft Power Automate(RPA).
    Microsoft Power Automate features none-coding macros. 
    Also CheckExpert+ is not binded by company security programs, which allows us to use RPA features freely.
</p>

## Pipeline
<p>

1. main_data.py (made into batch file)
   1. description: main_data.py writes individual excel functions for k amount of stocks.
   2. TODO: If! there's a holiday? RPA cannot account for that variable.
2. MICROSOFT RPA WORK
   1. RPA opens ./test.xlsx
   2. RPA saves values to separate xlsx files in Users/Documents
   3. RPA close all unnecessary windows
3. main_dins.py
   1. description: insert result.xlsx file from RPA into database
   2. This process only insert RAWborrow data
4. main_dsig.py
   1. description: Using RAWborrow database it calculates signal via signal threshold and insert it inside dbo.sig table.
   2. This is the end of the pipeline

</p>

## TODO
<p>
Let's hope we get proper REST API and websocket style data receiver from koscom in the future.
</p>