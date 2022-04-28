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

## How is the Signal Made?
<p>
   Original Plan: Use the autoencoder to detect whether or not the path of the signal 
      is out of the ordinary.
</p>

<p>
   Revised Plan: Using autoencoder takes up so much time, since we have to fit autoencoder to each and every stocks. 
      Therefore, we are using band methodology. If the signal goes outside of the bollingerband, 
      it is considered as "out of the ordinary". Result shows that the bollingerband methodology,
      produces similar result for initial signal generation - even faster at some point. 
      However, it does not provide long and stable signal like that of autoencoder. 
   Since marketing looooves "MACHINE LEARNING", we will state that we are using AutoEncoder based signal generating method.
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