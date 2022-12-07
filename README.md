# spreadsheet_translate


   Translates the spreadsheet - easy and quick
   by Aleksandrov M. 2022 in Amsterdam


   This program allows to translate entire Excel workbooks/spreadsheets while preserving the format of original file. 
   It also doesn't break math or functions in excel, just translates the text. The output is a new translated .xslx file.
    
    
   E.g. it useful when you are working in a different country.
    
   How to use? That's how:




   >> ./spreadsheet_translate.py -n file.xlsx -c "nl" -t "en"               Does everything but via the terminal

   Args:  -n   name of the file
          -c   current language of the spreadsheet (e.g. nl, en)
          -t   target language of the spreadsheet




   from spreadsheet_translate import Spreadsheet                            import the directory

   spreadsheet = Spreadsheet("file.xlsx", current="nl", target="en")        Open the workbook
   spreadsheet.translate()                                                  Translate all the spreadsheets
   spreadsheet.save()                                                       Save the workbook in the new file



   IMPORTANT: The translation is possible due to the free directory "translate",
   if your workbook/spreadsheet is too big you will run into an issue of
   "MYMEMORY WARNING: YOU USED ALL AVAILABLE FREE TRANSLATIONS FOR TODAY". So, to avoid it
   try to use a paid service and slightly change the code. Personally, I use API from Microsoft Azure,
   very easy to integrate into the code (±5 minutes). I have made commentaries inside the code,
   to show where the translation itself happens.
