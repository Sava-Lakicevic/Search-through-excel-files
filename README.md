# Search-through-excel-files
This Python script searches through all of the excel files in the current directory and all of its subdirectories. The scope of search depends on where it is launched from. I recommend to run it through the command prompt, change in to the directory where you need to run it from (unless you want to search through the entire computer). The output is the parent folder, the name of the file, the name of the sheet, the row number (starting from 0) and the column number (starting from 0). It take the input through the command prompt with the message "What are you looking for: ", reads through all of the files.

Output example:
  File: ['folder_name', 'file_name.xlsx']; Sheet: Sheet1; Location: (42,5)

It will run until a keyboard iterrupt (usually ctrl+C) or if you enter "end of work" into the input.
