# Schedule Generator
#### Description:
The purpose of **Schedule Generator** is to create a work schedule for a
specified calendar month as a formatted, ready-to-print Excel file, thereby
reducing the amount of time spent trying to manually create a work schedule.
The program's functionality was made with a specific company in mind, and as
such is limited in scope.

The program reads in a user-filled template which details the two staff working
on a particular weekday and what hours they're scheduled to work. If the
aforementioned template does not yet exist, the program will instead create a
new, blank template on the Desktop and tell the user to fill it out before
running the program again. If a template is found, the program will prompt the
user for a month and year, then read that template and generate a schedule for
that full calendar month, filling in the calendar with the content found in the
template.

The opening, parsing, writing, editing, and saving of spreadsheet files is
done through a third-party module called openpyxl.
