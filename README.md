# Libcat

This is a tool that searches OCLC Worldcat Search API. It then writes the results into an Excel file. 

*Currently a **work in progress**, still adding functionality, input/output validation, and other stuffs.*

### How to use

Replace the Excel file name with your own file name. It is currently set up as `template.xslx`. Make sure this Excel file is in the same directory as the `finder.py` file.

The terminal prompts the user for an ISBN and returns the title, author, and edition of the book.

#### Things I learned 
This was a nice way for me to learn how to work with a web scraper (`requests`), using an API (`WorldCat Search`), and writing to an Excel file (`openpyxl`)