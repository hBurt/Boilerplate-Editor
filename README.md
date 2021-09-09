# Boilerplate-Editor
Create PDF's and plain text from GUI input for job applications

**This project makes use of pythons subprocess module to run windows commands. As the project stands right now, it will not work on macOS or Linux.**

## Before Running
This python file makes use of Libre Office and its API to convert .docx file into pdf files. [Libre office can be found here](https://www.libreoffice.org/) and must be installed prior to use of this application.

The top portion of boilerplateGUI.pyw has a few constat variables which *must be edited* according to your computer:
* docx_template_path - this where the default template is located which is converted to pdf/txt
* libre_office_swriter_path - this is the path of swriter.exe which is installed with libre office
* pdf_output_directory - this is the output directory of your converted pdf files

Optional variables to take into consideration:
* converted_file_base_name - this is the base name of the file before the company name is added(eg. FirstNameLastNameCoverLetter_Microsoft)
* gui_start_position - this variable will change which screen the GUI will first appear on
* truncate_lines_before - remove n number of lines on top of txt output
* truncate_lines_after - remove n number of lines on bottom of txt output

Included in the project is a .bat file which will run the GUI.
To use this bat file you need to edit it to account for the installed location of python on your machine and the location of the boilerplateGUI.pyw file. 

**You must leave this bat file in the same folder as the boilerplateGUI.pyw file.** If you want to have this bat file elsewhere, right click on it and create a shortcut. From there you can then place that shortcut anywhere on your computer.

## Media
![Editor](https://i.imgur.com/F6d0kcE.png)






