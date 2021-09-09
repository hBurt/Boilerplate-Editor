import glob
import os
import re
import subprocess

from datetime import date
from docx import Document
import tkinter as tk

# constants

# gui screen start position assumes monitors are setup in parallel(left to right) and the primary monitor is on the left
# gui screen start position(1 = middle of first screen, 4 = middle of second screen)
gui_start_position = 1

# this truncates n number of lines from the output when copying to the clipboard(does not effect pdf output)
truncate_lines_before = 5
truncate_lines_after = 6

converted_file_base_name = "HarrisonStewardCoverLetter_"

docx_template_path = "CoverLetter_Template.docx"
libre_office_swriter_path = "C:\\Program Files\\LibreOffice\\program\\swriter.exe"
pdf_output_directory = "C:\\Users\\Harrison\\Documents\\"

default_curr_date = re.compile("zCurrentDatez")
default_job_title = re.compile("zJobTitlez")
default_company = re.compile("zCompanyz")
default_city_state = re.compile("zCityStatez")


# https://github.com/python-openxml/python-docx/issues/30#issuecomment-879593691
def paragraph_replace_text(paragraph, regex, replace_str):
    """Return `paragraph` after replacing all matches for `regex` with `replace_str`.

    `regex` is a compiled regular expression prepared with `re.compile(pattern)`
    according to the Python library documentation for the `re` module.
    """
    # --- a paragraph may contain more than one match, loop until all are replaced ---
    while True:
        text = paragraph.text
        match = regex.search(text)
        if not match:
            break

        # --- when there's a match, we need to modify run.text for each run that
        # --- contains any part of the match-string.
        runs = iter(paragraph.runs)
        start, end = match.start(), match.end()

        # --- Skip over any leading runs that do not contain the match ---
        for run in runs:
            run_len = len(run.text)
            if start < run_len:
                break
            start, end = start - run_len, end - run_len

        # --- Match starts somewhere in the current run. Replace match-str prefix
        # --- occurring in this run with entire replacement str.
        run_text = run.text
        run_len = len(run_text)
        run.text = "%s%s%s" % (run_text[:start], replace_str, run_text[end:])
        end -= run_len  # --- note this is run-len before replacement ---

        # --- Remove any suffix of match word that occurs in following runs. Note that
        # --- such a suffix will always begin at the first character of the run. Also
        # --- note a suffix can span one or more entire following runs.
        for run in runs:  # --- next and remaining runs, uses same iterator ---
            if end <= 0:
                break
            run_text = run.text
            run_len = len(run_text)
            run.text = run_text[end:]
            end -= run_len

    # --- optionally get rid of any "spanned" runs that are now empty. This
    # --- could potentially delete things like inline pictures, so use your judgement.
    # for run in paragraph.runs:
    #     if run.text == "":
    #         r = run._r
    #         r.getparent().remove(r)

    return paragraph


def convert_call_back(job_title, company_name, city_state, is_remote, is_copy_txt_to_clipboard):
    file_name = converted_file_base_name + company_name + ".docx"
    document = Document(docx=docx_template_path)

    today = date.today()

    # replace default words(city/state, jobTitle, companyName)
    for p in document.paragraphs:

        paragraph_replace_text(p, default_curr_date, today.strftime("%B %d, %Y"))
        if is_remote == 1:
            paragraph_replace_text(p, default_city_state, city_state + "\nRemote")
        else:
            paragraph_replace_text(p, default_city_state, city_state)
        paragraph_replace_text(p, default_company, company_name)
        paragraph_replace_text(p, default_job_title, job_title)

    document.save(file_name)

    if is_copy_txt_to_clipboard == 1:
        # create the txt file from our *.docx file
        with open("output.txt", "w") as text_file:
            for p in document.paragraphs:
                tmp_stripped_string = re.sub(r'\t', '', p.text)
                text_file.write(tmp_stripped_string)
                text_file.write('\n')

        # read the text file to prepare for truncation
        with open('output.txt', 'r') as fin:
            data = fin.read().splitlines(True)

        # write the data back to the text file
        with open('output.txt', 'w') as fout:
            fout.writelines(data[truncate_lines_before:-truncate_lines_after])

        # copy text file to clipboard
        os.system("clip < output.txt ")

        # delete text file
        os.remove("output.txt")

    else:
        subprocess.run(
            [libre_office_swriter_path, "--headless", "--convert-to", "pdf",
             file_name, "--outdir", pdf_output_directory])

    os.remove(file_name)


# Start GUI
window = tk.Tk()
window.iconphoto(False, tk.PhotoImage(file='scroll_x64.png'))
window.title("Boilerplate Editor")
window.geometry("400x225")

var1 = tk.IntVar()
var1.set(1)

var2 = tk.IntVar()
var2.set(0)

label_job_title = tk.Label(window, text="Job Title: ")
label_company_name = tk.Label(window, text="Company Name: ")
label_city_state = tk.Label(window, text="City/State: ")

entry_job_title = tk.Entry(window, width=40)
entry_company_name = tk.Entry(window, width=40)
entry_city_state = tk.Entry(window, width=40)

check_is_remote = tk.Checkbutton(window, text="Is Remote? ", variable=var1)
check_is_txt_file = tk.Checkbutton(window, text="Copy as text to clipboard", variable=var2)

button_convert = tk.Button(window, text="Convert",
                           command=lambda: convert_call_back(entry_job_title.get(), entry_company_name.get(),
                                                             entry_city_state.get(), var1.get(), var2.get()))

label_job_title.grid(row=0, column=0, pady=2, padx=20)
label_company_name.grid(row=1, column=0, pady=2, padx=20)
label_city_state.grid(row=2, column=0, pady=2, padx=20)

entry_job_title.grid(row=0, column=1, pady=10)
entry_company_name.grid(row=1, column=1, pady=10)
entry_city_state.grid(row=2, column=1, pady=10)

check_is_remote.grid(row=4, column=1, sticky=tk.W)
check_is_txt_file.grid(row=5, column=1, sticky=tk.W)

button_convert.grid(row=6, column=0, padx=(20, 0), columnspan=2, sticky=tk.N + tk.S + tk.W + tk.E)


def clear_values():
    entry_job_title.delete(0, tk.END)
    entry_company_name.delete(0, tk.END)
    entry_city_state.delete(0, tk.END)

    entry_job_title.focus_set()


def delete_excess_entries():
    for filename in glob.glob("C:/Users/Harrison/Documents/*"):
        if filename.__contains__(converted_file_base_name) and not filename.__contains__('Updated_Template'):
            print(os.path.basename(filename))
            os.remove(filename)


# Start menubar
menubar = tk.Menu(window)
file_menu = tk.Menu(menubar, tearoff=0)
file_menu.add_command(label="Clear Values", command=clear_values)
file_menu.add_command(label="Delete excess entries", command=delete_excess_entries)
menubar.add_cascade(label="File", menu=file_menu)

window.config(menu=menubar)

# Same size will be defined in variable for center screen in Tk_Width and Tk_height
tk_width = 400
tk_height = 225

# calculate coordination of screen and window form
x_left = int((window.winfo_screenwidth() / 2 - tk_width / 2) * gui_start_position)
y_top = int(window.winfo_screenheight() / 2 - tk_height / 2)

# Write following format for center screen
window.geometry("+{}+{}".format(x_left, y_top))

window.mainloop()
