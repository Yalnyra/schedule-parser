import openpyxl
import win32com.client
from docx_parser import *
import pandas as pd
import json
import re
import os

# It's easier to manipulate docx tables than binaries, so lets convert .doc to modern .docx
def convert_doc_to_docx(input_doc, output_docx):
    try:
        # If the .docx already exists, stop
        word = win32com.client.Dispatch("Word.Application")
        docx = word.Documents.Open(output_docx)
        docx.Close()
        word.Quit()
        return
    except Exception as e:
        try:
            doc = word.Documents.Open(input_doc)
            doc.SaveAs(output_docx, 16)  # 16 is the value for .docx format
            doc.Close()
            word.Quit()
            print(f"Converted {input_doc} to {output_docx}")
        except Exception as e:
            print(f"Conversion from .doc failed: {e}")

# Make it easier to manipulate data inside the table by turning it into a DataFrame
def extract_table_data_docx(doc_parser):
    data = []
    header = []
    # Scrape the table from the doc file
    for table in doc_parser.document.tables:
        # the first row must be the header of the DataFrame
        header_row = True
        # last cell of the previous row
        previous_cell = None
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                if previous_cell:
                    row_data.append(previous_cell)
                cell_data = cell.text.strip()
                row_data.append(cell_data)
            # Remove the last cell contents from the row and pad it to the next row
            if (len(row_data) != 0):
                previous_cell = row_data.pop()
            if not header_row:
                data.append(row_data)
                continue
            header = row_data
            header_row = False

    # Create a dataframe to manipulate the table
    df = pd.DataFrame(columns=header)
    # Delete duplicate columns from merged cells
    df = df.loc[:, ~df.columns.duplicated()]
    # Try appending the next row to the dataframe,
    # If it already fits the table header
    # todo fit all the rows to their respective columns
    for row in data:
        if len(row) != len(header):
            pass
        else:
            df = df.add(dict(zip(df.columns, row)))

    return df

# Map the data inside the document into a structured json
# TODO Map the xslx doc table onto the unified json structure
def shape_data(docx_file):
    # parse the document
    doc = DocumentParser(docx_file)
    doc_str = str(list(doc.parse()))
    faculty = re.findall(r"Факультет[ а-яА-ЯіїєґІЇЄҐ]+", doc_str)
    # If the doc doesn't specify the faculty, consider it empty/corrupted and return an empty dict
    if faculty is None:
        return {}
    faculty = faculty[0]
    # the schedule dictionary
    faculty_schedule = {faculty: {}}

    # Find all the specialities that are declared in the doc
    # Relevant for the economic faculty
    spec_matches = [(match.group(0), match.end()) for match in re.finditer(r'[«"]([а-яА-ЯіїєґІЇЄҐ`, ]*?)[»"]',
                                                                           doc_str,
                                                                           flags=re.MULTILINE)]

    specs = []
    year_span = 0
    year = 0
    # Drop the parentheses from the specialities and add them to the faculty
    for match, end in spec_matches:
        year_span = end
        spec = re.sub(r'[«"»]', "", match)
        faculty_schedule[faculty][spec] = {}
        specs.append(spec)
    # Define for which year the subjects are
    # If the next char starts with M (Impliying "МП"), then it's the masters year
    for s in doc_str[year_span:]:
        if s == 'М':
            year += 4
        # The next decimal is the year
        if s.isdecimal():
            year += int(s)
            break
    # Add the year to the specialities inside the faculty
    for spec in specs:
        faculty_schedule[faculty][spec][f'{year} рік навчання'] = {}
    # Create the schedule table dataframe
    schedule_df = extract_table_data_docx(doc)
    # todo filter the actual table for subjects, their specialities, weeks, day and time, groups, etc.
    # todo add the table data into the dictionary
    return faculty_schedule

# TODO Make parser method for the excel doc
def extract_table_data_xlsx(pyxl_parser):
    pass

# Dump the schedule dictionary into a serialized json
def convert_to_json(data, output):
    with open(output, 'w') as json_file:
        json.dump(data, json_file, indent=4, ensure_ascii=False)


if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # Relative paths to input and output files
    # Replace with the path to FEN .doc schedule file
    input_docx = os.path.join(script_dir, '3.doc')
    conv_docx = os.path.join(script_dir, '3.docx')
    # Replace with the desired output JSON file
    output_json = os.path.join(script_dir, 'schedule_data.json')
    convert_doc_to_docx(input_docx, conv_docx)
    schedule_dict = shape_data(conv_docx)
    convert_to_json(schedule_dict, output_json)
