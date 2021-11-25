from time import sleep
from typing import Tuple
from io import BytesIO
import os
import argparse
import re
from operator import attrgetter
from fitz import Rect
from fitz import open as pdf_open


def extract_info(input_file: str):
    """
    Extracts file info
    """
    # Open the PDF
    pdfDoc = pdf_open(input_file)
    output = {
        "File": input_file, "Encrypted": ("True" if pdfDoc.isEncrypted else "False")
    }
    # If PDF is encrypted the file metadata cannot be extracted
    if not pdfDoc.isEncrypted:
        for key, value in pdfDoc.metadata.items():
            output[key] = value
    # To Display File Info
    print("## File Information ##################################################")
    print("\n".join("{}:{}".format(i, j) for i, j in output.items()))
    print("######################################################################")
    return True, output


def search_for_text(lines, search_str):
    """
    Search for the search string within the document lines
    """
    for line in lines:
        if line[1] == search_str:
            yield line


def redact_matching_data(page, matched_values):
    """
    Redacts matching values
    """
    matches_found = 0
    # Loop throughout matching values
    for val in matched_values:
        matches_found += 1
        matching_val_area = page.search_for(val)
        # Redact matching values
        [page.addRedactAnnot(area, text=" ", fill=(0, 0, 0))
         for area in matching_val_area]
    # Apply the redaction
    page.apply_redactions()
    return matches_found


def annotate_matching_data(page, matched_values, action):
    """
    Annotate matching values
    """
    matches_found = 0
    for val in matched_values:
        matches_found += 1
        matching_val_areas = page.search_for(val[0])
        distinct_line_y_values = set(val.y0 for val in matching_val_areas)
        for y0 in distinct_line_y_values:
            line_areas = list(
                filter(lambda val: val.y0 == y0, matching_val_areas))
            matching_val_area = combine_area(line_areas)
            annotate_area(page, matching_val_area, action)
    return matches_found  


def annotate_area(page, area, action):
    """
    Annotates area on page
    """
    if action in ('Squiggly', 'Underline'):
        area.y0 += 2
        area.y1 += 2

    if action == 'Squiggly':
        annotation = page.add_squiggly_annot(area)
    elif action == 'Underline':
        annotation = page.add_underline_annot(area)
    elif action == 'Strikeout':
        annotation = page.add_strikeout_annot(area)
    elif action == 'Frame':
        annotation = page.add_rect_annot(area)
    else:
        annotation = page.add_highlight_annot(area)

    # To change the highlight color
    # highlight.setColors({"stroke":(0,0,1),"fill":(0.75,0.8,0.95) })
    # highlight.setColors(stroke = fitz.utils.getColor('white'), fill = fitz.utils.getColor('red'))
    # highlight.setColors(colors= fitz.utils.getColor('red'))

    # annotation.update()


def combine_area(line_areas):
    x0 = min(line_areas, key=attrgetter('x0')).x0
    x1 = max(line_areas, key=attrgetter('x1')).x1
    y0 = line_areas[0].y0
    y1 = line_areas[0].y1
    return Rect(x0, y0, x1, y1)


def process_data(input_file: str, output_file: str, search_str: str, pages: Tuple = None, teams: set = None, action: str = 'Highlight'):
    """
    Process the pages of the PDF File
    """
    with pdf_open(input_file) as pdf_input_file:

        if not teams:
            teams = get_teams(pdf_input_file)
        
        teams = list(teams)
        teams.sort()

        for team in teams:
            highlight_team(input_file, team, pages, action)
    

def highlight_team(input_file: str, team_name: str, pages: Tuple = None, action: str = 'Highlight'):
    """
    Highlights file for given team
    """
    
    pdf_input_file = pdf_open(input_file)

    output_buffer = BytesIO()
    total_matches = 0

    for pg in range(pdf_input_file.pageCount):

        # If required for specific pages
        if pages:
            if str(pg) not in pages:
                continue

        page = pdf_input_file[pg]

        text = page.get_text("text")
        regexString = r"((\w{1,5})-([A-Z]{2})\W*((?:\d*:)?\d{2}.\d{2}|NT)\W*(\d+)\W*([A-Z]\w*,\W[A-Z]\w*\W[A-Z]?)\W*(\d))"
        page_lines = re.findall(regexString, text)
        matched_values = search_for_text(page_lines, team_name)

        if matched_values:
            matches_found = annotate_matching_data(page, matched_values, action)
            total_matches += matches_found

    print(f"{total_matches} Swims Highlighted for Team: {team_name} In Input File: {input_file}")

    # Save to output
    pdf_input_file.save(output_buffer)

    # Save the output buffer to the output file
    with open(f"{os.path.splitext(input_file)[0]}-{team_name}{os.path.splitext(input_file)[1]}", mode='wb') as file:
        file.write(output_buffer.getbuffer())


def get_teams(pages):
    swimmer_line_regex_string = r"((\w{1,5})-([A-Z]{2})\W*((?:\d*:)?\d{2}.\d{2}|NT)\W*(\d+)\W*([A-Z]\w*,\W[A-Z]\w*\W[A-Z]?)\W*(\d))"
    teams = set()
    for page in pages:
        text = page.get_text("text")
        swimmer_lines = re.findall(swimmer_line_regex_string, text)
        teams = teams | {line[1] for line in swimmer_lines}
    return teams


def remove_highlght(input_file: str, output_file: str, pages: Tuple = None):
    # Open the PDF
    pdfDoc = pdf_open(input_file)
    # Save the generated PDF to memory buffer
    output_buffer = BytesIO()
    # Initialize a counter for annotations
    annot_found = 0
    # Iterate through pages
    for pg in range(pdfDoc.pageCount):
        # If required for specific pages
        if pages:
            if str(pg) not in pages:
                continue
        # Select the page
        page = pdfDoc[pg]
        annot = page.firstAnnot
        while annot:
            annot_found += 1
            page.deleteAnnot(annot)
            annot = annot.next
    if annot_found >= 0:
        print(f"Annotation(s) Found In The Input File: {input_file}")
    # Save to output
    pdfDoc.save(output_buffer)
    pdfDoc.close()
    # Save the output buffer to the output file
    with open(output_file, mode='wb') as file:
        file.write(output_buffer.getbuffer())


def process_file(**kwargs):
    """
    To process one single file
    Redact, Frame, Highlight... one PDF File
    Remove Highlights from a single PDF File
    """
    input_file = kwargs.get('input_file')
    output_file = kwargs.get('output_file')
    if output_file is None:
        output_file = input_file
    search_str = kwargs.get('search_str')
    pages = kwargs.get('pages')
    # Redact, Frame, Highlight, Squiggly, Underline, Strikeout, Remove
    action = kwargs.get('action')
    if action == "Remove":
        # Remove the Highlights except Redactions
        remove_highlght(input_file=input_file,
                        output_file=output_file, pages=pages)
    else:
        process_data(input_file=input_file, output_file=output_file,
                     search_str=search_str, pages=pages, action=action)


def process_folder(**kwargs):
    """
    Redact, Frame, Highlight... all PDF Files within a specified path
    Remove Highlights from all PDF Files within a specified path
    """
    input_folder = kwargs.get('input_folder')
    search_str = kwargs.get('search_str')
    # Run in recursive mode
    recursive = kwargs.get('recursive')
    #Redact, Frame, Highlight, Squiggly, Underline, Strikeout, Remove
    action = kwargs.get('action')
    pages = kwargs.get('pages')
    # Loop though the files within the input folder.
    for foldername, dirs, filenames in os.walk(input_folder):
        for filename in filenames:
            # Check if pdf file
            if not filename.endswith('.pdf'):
                continue
            # PDF File found
            inp_pdf_file = os.path.join(foldername, filename)
            print("Processing file =", inp_pdf_file)
            process_file(input_file=inp_pdf_file, output_file=None,
                         search_str=search_str, action=action, pages=pages)
        if not recursive:
            break


def is_valid_path(path):
    """
    Validates the path inputted and checks whether it is a file path or a folder path
    """
    if not path:
        raise ValueError("Invalid Path")
    if os.path.isfile(path):
        return path
    elif os.path.isdir(path):
        return path
    else:
        raise ValueError(f"Invalid Path {path}")


def parse_args():
    """Get user command line parameters"""
    parser = argparse.ArgumentParser(description="Available Options")
    parser.add_argument('-i', '--input_path', dest='input_path', type=is_valid_path,
                        required=True, help="Enter the path of the file or the folder to process")
    parser.add_argument('-a', '--action', dest='action', choices=['Redact', 'Frame', 'Highlight', 'Squiggly', 'Underline', 'Strikeout', 'Remove'], type=str,
                        default='Highlight', help="Choose whether to Redact or to Frame or to Highlight or to Squiggly or to Underline or to Strikeout or to Remove")
    parser.add_argument('-p', '--pages', dest='pages', type=tuple,
                        help="Enter the pages to consider e.g.: [2,4]")
    action = parser.parse_known_args()[0].action
    # if action != 'Remove':
    #     parser.add_argument('-s', '--search_str', dest='search_str'                            # lambda x: os.path.has_valid_dir_syntax(x)
    #                         , type=str, required=True, help="Enter a valid search string")
    path = parser.parse_known_args()[0].input_path
    if os.path.isfile(path):
        parser.add_argument('-o', '--output_file', dest='output_file', type=str  # lambda x: os.path.has_valid_dir_syntax(x)
                            , help="Enter a valid output file")
    if os.path.isdir(path):
        parser.add_argument('-r', '--recursive', dest='recursive', default=False, type=lambda x: (
            str(x).lower() in ['true', '1', 'yes']), help="Process Recursively or Non-Recursively")
    args = vars(parser.parse_args())
    # To Display The Command Line Arguments
    print("## Command Arguments #################################################")
    print("\n".join("{}:{}".format(i, j) for i, j in args.items()))
    print("######################################################################")
    return args


if __name__ == '__main__':
    # Parsing command line arguments entered by user
    args = parse_args()
    # If File Path
    if os.path.isfile(args['input_path']):
        # Extracting File Info
        extract_info(input_file=args['input_path'])
        # Process a file
        process_file(
            input_file=args['input_path'], output_file=args['output_file'],
            search_str=args['search_str'] if 'search_str' in (
                args.keys()) else None,
            pages=args['pages'], action=args['action']
        )
    # If Folder Path
    elif os.path.isdir(args['input_path']):
        # Process a folder
        process_folder(
            input_folder=args['input_path'],
            search_str=args['search_str'] if 'search_str' in (
                args.keys()) else None,
            action=args['action'], pages=args['pages'], recursive=args['recursive']
        )