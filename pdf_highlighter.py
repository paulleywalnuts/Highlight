from typing import Tuple
from io import BytesIO
import os
import argparse
import re
from operator import attrgetter
from fitz import Rect
from fitz import open as pdf_open
from tkinter import filedialog
from termcolor import colored
from ctypes import windll
windll.shcore.SetProcessDpiAwareness(1)


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

    # Adjust the selection vertically to not overlap text
    if action in ('Squiggly', 'Underline'):
        area.y0 += 2
        area.y1 += 2

    # Apply annotation
    if action == 'Squiggly':
        page.add_squiggly_annot(area)
    elif action == 'Underline':
        page.add_underline_annot(area)
    elif action == 'Strikeout':
        page.add_strikeout_annot(area)
    elif action == 'Frame':
        page.add_rect_annot(area)
    else:
        page.add_highlight_annot(area)

    # To change the highlight color
    # highlight.setColors({"stroke":(0,0,1),"fill":(0.75,0.8,0.95) })
    # highlight.setColors(stroke = fitz.utils.getColor('white'), fill = fitz.utils.getColor('red'))
    # highlight.setColors(colors= fitz.utils.getColor('red'))


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


def combine_area(line_areas):
    """
    Combine highlight rectangles into single object
    """
    x0 = min(line_areas, key=attrgetter('x0')).x0
    x1 = max(line_areas, key=attrgetter('x1')).x1
    y0 = line_areas[0].y0
    y1 = line_areas[0].y1
    return Rect(x0, y0, x1, y1)


def process_data(input_file: str, teams: set = None, pages: Tuple = None, action: str = 'Highlight'):
    """
    Process the pages of the PDF File
    """

    if input_file is None:
        input_file = filedialog.askopenfilename()

    if input_file is '':
        exit()

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
        matched_values = [line for line in page_lines if line[1] == team_name]

        if matched_values:
            matches_found = annotate_matching_data(page, matched_values, action)
            total_matches += matches_found

    print(f"{total_matches:3} Swims Highlighted for Team:", colored(f"{team_name}", "green"))

    # Save to output
    pdf_input_file.save(output_buffer)

    # Save the output buffer to the output file
    with open(f"{os.path.splitext(input_file)[0]}-{team_name}{os.path.splitext(input_file)[1]}", mode='wb') as file:
        file.write(output_buffer.getbuffer())


def get_teams(pages):
    """
    Pull all team code values from the heat sheet document
    """
    swimmer_line_regex_string = r"((\w{1,5})-([A-Z]{2})\W*((?:\d*:)?\d{2}.\d{2}|NT)\W*(\d+)\W*([A-Z]\w*,\W[A-Z]\w*\W[A-Z]?)\W*(\d))"
    teams = set()
    for page in pages:
        text = page.get_text("text")
        swimmer_lines = re.findall(swimmer_line_regex_string, text)
        teams = teams | {line[1] for line in swimmer_lines}
    return teams


def remove_highlight(input_file: str, output_file: str, pages: Tuple = None):
    """
    Remove all non-redaction highlights from file
    """
    
    # Initialize files in memory
    pdfDoc = pdf_open(input_file)
    output_buffer = BytesIO()
    annot_found = 0


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
    input_file = kwargs.get('input_path')
    teams = kwargs.get('teams')
    pages = kwargs.get('pages')
    action = kwargs.get('action')
    if action == "Remove":
        # Remove the Highlights except Redactions
        remove_highlight(input_file=input_file, pages=pages)
    else:
        process_data(input_file=input_file, teams=teams,
                     pages=pages, action=action)


def is_valid_path(path):
    """
    Validates the path inputted and checks whether it is a file path or a folder path
    """
    if not path:
        raise ValueError("No path given")
    if os.path.isfile(path):
        return path
    elif os.path.isdir(path):
        return path
    else:
        raise ValueError(f"Invalid Path: {path}")


def parse_args():
    """Get user command line parameters"""
    
    parser = argparse.ArgumentParser(description="Available Options")

    parser.add_argument('-i', '--input_path', dest='input_path', type=is_valid_path,
                        help="Enter the path of the file or the folder to process")
    parser.add_argument('-a', '--action', dest='action', choices=['Redact', 'Frame', 'Highlight', 'Underline', 'Squiggly', 'Strikeout', 'Remove'], type=str,
                        default='Highlight', help="Choose whether to Redact, Frame, Highlight, Underline, Squiggly Underline, Strikeout or Remove")
    parser.add_argument('-p', '--pages', dest='pages', type=tuple,
                        help="Enter the pages to consider e.g.: [2,4]")
    parser.add_argument('-o', '--output_file', dest='output_file', type=str
                            , help="Enter a valid output file")
    args = vars(parser.parse_args())

    # Display The Command Line Arguments
    print("## Command Arguments #################################################")
    print("\n".join("{}:{}".format(i, j) for i, j in args.items()))
    print("######################################################################")

    return args


if __name__ == '__main__':
    
    process_file(**parse_args())