from typing import Tuple
from io import BytesIO
import os
import argparse
from regex import findall, fullmatch
from operator import attrgetter
from fitz import Rect
from fitz import open as pdf_open
from tkinter import filedialog
from termcolor import colored
from ctypes import windll
windll.shcore.SetProcessDpiAwareness(1)


class HeatSheet:

    def __init__(self, file_name) -> None:
        self.file_name = None
        self.open(file_name)

    @property
    def teams(self):
        """All team codes in the heat sheat."""

        if self.__teams:
            return self.__teams

        teams = {swim.team for swim in self.individual_swims}
        teams = teams | {swim.team for swim in self.relay_swims}

        teams = [team.code for team in teams]
        teams.sort()

        self.__teams = teams
        return self.__teams

    @property
    def cuts(self):
        """All cut codes in the heat sheat."""

        if self.__cuts:
            return self.__cuts

        cuts = set()

        for page in self.pdf_data:
            text = page.get_text("text")
            cuts = cuts | HeatSheet.Cut.findall(text)

        cuts = list(cuts)
        cuts.sort()

        self.__cuts = cuts
        return self.__cuts

    @property
    def individual_swims(self):
        """All individual swims in the heat sheet."""

        if self.__ind:
            return self.__ind

        individual_swims = list()

        for page in self.pdf_data:
            text = page.get_text("text")
            individual_swims = individual_swims + \
                HeatSheet.IndividualSwim.findall(text, self.cuts)

        self.__ind = individual_swims
        return self.__ind

    @property
    def relay_swims(self):
        """All relay swims in the heat sheet."""

        if self.__rel:
            return self.__rel

        relay_swims = list()

        for page in self.pdf_data:
            text = page.get_text("text")
            relay_swims = relay_swims + \
                HeatSheet.RelaySwim.findall(text, self.cuts)

        self.__rel = relay_swims
        return self.__rel

    def open(self, file_name) -> None:
        if self.file_name != file_name:
            self.file_name = file_name
            self.__teams = None
            self.__cuts = None
            self.__ind = None
            self.__rel = None
        self.pdf_data = pdf_open(file_name)

    def save_as(self, save_path) -> None:
        output_buffer = BytesIO()
        self.pdf_data.save(output_buffer)
        with open(save_path, mode='wb') as file:
            file.write(output_buffer.getbuffer())

    def highlight_team(self, team_name: str, pages: Tuple = None, action: str = 'Highlight') -> None:
        total_matches = 0

        for pg in range(self.pdf_data.pageCount):

            # If required for specific pages
            if pages:
                if str(pg) not in pages:
                    continue

            page = self.pdf_data[pg]

            text = page.get_text("text")

            individual_swims = HeatSheet.IndividualSwim.findall(
                text, self.cuts)
            relay_swims = HeatSheet.RelaySwim.findall(text, self.cuts)

            matched_swims = [f"{swim}" for swim in relay_swims +
                             individual_swims if swim.team.code == team_name]

            if matched_swims:
                matches_found = self.__annotate_matching_data(
                    page, matched_swims, action)
                total_matches += matches_found

        print(f"{total_matches:3} Swims Highlighted for Team:",
              colored(f"{team_name}", "green"))

        # Create output directory if it does not exist
        output_directory = f"{os.path.dirname(self.file_name)}/Highlighted/{team_name}"
        if not os.path.isdir(output_directory):
            os.makedirs(output_directory)

        # Create output file path
        output_filename = f"{os.path.basename(os.path.splitext(self.file_name)[0])} - {team_name}{os.path.splitext(self.file_name)[1]}"
        output_path = output_directory + "/" + output_filename

        # Save and reload the unhighlighted file
        self.save_as(output_path)
        self.open(self.file_name)

    def __annotate_matching_data(self, page, matched_values, action):
        """
        Annotate matching values
        """
        matches_found = 0
        for val in matched_values:
            matches_found += 1
            matching_val_areas = page.search_for(val)
            distinct_line_y_values = set(round(val.y0, 3) for val in matching_val_areas)
            for y0 in distinct_line_y_values:
                line_areas = list(
                    filter(lambda val: round(val.y0, 3) == y0, matching_val_areas))
                matching_val_area = self.__combine_area(line_areas)
                self.__annotate_area(page, matching_val_area, action)
        return matches_found

    def __combine_area(self, line_areas):
        """
        Combine highlight rectangles into single object
        """
        x0 = min(line_areas, key=attrgetter('x0')).x0
        x1 = max(line_areas, key=attrgetter('x1')).x1
        y0 = line_areas[0].y0
        y1 = line_areas[0].y1
        return Rect(x0, y0, x1, y1)

    def __annotate_area(self, page, area, action):
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

    class IndividualSwim():

        age = r" \d+"
        name = r"[A-Z]\w*,\s[A-Z]\w*\s[A-Z]?"
        time = r"(?:(?:(?:\d*:)?\d{2}\.\d{2})|NT)"
        team = r"\w{1,5}(?=-)"
        lsc = r"(?<=-)[A-Z]{2}"
        team_code = f"{team}-{lsc}"
        lane = f"(?<={name}\n)[1-9]"

        individual_swim = f"{team_code}\n{time}\n{age}\n{name}\n{lane}\n"

        def __init__(self, team, swimmer, age, time, lane, cut_code) -> None:
            self.team = team
            self.swimmer = swimmer
            self.age = age
            self.time = time
            self.lane = lane
            self.cut_code = cut_code

        def __str__(self) -> str:
            return f"{self.team}\n{self.time}\n{self.age}\n{self.swimmer}\n{self.lane}\n" + (f"{self.cut_code}\n" if self.cut_code else "")

        @classmethod
        def from_string(cls, string, cuts):

            individual_swim = cls.individual_swim
            cut_codes = ""
            if cuts:
                cut_codes = "|".join(cuts)
                individual_swim = f"{individual_swim}(?:{cut_codes})?\n?"

            if not fullmatch(individual_swim, string):
                raise ValueError("String not formatted properly")

            team = HeatSheet.Team.from_string(
                findall(cls.team_code, string)[0])
            swimmer = findall(cls.name, string)[0]
            age = findall(cls.age, string)[0]
            time = findall(cls.time, string)[0]
            lane = findall(cls.lane, string)[0]
            cut_code = findall(cut_codes, string)
            if cut_code:
                cut_code = cut_code[0]

            return cls(team, swimmer, age, time, lane, cut_code)

        @classmethod
        def findall(cls, text, cuts):

            individual_swim = cls.individual_swim
            if cuts:
                cut_codes = "|".join(cuts)
                individual_swim = f"{individual_swim}(?:{cut_codes})?\n?"

            return [HeatSheet.IndividualSwim.from_string(swim, cuts) for swim in findall(individual_swim, text)]

    class RelaySwim():

        relay_letter = r"[A-Z]"
        time = r"(?:(?:(?:\d*:)?\d{2}\.\d{2})|NT)"
        team = r"\w{1,5}(?=-)"
        lsc = r"(?<=-)[A-Z]{2}"
        team_code = f"{team}-{lsc}"
        lane = f"(?<={team_code}\n)[1-9]"

        relay_swim = f"{relay_letter}\n{time}\n{team_code}\n{lane}\n"

        def __init__(self, team, letter, time, lane, cut_code) -> None:
            self.team = team
            self.letter = letter
            self.time = time
            self.lane = lane
            self.cut_code = cut_code

        def __str__(self) -> str:
            return f"{self.letter}\n{self.time}\n{self.team}\n{self.lane}\n" + (f"{self.cut_code}\n" if self.cut_code else "")

        @classmethod
        def from_string(cls, string, cuts):

            relay_swim = cls.relay_swim
            if cuts:
                cut_codes = "|".join(cuts)
                relay_swim = f"{relay_swim}(?:{cut_codes})?\n?"

            if not fullmatch(relay_swim, string):
                raise ValueError("String not formatted properly")

            team = HeatSheet.Team.from_string(
                findall(cls.team_code, string)[0])
            letter = findall(cls.relay_letter, string)[0]
            time = findall(cls.time, string)[0]
            lane = findall(cls.lane, string)[0]
            cut_code = findall(cut_codes, string)
            if cut_code:
                cut_code = cut_code[0]

            return cls(team, letter, time, lane, cut_code)

        @classmethod
        def findall(cls, text, cuts):

            relay_swim = cls.relay_swim
            if cuts:
                cut_codes = "|".join(cuts)
                relay_swim = f"{relay_swim}(?:{cut_codes})?\n?"

            return [HeatSheet.RelaySwim.from_string(swim, cuts) for swim in findall(relay_swim, text)]

    class Team():

        def __init__(self, code, lsc) -> None:
            self.code = code
            self.lsc = lsc

        def __str__(self) -> str:
            return f"{self.code}-{self.lsc}"

        def __repr__(self) -> str:
            return f"{self}"

        def __eq__(self, other) -> bool:
            return self.code == other.code and self.lsc == other.lsc

        def __hash__(self) -> int:
            return hash(repr(self))

        @classmethod
        def from_string(cls, string):

            team = r"\w{1,5}(?=-)"
            lsc = r"(?<=-)[A-Z]{2}"

            team_code = f"{team}-{lsc}"

            if not fullmatch(team_code, string):
                raise ValueError("String not formatted properly")

            code = findall(team, string)[0]
            lsc = findall(lsc, string)[0]
            return cls(code, lsc)

    class Cut():

        event_label = r"#\d+.*"
        event_sponsor = r"Sponsor:.*"
        cut_code = r"[A-Z]+"
        cut_code_description = r"\s.*"
        time = r"(?:(?:(?:\d*:)?\d{2}\.\d{2})|NT)"

        event_header = f"{event_label}\n(?:{event_sponsor}\n)?(?:{cut_code}(?:{cut_code_description})?\n(?:{time})\n)*"
        cut = f"{cut_code}(?:{cut_code_description})?\n{time}\n"

        def __init__(self, code, time_string) -> None:
            self.code = code
            self.time_string = time_string

        @classmethod
        def from_string(cls, string):

            if not fullmatch(cls.cut, string):
                raise ValueError("String not formatted properly")

            code = findall(cls.cut_code, string)[0]
            time_string = findall(cls.time, string)[0]
            return cls(code, time_string)

        @classmethod
        def findall(cls, text):

            cuts = set()

            event_headers = findall(cls.event_header, text)
            for event_header in event_headers:
                cut_times = findall(cls.cut, event_header)
                cuts = cuts | {HeatSheet.Cut.from_string(
                    cut_time).code for cut_time in cut_times}

            return cuts


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
    parser.add_argument('-o', '--output_file', dest='output_file',
                        type=str, help="Enter a valid output file")
    args = vars(parser.parse_args())

    # Display The Command Line Arguments
    print("## Command Arguments #################################################")
    print("\n".join("{}:{}".format(i, j) for i, j in args.items()))
    print("######################################################################")

    return args


def main():

    args = parse_args()

    input_files = args["input_path"]
    pages = args["pages"]
    action = args["action"]
    teams = None

    if input_files is None:
        input_files = filedialog.askopenfilenames()

    if input_files == '':
        exit()

    for input_file in input_files:

        print(f"For File:", colored(
            f"{os.path.basename(os.path.splitext(input_file)[0])}", "yellow"))

        heat_sheet = HeatSheet(input_file)

        if not teams:
            teams = heat_sheet.teams

        for team in teams:
            heat_sheet.highlight_team(team, pages, action)


if __name__ == '__main__':
    main()
