"""Prints rosters"""
import os
import sys
import logging
import time
from datetime import datetime
from tempfile import TemporaryDirectory
import yaml
import pandas as pd
from fpdf import FPDF
import dateparser

logger = logging.getLogger("roster-printer")

# debug flags
DEBUG = os.getenv("ROSTER_PRINTER_DEBUG", "False") == "True"
PRINT_ROSTERS = not DEBUG
USE_TEMPDIR = not DEBUG
LOG_LEVEL = logging.DEBUG if DEBUG else logging.INFO

class RosterPDF(FPDF):
    """Adds footer to base class"""
    def __init__(self, footer_str="", **kwargs):
        super().__init__(**kwargs)
        self.footer_str = footer_str

    def footer(self):
        # Position cursor at 1.5 cm from bottom:
        self.set_y(-15)
        self.set_font("helvetica", style="I", size=8)
        self.cell(0, 10, self.footer_str, align="R")

def find_latest_spreadsheet(search_dir: str, search_str: str) -> os.PathLike:
    """
    returns the newest file in `search_dir` that contains `search_str`
    """
    newest_file = ""

    files = os.listdir(search_dir)
    # matching_files = [file for file in files if file.contains(search_str)]

    for filename in files:
        if search_str not in filename:
            logger.debug(f"{search_str} not in basename of {filename}")
            continue

        full_path = os.path.join(search_dir, filename)

        if not newest_file or \
                os.path.getmtime(full_path) > os.path.getmtime(newest_file):
            newest_file = full_path
            logger.debug(f"new latest files: {newest_file=}")

    logger.debug(f"{newest_file=}")
    return newest_file


def check_for_required_config(config_to_check: dict) -> None:
    """Checks that all required keys in config are present"""

    required_keys = [
        "spreadsheet-pattern",
        "columns-to-print",
        "class-column-name",
        "search-dir",
    ]

    for key in required_keys:
        if key not in config_to_check.keys():
            raise KeyError(f"Required config `{key}` not found, check that file "
                           "at CONFIG_FILE contains key")

    logger.debug("config has all required keys")

def roster_to_pdf(roster: pd.DataFrame, file_path, title, **kwargs) -> None:
    """Creates a nicly formatted pdf of the `roster` at `file_path`"""

    modified_datetime = datetime.fromtimestamp(kwargs.get("spreadsheet_mtime")).strftime("%m/%d/%y, %H:%M:%S")
    print_datetime = kwargs.get("date_printed").strftime("%m/%d/%y, %H:%M:%S")
    session_date_str = kwargs.get("session_date_str", "")
    footer_strs = []
    if config.get("show-print-date", ""):
        footer_strs.append(f"Print time: {print_datetime}")
    if config.get("show-modified-time", ""):
        footer_strs.append(f"Data last modified: {modified_datetime}")
    footer_str = ", ".join(footer_strs)

    normal_cols = [x for x in config['columns-to-print'] if x not in config.get('use-extra-row', [])]

    # modified example from https://py-pdf.github.io/fpdf2/Maths.html#using-pandas
    pdf = RosterPDF(orientation=config.get("orientation", "P"),
               format="Letter",
               unit="pt",
               footer_str=footer_str)
    pdf.set_title(title)
    pdf.add_page()

    # create header
    pdf.set_font('helvetica', size=24)
    pdf.cell(text=title, new_y="NEXT", align="C", center=True)
    if session_date_str:
        pdf.cell(text=session_date_str, new_y="NEXT", align="C", center=True)
    pdf.ln(20)

    # add table
    pdf.set_font('helvetica', size=12)
    with pdf.table(
        borders_layout="MINIMAL",
        # cell_fill_color=200,
        # cell_fill_mode="ROWS", # this doesn't work when I want two rows with the same color
        text_align="CENTER",
    ) as table:
        # get current style for the table
        fontface = pdf.font_face()

        # fpdf table expects the header to be in the first row
        # in an iterable. Dataframes store them seperately, so we
        # must combine them for this to work nicely
        for n, data_row in enumerate([normal_cols] + roster.values.tolist()):
            row = table.row()
            # this works
            # so I need to enumerate it and count.
            if n % 2 == 0:
                fontface.fill_color = 200 # shaded gray
            else:
                fontface.fill_color = 255 # white

            for n, datum in enumerate(data_row):
                if n >= len(normal_cols):
                    # FIXME: keep correct shading.
                    if not pd.isna(datum):
                        extra_row = table.row(style=row.style)

                        extra_row.cell(datum, colspan=len(normal_cols),style=fontface, )
                elif pd.isna(datum):
                    row.cell("")
                else:
                    row.cell(datum)

    pdf.output(file_path)
    logger.debug(f"created pdf {file_path}")


def print_roster(roster: pd.DataFrame, title: str, tempdir: os.PathLike, **kwargs) -> None:
    """Prints a `roster` with the title: `title`"""

    tmp_file_path = os.path.join(tempdir, f"{title}.pdf")
    roster_to_pdf(roster, tmp_file_path, title=title, **kwargs)

    if PRINT_ROSTERS is True:
        logger.info(f"printing {title}.pdf")
        os.startfile(tmp_file_path, "print")
    else:
        logger.info(f"opening {title}.pdf")
        os.startfile(tmp_file_path, "open")


def print_all_sessions(roster: pd.DataFrame, **kwargs) -> None:
    """Prints all unique sessions in roster, max one per page"""

    # to use the os print utils, the item to print must be a file
    # using TemporaryDirectory() is nicer for cleanup
    # but hard to debug, hence the flag
    # and because of the context manager, more code has to be duplicated
    # FIXME: fix the duplication. variable as function?
    if USE_TEMPDIR is True:
        with TemporaryDirectory() as tempdir:
            for session in pd.unique(roster[config["class-column-name"]].values):
                logger.debug(f"printing session of name: {session}")
                # FIXME: is this the best way to query a variable col name?
                session_df = roster.query(f"{config['class-column-name']} == @session")[config["columns-to-print"]]
                logger.debug(f"{session_df}")
                print_roster(session_df, title=f"{session} {config['title-suffix']}", tempdir=tempdir, **kwargs)
            # Without this wait, the files get deleted before the print spooler gets them
            logging.info("waiting for print spooler to receive files")
            time.sleep(10)

    else:
        tempdir = ".temp"
        if not os.path.exists(tempdir):
            os.mkdir(tempdir)
        for session in pd.unique(roster[config["class-column-name"]].values):
            logger.debug(f"printing session of name: {session}")
            # FIXME: is this the best way to query a variable col name?
            session_df = roster.query(f"{config['class-column-name']} == @session")[config["columns-to-print"]]
            logger.debug(f"{session_df}")
            print_roster(session_df, title=f"{session} {config['title-suffix']}", tempdir=tempdir, **kwargs)
        logging.info("All files should be opened now, waiting 30s before exiting")
        time.sleep(30)

if __name__ == "__main__":
    # configure logging
    logging.basicConfig(stream=sys.stdout, level=LOG_LEVEL)

    # get config, accessed globally
    CONFIG_FILE = os.getenv("CONFIG_FILE", "./config.yaml")
    logger.debug(f"{CONFIG_FILE=}")
    with open(CONFIG_FILE, encoding="utf-8") as f:
        config = yaml.safe_load(f)
    check_for_required_config(config)

    newest_spreadsheet = find_latest_spreadsheet(config["search-dir"],
                                                 config["spreadsheet-pattern"])

    # start collecting metadata
    metadata = {}
    metadata['spreadsheet_mtime'] = os.path.getmtime(newest_spreadsheet)
    metadata['date_printed'] = datetime.now()

    # splitext() returns a tuple of (filename, extension)
    # we only want the extension, so we take the second element
    # and slice off the first character (the dot)
    extension = os.path.splitext(newest_spreadsheet)[1][1:]
    logger.debug(f"{extension=}")

    with open(newest_spreadsheet, "rb") as f:
        logger.debug(f"Read roster_df from {newest_spreadsheet}")
        if extension == "csv":
            roster_df = pd.read_csv(f)
        elif extension in ["xls", "xlsx", "xlsm", "xlsb", "odf", "ods", "odt"]:
            # list taken from https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html
            roster_df = pd.read_excel(f)
        else:
            logger.error(f"File type not supported: {newest_spreadsheet}")
            raise ValueError(f"File type not supported: {newest_spreadsheet}")

    logger.debug(f"{roster_df=}")
    session_date = roster_df[config.get("date-column")].values[0]
    if config.get('date-format', ""):
        metadata['session_date_str'] = dateparser.parse(session_date).strftime(config['date-format'])
    else:
        metadata['session_date_str'] = session_date

    if "modify-columns" in config.keys():
        logger.debug("modifying columns")
        for modify_data in config["modify-columns"]:
            if modify_data["new-name"] in roster_df.columns:
                raise ValueError(f"new column name {modify_data['new-name']} already exists in roster")
            # apply modifications
            logger.debug(f"merging columns {modify_data['old-columns']} into {modify_data['new-name']}")
            roster_df[modify_data["new-name"]] = roster_df[modify_data["old-columns"]].apply(
                lambda x: modify_data["separator"].join(x.dropna().astype(str)), axis=1)

        logger.debug(f"{roster_df.info()=}")

    print_all_sessions(roster_df, **metadata)

    logger.info("Printing complete! Please close the window")
