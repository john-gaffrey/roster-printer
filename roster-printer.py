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

logger = logging.getLogger("roster-printer")

# debug flags
DEBUG = os.getenv("ROSTER_PRINTER_DEBUG", "False") == "True"
PRINT_ROSTERS = not DEBUG
USE_TEMPDIR = not DEBUG
LOG_LEVEL = logging.DEBUG if DEBUG else logging.INFO

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

def modify_roster_columns(roster: pd.DataFrame, columns_to_merge: list[dict[str]]) -> pd.DataFrame:
    """Returns a DataFrame with the all of the specified columns to be updated.
       columns_to_merge is a list of dicts, where each dict has the keys:
       'new-name' (the name of the new column) and 'old-columns' (a list of columns to merge)
       This can also be used to rename columns by setting 'old-columns' to a list of 1 column"""
    result = roster.copy()

    for merge in columns_to_merge:
        if "new-name" not in merge.keys() or "old-columns" not in merge.keys():
            raise KeyError("merge dict must contain keys 'new-name' and 'old-columns'")
        if merge["new-name"] in roster.columns:
            raise ValueError(f"new column name {merge['new-name']} already exists in roster")
        # merge the columns
        logger.debug(f"merging columns {merge['old-columns']} into {merge['new-name']}")
        result[merge["new-name"]] = roster[merge["old-columns"][0]]
        # result[merge["new-name"]] = roster[merge["old-columns"]].apply(
        #     lambda x: ' '.join(x.dropna().astype(str)), axis=1)
        
        # drop the old columns
        logger.debug(f"dropping columns {merge['old-columns']}")
        result = roster.drop(columns=merge["old-columns"])

    logger.debug(result.info())
    return result

def roster_to_pdf(roster: pd.DataFrame, file_path, title) -> None:
    """Creates a nicly formatted pdf of the `roster` at `file_path`"""

    # modified example from https://py-pdf.github.io/fpdf2/Maths.html#using-pandas
    pdf = FPDF(orientation="P", format="Letter", unit="pt")
    pdf.set_title(title)
    pdf.add_page()

    # create header
    pdf.set_font('helvetica', size=24)
    pdf.cell(text=title, new_y="NEXT", align="C", center=True)
    pdf.cell(text=datetime.today().strftime("%m/%d/%Y"), new_y="NEXT", align="C", center=True)
    pdf.ln(20)

    # add table
    pdf.set_font('helvetica', size=12)
    with pdf.table(
        borders_layout="MINIMAL",
        cell_fill_color=200,
        cell_fill_mode="ROWS",
        text_align="CENTER",
    ) as table:
        # fpdf table expects the header to be in the first row
        # in an iterable. Dataframes store them seperately, so we
        # must combine them for this to work nicely
        for data_row in [list(roster)] + roster.values.tolist():
            row = table.row()
            for datum in data_row:
                row.cell(datum)

    pdf.output(file_path)
    logger.debug(f"created pdf {file_path}")


def print_roster(roster: pd.DataFrame, title: str, directory: os.PathLike) -> None:
    """Prints a `roster` with the title: `title`"""

    tmp_file_path = os.path.join(directory, f"{title}.pdf")
    roster_to_pdf(roster, tmp_file_path, title=title)

    if PRINT_ROSTERS is True:
        logger.info(f"printing {title}.pdf")
        os.startfile(tmp_file_path, "print")
    else:
        logger.info(f"opening {title}.pdf")
        os.startfile(tmp_file_path, "open")


def print_all_sessions(roster: pd.DataFrame, config: dict) -> None:
    """Prints all unique sessions in roster, 1 per page"""
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
                # session_df = filter_roster_columns(roster_df, config["columns-to-print"])
                logger.debug(f"{session_df}")
                print_roster(session_df, title=f"{session} {config['title-suffix']}", directory=tempdir)
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
            session_df = roster.query(f"{config['class-column-name']} == @session")

            # filter to just the desired columns
            session_df = session_df[config["columns-to-print"]]
            logger.debug(f"{session_df}")
            print_roster(session_df, title=f"{session} {config['title-suffix']}", directory=tempdir)
        time.sleep(30)
    
if __name__ == "__main__":
    # configure logging
    logging.basicConfig(stream=sys.stdout, level=LOG_LEVEL)

    # get config
    CONFIG_FILE = os.getenv("CONFIG_FILE", "./config.yaml")
    logger.debug(f"{CONFIG_FILE=}")

    with open(CONFIG_FILE, encoding="utf-8") as f:
        config = yaml.safe_load(f)
    check_for_required_config(config)

    newest_spreadsheet = find_latest_spreadsheet(config["search-dir"],
                                                 config["spreadsheet-pattern"])
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
            raise(ValueError(f"File type not supported: {newest_spreadsheet}"))

    logger.debug(f"{roster_df=}")

    if "modify-columns" in config.keys():
        logger.debug("modifying columns")
        # roster_df = modify_roster_columns(roster_df, config["modify-columns"])
        for merge in config["modify-columns"]:
            if merge["new-name"] in roster_df.columns:
                raise ValueError(f"new column name {merge['new-name']} already exists in roster")
            # merge the columns
            logger.debug(f"merging columns {merge['old-columns']} into {merge['new-name']}")
            roster_df[merge["new-name"]] = roster_df[merge["old-columns"]].apply(
                lambda x: ' '.join(x.dropna().astype(str)), axis=1)

            # drop the old columns
            logger.debug(f"dropping columns {merge['old-columns']}")
            roster_df = roster_df.drop(columns=merge["old-columns"])
        logger.debug(f"{roster_df.info()=}")

    print_all_sessions(roster_df, config)
