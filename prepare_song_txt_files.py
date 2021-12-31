#!/usr/bin/env python

"""This program preprocesses TXT files for faster manual editing afterwards."""

from typing import Tuple, Dict, Optional
import argparse
import glob
import os.path
from collections import namedtuple
import io
import pandas as pd
import warnings
import codecs
import json


__author__ = "Anton Höß"
__copyright__ = "Copyright 2021"
__credits__ = list()
__license__ = "BSD"
__version__ = "0.1"
__maintainer__ = "Anton Höß"
__email__ = "anton.hoess42@gmail.com"
__status__ = "Development"


class Txt:
    def __init__(self, header: str, body: str, fn_out: str):
        self.header = header
        self.body = body
        self.fn_out = fn_out
    # end def
# end class


class PreprocessTxt:
    """The preprocessing class."""

    def __init__(self, fn_db: str, db_header_offset: int, cols: Dict[str, str], country_lang_assignment: Dict[str, str], song_name_assignment: Optional[str] = None) -> None:
        """Initializes the Converter.

        Parameters
        ----------
        fn_db : str
            The filename of the Excel database  file to read the metadata from.
        db_header_offset : int
            The number of rows to ignore before interpreting data as header and body.
        db_header_offset : dict of int, int
            Dictionary holding the required columns and their internally used names.
        country_lang_assignment : dict of str, str
            The assignment dictionary of country to language codes.
        song_name_assignment : str
            The song name assignment filename.
        """

        self._fn_db = fn_db
        self._db_header_offset = db_header_offset
        self._cols = cols
        self._country_lang_assignment = country_lang_assignment
        self._song_name_assignment = song_name_assignment

        # Read data from Excel database
        warnings.filterwarnings("ignore")  # Suppress certain warning
        self._df = pd.read_excel(self._fn_db, header=self._db_header_offset, usecols=list(self._cols.keys()), dtype=str)
        warnings.filterwarnings("default")
        self._df.rename(columns=self._cols, inplace=True)  # Rename the columns for more flexibility afterwards

        # Read song name assignments to dictionary
        self._name_changes = dict()

        if self._song_name_assignment is not None:
            if not os.path.isfile(self._song_name_assignment):
                raise FileNotFoundError(f"Specified song name assignment file {self._song_name_assignment} not found!")
            else:
                with io.open(self._song_name_assignment, mode="r", encoding="utf-8") as f:
                    data = f.read()
                    data = data.splitlines()

                    for line in data:
                        line = line.strip()
                        if line == "" or len(line) >= 1 and line[0] == "#":  # Don't allow inline comments tho prevent problems in case "#" is part of the song name
                            continue
                        song_name_file, song_name_db = line.split("=")
                        self._name_changes[song_name_file] = song_name_db.replace(r"\n", "\n")
                    # end for
                # end with
            # end if
        # end if
    # end def

    def preprocess(self, path: str, filename: str) -> Optional[Txt]:
        """Does the preprocessing.

        Parameters
        ----------
        path : str
            The path of the file to preprocess.
        filename : str
            The filename itself.

        Returns
        -------
        txt : Txt
            Txt-object holding the resulting data.
        """

        # Check if filename is valid and can be modified
        if len(filename) < 14 or filename[-14:-10] != "_AP_":
            print(f"Invalid filename {filename}!")
            return None
        # end if

        # Load file and preprocess its data
        cur_song_title = None
        with io.open(os.path.join(path, filename), mode="r", encoding="utf-8") as f:
            data = f.read()

            # Line operations
            #################
            data = data.splitlines()

            # Strip lines
            for i in range(len(data)):
                data[i] = data[i].strip()
            # end for

            # Remove title block
            title_idx = -1
            for i, line in enumerate(data):
                if line == "Titel:":
                    title_idx = i
                    break
                # end if
            # end for

            cur_song_title = data[title_idx + 1]

            if title_idx != -1:
                for i in range(3):  # Delete three lines ("Titel:"; <the title itself>; newline
                    del data[title_idx]
                # end for
            # end if

            # Merge lines to a single string again
            data = "\n".join(data)

            # Whole file operations
            #######################
            data = data.replace(" ", " ")
            data = data.replace("͜", "")
            data = data.replace("Verse:\n", "")
            data = data.replace("REFRAIN:", "Refrain:")
            data = data.replace("Refrain:\n", "R. ")
            data = data.replace("CODA:", "Coda:")
            data = data.replace("BRIDGE:", "Bridge:")
        # end with

        # Resolve names that are not matching between the TXT files and the entry in the database
        tmp_cur_song_title = cur_song_title
        if cur_song_title in self._name_changes:
            tmp_cur_song_title = self._name_changes[cur_song_title]

        # Get all associated data from db
        entry = self._df.loc[self._df['TITLE'] == tmp_cur_song_title]  # old: entry = self._df[self._df["TITLE"].str.fullmatch(tmp_cur_song_title).fillna(False)]
        if entry.size == 0:
            raise ValueError(f"Entry with title \"{tmp_cur_song_title}\" not found in Excel database.")

        title = cur_song_title
        title_original = entry.TITLE_ORIGINAL.values[0]
        title_original_lang = entry.TITLE_ORIGINAL_LANG.values[0]
        ref_no = entry.REF_NO.values[0]
        copy_right = entry.COPYRIGHT.values[0].splitlines()
        # year_of_translation = entry.YEAR_OF_TRANSLATION.values[0]

        # Strip entries (just to be sure)
        title = title.strip()
        if not pd.isna(title_original):
            title_original = title_original.strip()
        if not pd.isna(title_original_lang):
            lang = self._country_lang_assignment.get(title_original_lang.strip())
            if lang is not None:
                title_original_lang = lang
            else:
                raise ValueError(f'TITLE_ORIGINAL_LANG ({title_original_lang}) not in country language assignment dict ({self._country_lang_assignment}).')
        if not pd.isna(ref_no):
            ref_no = ref_no.strip()
        else:
            ref_no = "NOREF"
        copy_right = [line.strip() for line in copy_right]
        # if not pd.isna(year_of_translation):
        #     year_of_translation = year_of_translation.strip()

        # Compile header
        header = ""
        header += f"TITLE={title}\n"
        if not pd.isna(title_original):
            header += f"TITLE_ORIGINAL={title_original}\n"
            if not pd.isna(title_original_lang):
                header += f"TITLE_ORIGINAL_LANG={title_original_lang}\n"
        header += f"REF_NO={ref_no}\n"
        # No CAPO
        header += f"AUTHORS={copy_right[0]}\n"
        header += f"COPYRIGHT={copy_right[1]}\n"

        # Create result object
        return Txt(header=header, body=data, fn_out=f"{cur_song_title.replace(':', '')} {ref_no}.txt")
    # end def

    @staticmethod
    def save(txt: Txt, output: str, force_overwrite: bool = False) -> None:
        """Saves the preprocessed TXT file.

        Parameters
        ----------
        txt : Txt
            Txt-object holding the resulting data that shall be saved.
        output : str
            The output folder.
        force_overwrite : bool, default False
            Indicates if the output file(s) shall get overwritten if they already exist.
        """

        # Save the file
        if not os.path.exists(output):
            os.mkdir(output)

        fn_out = os.path.join(output, txt.fn_out)

        if not os.path.exists(fn_out) or force_overwrite:
            with codecs.open(fn_out, "w", "utf-8") as f:
                f.write(txt.header + "\n" + txt.body)

        else:
            raise FileExistsError(f"File {fn_out} already exists and therefore not got overwritten.")
    # end def
# end class


def main() -> None:
    """The main function which parses the program arguments and performs the conversion of the specified files."""

    def get_base_path_from_wildcard_path(wc_path: str) -> Tuple[str, int]:
        wc_path = os.path.normpath(wc_path)
        parts = wc_path.split(os.sep)

        # Special treatment for Windows drive letters
        drive = os.path.splitdrive(wc_path)[0]

        if parts[0] == drive:
            parts[0] = drive + os.path.sep

        cur_path = ""
        parts_cnt = 0
        for part in parts:
            tmp_path = os.path.join(cur_path, part)

            if not os.path.exists(tmp_path):
                break

            parts_cnt += 1
            cur_path = os.path.join(cur_path, part)
        # end for

        return cur_path, parts_cnt
    # end def

    def str2bool(v):
        if isinstance(v, bool):
            return v
        # end if

        if v.lower() in ('yes', 'true', 't', 'y', '1'):
            return True

        elif v.lower() in ('no', 'false', 'f', 'n', '0'):
            return False

        else:
            raise argparse.ArgumentTypeError('Boolean value expected.')
        # end if
    # end def

    # Set up parser for command line arguments
    parser = argparse.ArgumentParser(description="Preprocesses TXT files.")
    parser.add_argument("paths", type=str, nargs='+',
                        help=r"Filenames or folder base paths. Accepts wildcards as well as "
                             r"\**\ for recursive search. Example: ./**/*.txt.")
    parser.add_argument("--output", type=str, default=".",
                        help="Output folder")
    parser.add_argument("--force_overwrite", type=str2bool, default=False,
                        help="Overwrite existing output files.")
    parser.add_argument("--suppress_error_output", type=str2bool, default=True,
                        help="Suppress the error output (traceback) and only print the error string "
                             "without any further information.")
    parser.add_argument("--excel_database", type=str, required=True,
                        help="Excel database file.")
    parser.add_argument("--db_header_offset", type=int, required=False, default=8,
                        help="Number of lines to skip before interpreting the remaining data as header and data.")
    parser.add_argument("--cols", type=json.loads, required=False,
                        default='{'
                                '"Ref.-Nr.:": "REF_NO",'
                                '"Titel": "TITLE",'
                                '"Originaltitel": "TITLE_ORIGINAL",'
                                '"Ursprungsland": "TITLE_ORIGINAL_LANG",'
                                '"Gesamte Copyrightangabe © (extern)": "COPYRIGHT",'
                                '"Übersetzungsjahr": "YEAR_OF_TRANSLATION"'
                                '}',
                        help="Excel database file.")
    parser.add_argument("--song_name_assignment", type=str, required=False, default=None,
                        help="Song name assignment filename. Necessary to assign TXT files with database entries in cases they were not written the same."
                             "The format of each line is as follows: \"<TITLE_TXT>=<TITLE_DB>\". Comment lines need to start with #.")
    parser.add_argument("--country_lang_assignment", type=json.loads, required=False,
                        default='{'
                                '"AT": "DE",'
                                '"DE": "DE",'
                                '"EN": "EN",'
                                '"USA": "EN",'
                                '"FR": "FR",'
                                '"IT": "IT",'
                                '"NL": "NL",'
                                '"PL": "PL"'
                                '}',
                        help="Country to language assignment (ISO 639).")

    # Parse arguments
    args = parser.parse_args()

    FilePath = namedtuple("FilePath", "base_path rel_path filename")
    file_paths = list()

    for path in args.paths:
        # If argument is a path, add it ...
        if os.path.isfile(path):
            p, f = os.path.basename(path)
            file_paths.append(FilePath(p, "", f))
        else:
            # ... otherwise it is a filter using wildcards, so evaluate it and add all matched files
            base_path, base_path_parts_cnt = get_base_path_from_wildcard_path(path)

            for fn in glob.glob(path, recursive=True):
                if os.path.isfile(fn):
                    p, f = os.path.split(fn)
                    file_paths.append(FilePath(base_path, p[len(base_path) + 1:], f))
                # end if
            # end for
        # end if
    # end for

    # Load database from Excel file - just do it once for all songs at it takes some time
    fn_song_db = args.excel_database
    pp = None
    try:
        # Load preprocessor
        pp = PreprocessTxt(fn_db=fn_song_db, db_header_offset=args.db_header_offset, cols=args.cols, country_lang_assignment=args.country_lang_assignment,
                           song_name_assignment=args.song_name_assignment)

    except Exception as e:
        print(f"\n=> ERROR on opening Excel database occurred: {e}")
        if not args.suppress_error_output:
            raise e
        # end if
    # end try

    # Process all files
    for base_path, rel_path, file in file_paths:
        print(f"Processing file \"{file}\"...")

        try:
            out_path = os.path.join(args.output, rel_path)
            if not os.path.exists(out_path):
                os.makedirs(out_path)

            # Preprocess...
            txt = pp.preprocess(path=os.path.join(base_path, rel_path), filename=file)

            if txt:
                # Save result
                pp.save(txt=txt, output=out_path, force_overwrite=args.force_overwrite)
                print("Finished!")
            else:
                print("Error!")

        except Exception as e:
            print(f"\n=> ERROR occurred: {e}")
            if not args.suppress_error_output:
                raise e
            # end if
        # end try
    # end for

# end def


if __name__ == "__main__":
    main()
# end if
