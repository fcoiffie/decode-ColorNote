#!/usr/bin/env python3
from Crypto.Cipher import AES
from Crypto.Hash import MD5
import logging
import binascii
import json
import struct
import os
import glob
from datetime import datetime
from optparse import OptionParser
# ---------------------------------------------------------------------------NEW
import csv
import sys
import traceback
from datetime import timezone
from os.path import abspath, dirname
# ---------------------------------------------------------------------------END

# ---------------------------------------------------------------------------NEW
def get_excel_date(dt):
    # Microsoft Excel date is number of days since 1899-12-30
    # Microsoft Excel date = 1 = 1900-01-01
    # seconds per day = 60*60*24 = 86400
    UTC = timezone.utc
    dt_python = dt.replace(tzinfo=UTC)
    dt_excel_zero = datetime(1899, 12, 30, tzinfo=UTC)
    return (dt_python - dt_excel_zero).total_seconds() / 86400
# ---------------------------------------------------------------------------END

# ---------------------------------------------------------------------------NEW
def get_html_table(id, caption, list_head, list_data, datetime, indent):
    # aligns numbers right for all key values except "note" and "title"
    # with option to indent tags
    tab = " " * indent if indent else ""
    index_note = list_head.index("note")
    index_title = list_head.index("title")
    list = []
    list.append('<table id="' + id + '">')
    list.append(tab * 1 + '<caption>' + caption + '</caption>')
    list.append(tab * 1 + '<thead>')
    list.append(tab * 2 + '<tr>')
    list.append(tab * 3 + '<th>' + '</th><th>'.join(str(x) for x in list_head) + '</th>')
    list.append(tab * 2 + '</tr>')
    list.append(tab * 1 + '</thead>')
    list.append(tab * 1 + '<tbody>')
    for lrow in list_data:
        list.append(tab * 2 + '<tr>')
        list_row = []
        for index, item in enumerate(lrow):
            if (index != index_note) and (index != index_title):
                try:
                    number_test = int(item)
                    list_row.append('<td style="text-align: right;">' + str(item) + '</td>')
                except:
                    list_row.append('<td>' + str(item) + '</td>')
            else:
                list_row.append('<td>' + str(item) + '</td>')
        row = "".join(list_row)
        list.append(tab * 3 + row)
        list.append(tab * 2 + '</tr>')
    list.append(tab * 1 + '</tbody>')
    list.append(tab * 1 + '<tfoot>')
    list.append(tab * 2 + '<tr>')
    list.append(tab * 3 + '<td colspan="' + str(len(list_head)) + '">created: ' + datetime + '</td>')
    list.append(tab * 2 + '</tr>')
    list.append(tab * 1 + '</tfoot>')
    list.append('</table>')
    return list
# ---------------------------------------------------------------------------END

class PBEWITHMD5AND128BITAES_CBC_OPENSSL:
    def __init__(self, password, salt, iterations):
        # Iterations aren't used
        (self._key, self._iv) = self._get_derived_key_and_iv(password, salt)

    def _get_derived_key_and_iv(self, password, salt):
        """
        Returns tuple of key(16 bytes) and iv(16 bytes) for AES
        Logic:
        This code is inspired by :
        /**
        * Generator for PBE derived keys and ivs as usd by OpenSSL.
        * <p>
        * The scheme is a simple extension of PKCS 5 V2.0 Scheme 1 using MD5 with an
        * iteration count of 1.
        * <p>
        */
        public class OpenSSLPBEParametersGenerator
        :param password: password used for encryption/decryption
        :param salt: salt
        :return: (16 bytes dk, 16 bytes iv)
        """

        hasher = MD5.new()
        hasher.update(password)
        hasher.update(salt)
        result = hasher.digest()
        key = result

        hasher = MD5.new()
        hasher.update(result)
        hasher.update(password)
        hasher.update(salt)
        result = hasher.digest()
        iv = result

        # key, iv
        return key, iv

    def decrypt(self, data):
        encoder = AES.new(self._key, AES.MODE_CBC, self._iv)
        return encoder.decrypt(data)

class Note:
    def __init__(self, json):
        self._json = json
    def __repr__(self):
        return json.dumps(self._json, sort_keys=True, indent=4)
    def get_uuid(self):
        return self._json['uuid']
    def get_created_date(self):
        return datetime.fromtimestamp(self._json['created_date'] / 1000)
    def get_minor_modified_date(self):
        return datetime.fromtimestamp(self._json['minor_modified_date'] / 1000)
    def get_modified_date(self):
        return datetime.fromtimestamp(self._json['modified_date'] / 1000)
    def is_archived(self):
        return self._json['space'] == 16
    def get_title(self):
        return self._json['title']
    def get_note(self):
        return self._json['note']

class NotesSet:
    def __init__(self):
        self._notes = {}
    def has_uuid(self, uuid):
        return uuid in self._notes.keys()
    def update_if_newer(self, note):
        if not self.has_uuid(note.get_uuid()):
            self._notes[note.get_uuid()] = note
        elif self._notes[note.get_uuid()].get_minor_modified_date() < note.get_minor_modified_date():
            self._notes[note.get_uuid()] = note
        #else:
            # Nothing to do
    def get(self):
        for (k,n) in sorted(self._notes.items(), key=lambda item: item[1].get_modified_date()):
            #print(n)
            yield n

##
# MAIN
def main():
    # -----------------------------------------------------------------------NEW
    """
    If the ColorNote backup directory is not given on the command line, the
    script will search in directory 'backup_dirname' using paths as follows:
        1. specified_backup_path
        2. user document path     (if specified_backup_path = "")
    Similarly, tmp and output directories 'tmp_dirname' and 'output_dirname'
    will be created using paths as follows:
        1. specified_tmp_path
        2. user document path     (if specified_tmp_path = "")
    and:
        1. specified_output_path
        2. user document path     (if specified_output_path = "")
    For example:
        backup  specified  ->  r"C:\<specified_backup_path>"
        backup  user       ->  r"C:\<user path>\Documents\<backup_dirname>"
        output  specified  ->  r"C:\<specified_output_path>"
        output  user       ->  r"C:\<user path>\Documents\<output_dirname>"
        tmp     specified  ->  r"C:\<specified_tmp_path>"
        tmp     user       ->  r"C:\<user path>\Documents\<tmp_dirname>"
    """
    specified_backup_path = ""
    specified_tmp_path = "ColorNote_tmp"
    specified_output_path = "ColorNote_output"
    backup_dirname = "backup"
    tmp_dirname = "ColorNote_tmp"
    output_dirname = "ColorNote_output"
    html_template = r"decode-ColorNote-CSV-HTML-TEMPLATE.html"
    out_name_csv = "ColorNote_backup"
    out_extn_csv = ".csv"
    out_sep_csv = "_"
    out_name_html = "ColorNote_backup"
    out_extn_html = ".html"
    out_sep_html = "_"

    # define 'json_keys_select'
    # define as empty list to use all available keys
    # 'json_keys_select' must include key: "color_index" - for html output
    json_keys_select = [] # empty list
    json_keys_select = ["_id", "color_index",
                        "created_date", "minor_modified_date", "modified_date",
                        "note", "revision", "title"]
    # json keys available:
    #     "_id", "account_id", "active_state", "color_index",
    #     "created_date", "dirty", "encrypted", "folder_id",
    #     "importance", "latitude", "longitude", "minor_modified_date",
    #     "modified_date", "note", "note_ext", "note_type",
    #     "reminder_base", "reminder_date", "reminder_duration",
    #     "reminder_last", "reminder_option", "reminder_repeat",
    #     "reminder_repeat_ends", "reminder_type", "revision", "space",
    #     "staged", "status", "tags", "title", "type", "uuid"
    # -----------------------------------------------------------------------END

    _salt = b'ColorNote Fixed Salt'
    _iterations = 20 # In fact, not required for derivation

    logger = logging.getLogger()
    #logger.setLevel(logging.DEBUG)

    parser = OptionParser()
    parser.add_option("-p", "--password", action="store", type="string",
                    dest="password", default="0000",
                    help="password for uncrypting backup notes")
    parser.add_option("-q", "--quiet",
                    action="store_false", dest="verbose", default=True,
                    help="don't print status messages to stdout")

    (options, args) = parser.parse_args()

    #if len(args) != 1:
    #    parser.error("ColorNote backup directory is missing")
    #if not os.path.isdir(args[0]):
    #    parser.error("Argument '{}' is not a directory or doesn't exist".format(args[0]))

    # -----------------------------------------------------------------------NEW
    # assign generic paths
    script_path = dirname(abspath(__file__))
    user_home_path = os.path.expanduser("~")
    user_docs_path = os.path.join(user_home_path, "Documents")
    # -----------------------------------------------------------------------END

    # -----------------------------------------------------------------------NEW
    # check backup path
    user_backup_path = os.path.join(user_docs_path, backup_dirname)
    user_backup_path = user_backup_path if os.path.isdir(user_backup_path) else ""
    if len(args) != 1:
        backup_directory = specified_backup_path if os.path.isdir(specified_backup_path) else user_backup_path
    else:
        backup_directory = args[0]
    if not os.path.isdir(backup_directory):
        parser.error("Argument '{}' is not a directory or doesn't exist".format(backup_directory))
    # -----------------------------------------------------------------------END

    #backup_directory = args[0] # ORIGINAL

    notes = NotesSet()

    decoder = PBEWITHMD5AND128BITAES_CBC_OPENSSL(options.password.encode('utf-8'), _salt, _iterations)

    # ----------------------------------------------------------MODIFIED AND NEW
    # renamed 'decoded_doc' as 'decoded_backup_file'
    backup_files = []
    for type in ("*.backup", "*.dat", "*.doc"):
        backup_files.extend(glob.iglob(os.path.join(backup_directory, "**", type), recursive=True))
    logging.debug(backup_files)
    for bakfile in backup_files:
    #for bakfile in glob.iglob(os.path.join(backup_directory, '**', '*.dat'), recursive=True): # ORIGINAL
        logging.debug(bakfile)

        backup_file = open(bakfile, "rb").read()

        bakfile_type = os.path.splitext(bakfile)[1].lower()

        # handle file types - assign 'magic_offset'
        match bakfile_type:
            case ".backup":
                magic_offset = 28 # 12 also appears to work
            case ".dat":
                magic_offset = 0
            case ".doc":
                magic_offset = 28
            case _:
                print("Backup file type not recognised.  Require '.backup' , '.dat' or '.doc'.")
                Exit

        decoded_backup_file = decoder.decrypt(backup_file[magic_offset:])

        #open("/tmp/notes.bin", "wb").write(decoded_doc) # ORIGINAL

        # create tmp path
        user_tmp_path = os.path.join(user_docs_path, tmp_dirname)
        tmp_directory = specified_tmp_path if specified_tmp_path else user_tmp_path
        os.makedirs(tmp_directory, exist_ok=True)

        # write to tmp file
        open(os.path.join(tmp_directory, "notes.bin"), "wb").write(bytes(str(decoded_backup_file), "utf-8"))
        #open(os.path.join(tmp_directory, "notes.bin"), "wb").write(decoded_backup_file)

        # locate substring to give start offset = idx + 4
        substring = b'{"_id":1,"title"'
        offset = decoded_backup_file.find(substring)
        extract = decoded_backup_file[offset:offset+len(substring)].decode("utf-8")
        logging.debug(f"{offset: <10}: {extract}")

        # handle file types - assign 'idx'
        match bakfile_type:
            case ".backup":
                idx = offset - 4
            case ".dat":
                idx = 0x10
            case ".doc":
                idx = offset - 4
            case _:
                idx = offset - 4

        #idx = 0x10 # ORIGINAL
        while idx + 4 < len(decoded_backup_file):
            # File is padded with something like 0f0f0f0f or 0b0b0b0b...
            if (decoded_backup_file[idx] == decoded_backup_file[idx+1] and
                decoded_backup_file[idx+1] == decoded_backup_file[idx+2] and
                decoded_backup_file[idx+2] == decoded_backup_file[idx+3]):
                break
            (chunk_length,) = struct.unpack(">L", decoded_backup_file[idx:idx+4])
            logging.debug("Chunk length: {}".format(chunk_length))
            chunk = decoded_backup_file[idx+4:idx+chunk_length+4]
            logging.debug("Chunk: {}".format(chunk))
            json_chunk = json.loads(chunk.decode("utf-8"))
            notes.update_if_newer(Note(json_chunk))
            idx += chunk_length + 4
    # -----------------------------------------------------------------------END

    #for n in notes.get():
    #    if not n.is_archived():
    #        print('--------')
    #        logging.debug(n)
    #        print(n.get_title())
    #        print("Created at {}\t Modified at {}".format(n.get_created_date(), n.get_modified_date()))
    #        print(n.get_note())

    # -----------------------------------------------------------------------NEW
    # define date time strings
    dt = datetime.now()
    dtiso = dt.isoformat()
    dtymdhms = dt.strftime("%Y%m%d_%H%M%S")
    dtymd = dt.strftime("%Y%m%d")
    # -----------------------------------------------------------------------END

    # -----------------------------------------------------------------------NEW
    # create output path
    user_output_path = os.path.join(user_docs_path, output_dirname)
    output_directory = specified_output_path if specified_output_path else user_output_path
    os.makedirs(output_directory, exist_ok=True)

    # write to csv file
    file_name = out_name_csv + out_sep_csv + dtymdhms + out_extn_csv
    output_file = os.path.join(output_directory, file_name)
    #with open(output_file, "a", newline="", encoding="iso-8859-1") as out_file:
    with open(output_file, "a", newline="", encoding="utf-8") as out_file:
        csvwriter = csv.writer(out_file, delimiter=",")

        list_html = []

        # write json keys as headers
        json_keys = []
        if json_keys_select:
            json_keys = json_keys_select
        else:
            for n in notes.get():
                njson = json.loads(str(n))
                json_keys = list(njson.keys())
                break
        csvwriter.writerow(json_keys)

        # write json values as rows
        for n in notes.get():
            if not n.is_archived():
                print("-"*50)
                logging.debug(n)
                print(n.get_title())
                print("Created at {}\t Modified at {}".format(n.get_created_date(), n.get_modified_date()))
                print(n.get_note())
                njson = json.loads(str(n))
                row_csv = []
                row_html = []
                for key in json_keys:
                    match key:
                        case "created_date":
                            value_csv = get_excel_date(n.get_created_date())
                            value_html = n.get_created_date().isoformat()
                            #value = n.get_created_date().strftime("%Y-%m-%d %H:%M:%S.%f")
                        case "minor_modified_date":
                            value_csv = get_excel_date(n.get_minor_modified_date())
                            value_html = n.get_minor_modified_date().isoformat()
                            #value = n.get_minor_modified_date().strftime("%Y-%m-%d %H:%M:%S.%f")
                        case "modified_date":
                            value_csv = get_excel_date(n.get_modified_date())
                            value_html = n.get_modified_date().isoformat()
                            #value = n.get_modified_date().strftime("%Y-%m-%d %H:%M:%S.%f")
                        case _:
                            value_csv = njson[key]
                            value_html = njson[key]
                    row_csv.append(value_csv)
                    row_html.append(value_html)
                csvwriter.writerow(row_csv)
                #csvwriter.writerow(njson.values())
                list_html.append(row_html)
                logging.debug(njson)
                logging.debug(njson.keys())
                logging.debug(njson.values())
                logging.debug(njson["note"])
    # -----------------------------------------------------------------------END

    # -----------------------------------------------------------------------NEW
    list_table = get_html_table("colornote_table", "ColorNote Table",
                                json_keys, list_html, dtiso, 0)
    search_table_placeholder = "COLORNOTE_TABLE_PLACEHOLDER"
    replace_table_text = "\n".join(list_table)
    search_index_placeholder = "COLOR_INDEX_COLUMN_PLACEHOLDER"
    replace_index_text = str(json_keys.index("color_index"))
    # read from html template file
    with open(html_template, "r") as in_file:
        data = in_file.read()
        data = data.replace(search_table_placeholder, replace_table_text)
        data = data.replace(search_index_placeholder, replace_index_text)
    # write to html file
    file_name = out_name_html + out_sep_html + dtymdhms + out_extn_html
    output_file = os.path.join(output_directory, file_name)
    #with open(output_file, "w", newline="", encoding="iso-8859-1") as out_file:
    with open(output_file, "w", newline="", encoding="utf-8") as out_file:
        out_file.write(data)
    # -----------------------------------------------------------------------END

if __name__ == "__main__":
    #main() # ORIGINAL
    # -----------------------------------------------------------------------NEW
    #input()
    try:
        main()
        print("\nFinished.")
    except:
        print(sys.exc_info()[0])
        print(traceback.format_exc())
        print("\nProgram terminated.")
    finally:
        print("\nPlease press Enter to exit...", end="")
        input()
    # -----------------------------------------------------------------------END
