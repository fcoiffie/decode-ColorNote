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
from os.path import abspath, dirname
# ---------------------------------------------------------------------------NEW
import csv
import traceback
from datetime import timezone
# ------------------------------------------------------------------------------

# ---------------------------------------------------------------------------NEW
def get_excel_date(dt):
    # Microsoft Excel date is number of days since 1899-12-30
    # Microsoft Excel date = 1 = 1900-01-01
    # seconds per day = 60*60*24 = 86400
    UTC = timezone.utc
    dt_python = dt.replace(tzinfo=UTC)
    dt_excel_zero = datetime(1899, 12, 30, tzinfo=UTC)
    return (dt_python - dt_excel_zero).total_seconds() / 86400
# ------------------------------------------------------------------------------

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
        # ----------------------------------------------------------------------
        list_row = []
        for index, item in enumerate(lrow):
            if (index != index_note) and (index != index_title):
                try:
                    tmp = int(item)
                    list_row.append('<td style="text-align: right;">' + str(item) + '</td>')
                except:
                    list_row.append('<td>' + str(item) + '</td>')
            else:
                list_row.append('<td>' + str(item) + '</td>')
        row = "".join(list_row)
        list.append(tab * 3 + row)
        # ----------------------------------------------------------------------
        list.append(tab * 2 + '</tr>')
    list.append(tab * 1 + '</tbody>')
    list.append(tab * 1 + '<tfoot>')
    list.append(tab * 2 + '<tr>')
    list.append(tab * 3 + '<td colspan="' + str(len(list_head)) + '">created: ' + datetime + '</td>')
    list.append(tab * 2 + '</tr>')
    list.append(tab * 1 + '</tfoot>')
    list.append('</table>')
    return list
# ------------------------------------------------------------------------------

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
    # for default_backup_path and default_tmp_path, either assign null string ""
    # to use user 'Documents' folder or specify full path, for example:
    #     backup    null string   ->  r"C:\<user path>\Documents\backup"
    #     backup    full path     ->  r"C:\<some path>\backup"
    #     tmp       null string   ->  r"C:\<user path>\Documents\tmp"
    #     tmp       full path     ->  r"C:\<some path>\tmp"
    default_backup_path = ""
    default_tmp_path = ""
    html_template = r"decode-ColorNote-CSV-HTML-TEMPLATE.html"
    out_name_csv = "ColorNote_backup"
    out_extn_csv = ".csv"
    out_sep_csv = "_"
    out_name_html = "ColorNote_backup"
    out_name_html_tab = "ColorNote_backup_tabulate"
    out_extn_html = ".html"
    out_sep_html = "_"
    json_keys_select = ["_id", "color_index", "created_date",
                        "minor_modified_date", "modified_date", "note",
                        "revision", "title"]
    # json_keys_select must include key: "color_index" - for html table creation
    # json keys available:
    #     "_id", "account_id", "active_state", "color_index",
    #     "created_date", "dirty", "encrypted", "folder_id",
    #     "importance", "latitude", "longitude", "minor_modified_date",
    #     "modified_date", "note", "note_ext", "note_type",
    #     "reminder_base", "reminder_date", "reminder_duration",
    #     "reminder_last", "reminder_option", "reminder_repeat",
    #     "reminder_repeat_ends", "reminder_type", "revision", "space",
    #     "staged", "status", "tags", "title", "type", "uuid"
    # --------------------------------------------------------------------------

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

    if len(args) != 1:
        parser.error("ColorNote backup directory is missing")
    if not os.path.isdir(args[0]):
        parser.error("Argument '{}' is not a directory or doesn't exist".format(args[0]))

    backup_directory = args[0]

    notes = NotesSet()

    decoder = PBEWITHMD5AND128BITAES_CBC_OPENSSL(options.password.encode('utf-8'), _salt, _iterations)

    for bakfile in glob.iglob(os.path.join(backup_directory, '**', '*.doc'), recursive=True):
        logging.debug(bakfile)

        doc = open(bakfile, "rb").read()

        decoded_doc = decoder.decrypt(doc[28:])

        #open("/tmp/notes.bin", "wb").write(decoded_doc) # ORIGINAL

        # -------------------------------------------------------------------NEW
        # create output path
        directory = dirname(abspath(__file__))
        tmp_path = os.path.join(directory, "tmp")
        os.makedirs(tmp_path, exist_ok=True)
        open(os.path.join(tmp_path, "notes.bin"), "wb").write(bytes(str(decoded_doc),"utf-8"))
        # ----------------------------------------------------------------------

        # -------------------------------------------------------------------NEW
        # locate substring to give start offset = idx + 4
        substring = b'{"_id":1,"title"'
        offset = decoded_doc.find(substring)
        extract = decoded_doc[offset:offset+len(substring)].decode("utf-8")
        logging.debug(f"{offset: <10}: {extract}")
        idx = offset - 4
        # ----------------------------------------------------------------------

        #idx = 0x10 # ORIGINAL
        while idx + 4 < len(decoded_doc):
            # File is padded with something like 0f0f0f0f or 0b0b0b0b...
            if (decoded_doc[idx] == decoded_doc[idx+1] and
                decoded_doc[idx+1] == decoded_doc[idx+2] and
                decoded_doc[idx+2] == decoded_doc[idx+3]):
                break
            (chunk_length,) = struct.unpack(">L", decoded_doc[idx:idx+4])
            logging.debug("Chunk length: {}".format(chunk_length))
            chunk = decoded_doc[idx+4:idx+chunk_length+4]
            logging.debug("Chunk: {}".format(chunk))
            json_chunk = json.loads(chunk.decode("utf-8"))
            notes.update_if_newer(Note(json_chunk))
            idx += chunk_length + 4

    #for n in notes.get():
    #    if not n.is_archived():
    #        print('--------')
    #        logging.debug(n)
    #        print(n.get_title())
    #        print("Created at {}\t Modified at {}".format(n.get_created_date(), n.get_modified_date()))
    #        print(n.get_note())

    # -----------------------------------------------------------------------NEW
    dt = datetime.now()
    dtiso = dt.isoformat()
    dtymdhms = dt.strftime("%Y%m%d_%H%M%S")
    dtymd = dt.strftime("%Y%m%d")
    # --------------------------------------------------------------------------

    # -----------------------------------------------------------------------NEW
    file_path = out_name_csv + out_sep_csv + dtymdhms + out_extn_csv
    #with open(file_path, "a", newline="", encoding="iso-8859-1") as out_file:
    with open(file_path, "a", newline="", encoding="utf-8") as out_file:
        csvwriter = csv.writer(out_file, delimiter=",")

        list_html = []

        # write json keys as headers
        json_keys = []
        if json_keys_select:
            json_keys = json_keys_select
        else:
            for n in notes.get():
                njson = json.loads(str(n))
                json_keys = njson.keys()
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
    # --------------------------------------------------------------------------

    # -----------------------------------------------------------------------NEW
    list_table = get_html_table("colornote_table", "ColorNote Table",
                                json_keys, list_html, dtiso, 0)
    search_table_placeholder = "COLORNOTE_TABLE_PLACEHOLDER"
    replace_table_text = "\n".join(list_table)
    search_index_placeholder = "COLOR_INDEX_COLUMN_PLACEHOLDER"
    replace_index_text = str(json_keys.index("color_index"))
    # read from source file
    with open(html_template, "r") as in_file:
        data = in_file.read()
        data = data.replace(search_table_placeholder, replace_table_text)
        data = data.replace(search_index_placeholder, replace_index_text)
    # write to target file
    file_path = out_name_html + out_sep_html + dtymdhms + out_extn_html
    #with open(file_path, "w", newline="", encoding="iso-8859-1") as out_file:
    with open(file_path, "w", newline="", encoding="utf-8") as out_file:
        out_file.write(data)
    # --------------------------------------------------------------------------

if __name__ == "__main__":
    main()
