#!/usr/bin/env python3
from Crypto.Cipher import AES
from Crypto.Hash import MD5
import logging
import json
import struct
import os
import glob
from datetime import datetime
from argparse import ArgumentParser
from os.path import abspath, dirname
# ---------------------------------------------------------------------------NEW
import csv
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

class PBEWITHMD5AND128BITAES_CBC_OPENSSL:
    def __init__(self, password, salt):
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
    def __init__(self, json_string):
        self._json = json_string
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
        for n in sorted(self._notes.values(), key=lambda item: item.get_modified_date()):
            #print(n)
            yield n

##
# MAIN
def main():
    # -----------------------------------------------------------------------NEW
    out_name = "ColorNote_backup"
    out_extn = ".csv"
    out_sep = "_"
    # json_keys = ["_id", "account_id", "active_state", "color_index",
    #              "created_date", "dirty", "encrypted", "folder_id",
    #              "importance", "latitude", "longitude", "minor_modified_date",
    #              "modified_date", "note", "note_ext", "note_type",
    #              "reminder_base", "reminder_date", "reminder_duration",
    #              "reminder_last", "reminder_option", "reminder_repeat",
    #              "reminder_repeat_ends", "reminder_type", "revision", "space",
    #              "staged", "status", "tags", "title", "type", "uuid"]
    json_keys_select = ["_id", "color_index", "created_date",
                        "minor_modified_date", "modified_date", "note",
                        "revision", "title"]
    # --------------------------------------------------------------------------

    _salt = b'ColorNote Fixed Salt'

    parser = ArgumentParser()
    parser.add_argument("-p", "--password", action="store", type=str,
                    dest="password", default="0000",
                    help="password for uncrypting backup notes")
    parser.add_argument("--csv",
                    action="store_true", default=False,
                    help="output as a CSV file")
    parser.add_argument("-v", "--verbose",
                    action="store_true", dest="verbose", default=False,
                    help="[For debug] verbose output")
    parser.add_argument("--binary",
                    action="store_true", default=False,
                    help="[For debug] dump the decrypted binary in tmp/notes.bin")
    parser.add_argument('colornote_backup_dir', type=str,
                        help='ColorNote backup directory (with .doc files)')


    options = parser.parse_args()

    if not os.path.isdir(options.colornote_backup_dir):
        parser.error("Argument '{}' is not a directory or doesn't exist".format(options.colornote_backup_dir))


    logging.basicConfig(level=logging.DEBUG if options.verbose else logging.WARNING)

    notes = NotesSet()

    decoder = PBEWITHMD5AND128BITAES_CBC_OPENSSL(options.password.encode('utf-8'), _salt)

    for bakfile in glob.iglob(os.path.join(options.colornote_backup_dir, '**', '*.doc'), recursive=True):
        logging.debug("Parsing %s...", bakfile)

        doc = open(bakfile, "rb").read()

        decoded_doc = decoder.decrypt(doc[28:])

        # For debug purpose: create output path and dump the decrypted binary
        if options.binary:
            directory = dirname(abspath(__file__))
            tmp_path = os.path.join(directory, "tmp")
            os.makedirs(tmp_path, exist_ok=True)
            open(os.path.join(tmp_path, "notes.bin"), "wb").write(bytes(str(decoded_doc),"utf-8"))

        # locate substring to give start offset = idx + 4
        substring = b'{"_id":1,"title"'
        offset = decoded_doc.find(substring)
        extract = decoded_doc[offset:offset+len(substring)].decode("utf-8")
        logging.debug("%10d: %s", offset, extract)
        idx = offset - 4

        while idx + 4 < len(decoded_doc):
            # File is padded with something like 0f0f0f0f or 0b0b0b0b...
            if (decoded_doc[idx] == decoded_doc[idx+1] and
                decoded_doc[idx+1] == decoded_doc[idx+2] and
                decoded_doc[idx+2] == decoded_doc[idx+3]):
                break
            (chunk_length,) = struct.unpack(">L", decoded_doc[idx:idx+4])
            logging.debug("Chunk length: %d", chunk_length)
            chunk = decoded_doc[idx+4:idx+chunk_length+4]
            logging.debug("Chunk: %s", chunk.decode())
            json_chunk = json.loads(chunk.decode("utf-8"))
            notes.update_if_newer(Note(json_chunk))
            idx += chunk_length + 4


    if options.csv:
        dtn = datetime.now()
        dtymdhms = dtn.strftime("%Y%m%d_%H%M%S")
        file_path = out_name + out_sep + dtymdhms + out_extn
        out_file = open(file_path, 'a', newline='', encoding='utf-8')
        csvwriter = csv.writer(out_file, delimiter=',')

        # write headers
        csvwriter.writerow(json_keys_select)

    for n in notes.get():
        if not n.is_archived():
            print("-"*50)
            logging.debug(n)
            print(n.get_title())
            print("Created at {}\t Modified at {}".format(n.get_created_date(), n.get_modified_date()))
            print(n.get_note())
            if options.csv:
                njson = json.loads(str(n))
                row = []
                for key in json_keys_select:
                    match key:
                        case "created_date":
                            #value = n.get_created_date().isoformat()
                            #value = n.get_created_date().strftime("%Y-%m-%d %H:%M:%S.%f")
                            value = get_excel_date(n.get_created_date())
                        case "minor_modified_date":
                            #value = n.get_minor_modified_date().isoformat()
                            #value = n.get_minor_modified_date().strftime("%Y-%m-%d %H:%M:%S.%f")
                            value = get_excel_date(n.get_minor_modified_date())
                        case "modified_date":
                            #value = n.get_modified_date().isoformat()
                            #value = n.get_modified_date().strftime("%Y-%m-%d %H:%M:%S.%f")
                            value = get_excel_date(n.get_modified_date())
                        case _:
                            value = njson[key]
                    row.append(value)
                csvwriter.writerow(row)
                logging.debug(njson)
                logging.debug(njson.keys())
                logging.debug(njson.values())
                logging.debug(njson["note"])

    if options.csv:
        out_file.close()

if __name__ == "__main__":
    main()
