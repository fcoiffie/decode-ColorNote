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

		open("/tmp/notes.bin", "wb").write(decoded_doc)

		idx = 0x10
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

	for n in notes.get():
		if not n.is_archived():
			print('--------')
			logging.debug(n)
			print(n.get_title())
			print("Created at {}\t Modified at {}".format(n.get_created_date(), n.get_modified_date()))
			print(n.get_note())

if __name__ == "__main__":
	main()
