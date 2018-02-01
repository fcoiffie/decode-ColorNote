from Crypto.Cipher import AES
from Crypto.Hash import MD5
import binascii
import json
import struct

def _get_derived_key_and_iv(password, salt):
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
	hasher.update(_password)
	hasher.update(_salt)
	result = hasher.digest()

	key = result

	hasher = MD5.new()
	hasher.update(result)
	hasher.update(_password)
	hasher.update(_salt)
	result = hasher.digest()

	iv = result

	# key, iv
	return key, iv

##
# MAIN
# Just a test.

_password = b'0000'
_salt = b'ColorNote Fixed Salt'
#_iterations = 20 # In fact, not required for derivation


doc = open("/home/fcoiffie/Documents/ColorNote/backup/1517196902886-AUTO.doc", "rb").read()

(key, iv) = _get_derived_key_and_iv(_password, _salt)
encoder = AES.new(key, AES.MODE_CBC, iv)

decoded_doc = encoder.decrypt(doc[28:])

# Remove padding done with 0xF0
decoded_doc = decoded_doc.rstrip(b'\x0f')
#print(decoded_doc)

json_chunks = []
idx = 0x10
while idx < len(decoded_doc):
	(chunk_length,) = struct.unpack(">L", decoded_doc[idx:idx+4])

	chunk = decoded_doc[idx+4:idx+chunk_length+4]
	json_chunk = json.loads(chunk.decode("utf-8"))
	json_chunks.append(json_chunk)
	idx += chunk_length + 4

for j in json_chunks:
	print(json.dumps(j, sort_keys=True, indent=4))

#json = json.loads(decoded_doc[0x14:].decode("utf-8"))
#print()

#print(json.dumps(json, sort_keys=True, indent=4))
