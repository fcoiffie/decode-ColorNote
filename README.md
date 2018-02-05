# ColorNote decoder

This small Python script decrypts .doc ColorNote  (https://www.colornote.com) backup files, converts to JSON and dumps their contents on the console.

The AES-128-CBC decrypt part is based on (Many thanks to olejorgenb) : https://github.com/olejorgenb/ColorNote-backup-decryptor

## Requirements

`pycrypto` package is used for uncrypting the .doc files:
```
  pip install pycrypto
```

## TODO
[ ] Load Color tag  
[ ] Export to HTML  
[ ] Add new options for sorting/filtering notes
