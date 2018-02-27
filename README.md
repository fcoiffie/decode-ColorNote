# ColorNote decoder

This small Python script decrypts .doc ColorNote  (https://www.colornote.com) backup files, converts to JSON and dumps their contents on the console.

The AES-128-CBC decrypt part is based on (Many thanks to olejorgenb) : https://github.com/olejorgenb/ColorNote-backup-decryptor

_Example_:
```
$ ./decode-ColorNote.py ~/Documents/ColorNote/
--------
My todolist
Created at 2018-02-05 15:16:03.559000    Modified at 2018-02-09 17:27:20.633000
[ ] Item 3
[ ] Item 2
[V] Item 1
```

_Note_: I synchronize my backup files between my smartphone and my PC with: https://syncthing.net/

## Requirements

`pycrypto` package is used for uncrypting the .doc files:
```
  pip install pycrypto
```

## TODO
- [ ] Load Color tag  
- [ ] Export to HTML  
- [ ] Add new options for sorting/filtering notes
