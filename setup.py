#!/usr/bin/python
import os

os.system("sudo cp spreadsheetScript /usr/share/man/man1/spreadsheetScript.1")
os.system("sudo gzip /usr/share/man/man1/spreadsheetScript.1")
