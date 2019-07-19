# lang-detect
Parses xml files and outputs any English text in an Excel

Created in Python 3.7.1

Non-standard dependent libraries: polyglot, xlsxwriter

I wrote this script to help look for any untranslated English text in DITA xml files.

When run from the command line, this script will prompt the user to select a folder. It will then find all xml files in the folder, parse them, and look for any text that polyglot detects as being English. The filename, text, language code, and confidence level will be output in an Excel file.
