# -*- coding: utf-8 -*-
"""
Created on Mon Mar  4 17:30:13 2019

@author: a-whalen
"""

import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import messagebox
from tkinter.filedialog import askdirectory
import os
import re
import datetime
from polyglot.detect import Detector
import xlsxwriter

def parse_xml(filepath):
    """
    filepath: complete path to xml file (including filename)
    returns list containing all non-space text in file
    """
    tree = ET.parse(filepath)
    root = tree.getroot()
    text = []
    for x in root.itertext():
        if x:
            text.append(x)
    return text
    
def get_xml():
    """
    Prompts user to select a folder,
    and returns a list of xml filepaths
    """
    filelist = []
    filepath = askdirectory()
    for root, dirs, files in os.walk(filepath):
        for file in files:
            if file.endswith("xml"):
                filelist.append(os.path.join(os.path.normpath(root), file))
    return filelist

def get_languages(filelist):
    """
    Finds filename, text, confidence for all text detected as English
    """
    output = (["Filename", "Text", "Language Code", "Confidence"],)
    for file in filelist:
        text = parse_xml(file)
        for i in text:
            item = i.strip()
            if item:
                regex1 = re.match(r"\([\d|A-Z]\)", item)
                if not regex1:
                    for language in Detector(item, quiet=True).languages:
                        if language.code == "en":
                            output += ([file, item, language.code, language.confidence],)
    return output

def create_excel(output):
    """
    Creates an Excel file listing the filename, text, langcode, 
    and confidence for all text detected as English
    """
    savedate = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    savename = savedate + "_" + "LanguageReport.xlsx"
    workbook = xlsxwriter.Workbook(savename)
    worksheet = workbook.add_worksheet()
    row = 0
    col = 0
    wrap_text = workbook.add_format()
    wrap_text.set_text_wrap()
    worksheet.set_column(0, 3, 50)
    for filename, text, langcode, confidence in (output):
       worksheet.write(row, col, filename, wrap_text)
       worksheet.write(row, col + 1, text, wrap_text)
       worksheet.write(row, col + 2, langcode, wrap_text)
       worksheet.write(row, col + 3, confidence, wrap_text)
       worksheet.set_row(row, 50)
       row += 1
    workbook.close()
    return

base = tk.Tk()
base.withdraw()

filelist = get_xml()
output = get_languages(filelist)
create_excel(output)
messagebox.showinfo('言語検出ツール', '終了しました！')



