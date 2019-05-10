#coding=utf-8

import sys
reload(sys)
sys.setdefaultencoding("utf-8")

import pickle
import re
import  codecs
import string
import shutil
from win32com import client as wc
import docx
import os
 
def doSaveAas():

    word = wc.Dispatch('Word.Application')
    for filename in os.listdir(r'C:/Users/mazy/Desktop/books/'):
    	if "docx" in filename or "doc" not in filename:
    		continue
    	print(filename)
    	doc = word.Documents.Open("C:/Users/mazy/Desktop/books/"+filename)       
    	doc.SaveAs("C:/Users/mazy/Desktop/books/"+filename+"x", 12, False, "", True, "", False, False, False, False)    
    	doc.Close()
    word.Quit()


    # doc = word.Documents.Open("C:/Users/mazy/Desktop/books/test.doc")       
    # doc.SaveAs("C:/Users/mazy/Desktop/books/test.docx", 12, False, "", True, "", False, False, False, False)    
    # doc.Close()
    # word.Quit()
 
doSaveAas()