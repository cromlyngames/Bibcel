# -*- coding: utf-8 -*-
"""
Created on Fri Dec 17 15:02:13 2021

@author: pb910
"""

# a wee script that converts many downloaded bibtex refs into a single bibtex file
# it looks in target folder and all sub folders,
# gathers up all files of target type and puts them in a single file


outfile = "glaze_colour.bib"
folder =  "globholder"
#unique folder name you've put the targets in
filetype = ".bib"

import glob2  # script for file collecting up
pathy =  "*" + filetype
list_of_files = glob2.glob(pathy)

outfile = input("output filename pls. \n >>>__")
if outfile[-4:] != ".bib":
    outfile = outfile + ".bib"


for f in list_of_files:
    #don't add the outfile to itself!
    if f != outfile:
        try:   # it can break on umlauts or other advanced characters
            f1 = open(outfile, 'a+', encoding='utf-8')
            f2 = open(f, 'r')
            f1.write(f2.read())
            f1.close()
            f2.close()
        except:
            print("error with file ", f)
            # I normally manually copy these ones over.
    
    
        

