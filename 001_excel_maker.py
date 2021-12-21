# -*- coding: utf-8 -*-
"""
Created on Fri Dec 17 15:34:31 2021

@author: pb910
"""

# takes bibtex files 

import xlwings as xw # modern excel library
### NOTE xlwings/Excel Ranges are 1-based. (not python's default of 0 based)
import bibtexparser as bt
from bibtexparser.bparser import BibTexParser
from bibtexparser.customization import homogenize_latex_encoding

def Initial_col_namer(target_bib):
    print("opening ", target_bib)
    # https://bibtexparser.readthedocs.io/en/master/tutorial.html
    with open(target_bib, encoding='utf-8') as bibtex_file:
        parser = BibTexParser()
        parser.customization = homogenize_latex_encoding
        bib_database = bt.load(bibtex_file, parser=parser)
    print("mining ", target_bib)
    
    # database is a list of dictionaries, where each dict is one bibtex entry
    #print(bib_database.entries[1].keys())
        
    # get all of the diict keyz to get all the future column headings
    list_keyz = []
    for b in bib_database.entries:
        keyz = b.keys()
        for k in keyz:
            list_keyz.append(k)
    set_keyz = set(list_keyz)
    sl_keyz = list(set_keyz) #cast it back to indexed list for easier iterations later
    #print(set_keyz)
    
    first_cols =['title', 'source_bib', 'read', 'value']
    
    keywordz = [] 
    for b in bib_database.entries:
        try:
            keywrd_str =b['keywords']
        except:
            keywrd_str = ''  #not all enteries have the keywords entry
        #spilt the string into words, and append to overall list
        keyword_list = keywrd_str.split(';')
        for w in keyword_list:
            keywordz.append(w)
    set_wordz = set(keywordz) #cast it back to indexed list for easier iterations later
    sl_wordz = list(set_wordz)
    sl_wordz.sort()     # put the keywords into alphabetical order
    
    wb = xw.Book('reading_lists.xlsx')
    sheet = wb.sheets['Sheet1']
    all_colz =  first_cols+sl_keyz+sl_wordz
    for i in range(1,len(all_colz)):
        sheet.range((1,i)).value = all_colz[i-1]
    return(all_colz)



def get_col_headers():
    print("error, need a list of col headers.")
    print("This fucntion is on the todo list")
    
    
    
def get_next_empty_line(sheet):
    iterate = True
    row = 1
    while iterate:
        tempval = sheet.range((row,1)).value
        if tempval != None:
            row += 1
        else:
            iterate = False
    print("starting at row" ,row)
    return(row)
    
def col_filler(col_headers):
    
    if col_headers == [] or col_headers == None:
        col_headers = get_col_headers()
    
    target_bib = input("target_bib pls. \n >>> ")
    if target_bib[-4:] != ".bib":
        target_bib = target_bib + ".bib"
    
    print("opening ", target_bib)
    # https://bibtexparser.readthedocs.io/en/master/tutorial.html
    with open(target_bib, encoding='utf-8') as bibtex_file:
        parser = BibTexParser()
        parser.customization = homogenize_latex_encoding
        bib_database = bt.load(bibtex_file, parser=parser)
    print("mining ", target_bib)
    
    print("opening prepared excel file")
    wb = xw.Book('reading_lists.xlsx')
    sheet = wb.sheets['Sheet1']
    
    start_at = get_next_empty_line(sheet)
    
    for row in range(start_at, start_at+len(bib_database.entries)): #step through the bibtex enteries
        entry = bib_database.entries[row-start_at]    
        for col in range(1, len(col_headers)+1): #step through the cols
            
            try: 
                entry['keywords']
                haskeywords = True
            except KeyError:
                haskeywords = False
                
        
            try: #looking up col_header in entry dictionary
                tval = entry[col_headers[col-1]]
                sheet.range((row,col)).value = tval
            except KeyError:
                if col_headers[col-1] == 'source_bib':
                    tval = target_bib
                    sheet.range((row,col)).value = tval
                elif col_headers[col-1] == 'read' or col_headers[col-1] == 'value':
                    tval = "" # and don't update sheet
                elif haskeywords == True: 
                     if col_headers[col-1] in entry['keywords']:
                         tval = 1
                         sheet.range((row,col)).value = tval
                
    return()
                   
          

#MAIN
target_bib = 'test.bib'
colz_headers = Initial_col_namer(target_bib)
col_filler(colz_headers)
