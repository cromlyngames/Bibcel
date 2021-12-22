# Bibcel
Python tools for linking bibtex files and an excel reading list manager.

## Design Statement

I want to maintain my reading list of papers and articles. These will be split into smaller themed bibtex files, rather than have a single clunky one. I want to have a single speadsheet tracking these, whether I've read them, found them or predict that they'll be valable, and tracking keywords from the author and added by me.
In two years time, I want to be able to easily locate pertinant information.

built on python-bibtexparser and xlwings

## Not quite what you were looking for?
try
* https://github.com/sciunto-org/python-bibtexparser/issues/284  Bibtex tidy up tools likely to be merged into python-bibtexparser in the future
* https://github.com/vitorfs/parsifal  A systematic literature review toolset
* https://github.com/Phyks/BMC Download and manage pdfs from urls / doi

## Contents:
globber_script : 
* barebones trivial function to merge a lot of downlaoded bibtex refs into a single file

001_excel_maker : currently two functions: 
* reads in a bibtex file to create an excel sheet with column headers for entry info and keywords 
* reads in a bibtex file and appends the data to the next empty row(s)

## Todo:
* Add a better coloumn header finder to the append bibtex file function
* Add a duplicate cull function
* ~~Improve keyword splitter. (currently works with ; only. Some bibtex use , ). There's a regex option in bibtex parser that might be valid if I can figure it out
* Add new function to write out updated bibtex files from the excel sheet info (eg to allow bulk update of entryid)
