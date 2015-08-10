#!/usr/bin/env python

# -*- coding: utf-8 -*-
#

from pptx import Presentation
from docx import Document
from itertools import groupby
from operator import itemgetter
from zipfile import ZipFile
from xml.dom.minidom import parse

import sys
import re
import os
import shutil
import glob
import tempfile
import operator
import getopt
import pdb
import string

def parseslidenotes(pptxfile, words, booknum):
    tmpd = tempfile.mkdtemp()
    ZipFile(pptxfile).extractall(path=tmpd, pwd=None)
    path = tmpd + '/ppt/notesSlides/'
    for infile in glob.glob(os.path.join(path, '*.xml')):
        #parse each XML notes file from the notes folder.
        dom = parse(infile)
        noteslist = dom.getElementsByTagName('a:t')

        # The page number is part of the filename
        page = int(re.sub(r'\D', "", infile.split("/")[-1]))

        # Create dictionary entry without content
        words[str(booknum) + ":" + str(page)] = ""
        text = ''

        for node in noteslist:
            xmlTag = node.toxml()
            xmlData = xmlTag.replace('<a:t>', '').replace('</a:t>', '')
            #concatenate the xmlData to the text for the particular slideNumber index.
            text += " " + xmlData

        # Convert to ascii to simplify
        text = text.encode('ascii', 'ignore')
        words[str(booknum) + ":" + str(page)] = " " + text


    path = tmpd + '/ppt/slides/'
    for infile in glob.glob(os.path.join(path, '*.xml')):
        #parse each XML notes file from the notes folder.
        dom = parse(infile)

        noteslist = dom.getElementsByTagName('a:t')

        page = int(re.sub(r'\D', "", infile.split("/")[-1]))
        text = ''

        for node in noteslist:
            xmlTag = node.toxml()
            xmlData = xmlTag.replace('<a:t>', '').replace('</a:t>', '')
            text += " " + xmlData

        # Convert to ascii to simplify
        text = text.encode('ascii', 'ignore')
        words[str(booknum) + ":" + str(page)] += " " + text

    # Remove all the files created with unzip
    shutil.rmtree(tmpd)

    # Remove double-spaces which happens in the content occasionally
    for page in words:
        words[page] = ''.join(ch for ch in words[page] if ch not in set([',','(',')']))
        words[page] = re.sub('\. ', " ", words[page])
        words[page] = ' '.join(words[page].split())

    return words

# Parse the text on slides using the python-pptx module, return words
def parseslidetext(prs, words, booknum):
    page = 0
    for slide in prs.slides:
        page += 1
        text_runs = ""
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    try:
                        text_runs += (" " + run.text)
                    except TypeError:
                        continue

        # SEC575 Specific -- skip slides that are course outline slides
        pdb.set_trace()
        if text_runs.find("Course") and text_runs.find("Roadmap"): 
            words[str(booknum)+ ":" + str(page)] = ""
        else:
            words[str(booknum)+ ":" + str(page)] = text_runs
    return words

# Validate the contents of the concordance file
def checkconcordance(concordancefile):
    page = ""
    cspage = ""
    ret=0
    lineno=0
    for line in open(concordancefile):
        expression = None
        lineno+=1
        if line[0] == "#" or line == "\n" or line == "\r\n" or line.isspace(): continue
        try:
            key,expression = line.strip().split(";")
        except ValueError:
            # Explicit search term, continue
            continue
        if expression != None:
            try:
                eval(expression)
            except Exception, e:
                ret=1
                sys.stdout.write("Error processing concordance file line " + str(lineno) + ": ")
                sys.stdout.write(str(e))
                sys.stdout.write("\n")
            continue
    return ret

# Take the index of entries, sort and reduce the page numbers
# into ranges (e.g. 1:3,1:4,1:5,1:8 into 1:3-5,1:8
# This is awful and I hope I never have to edit this code.
def indexreduce(index):
    for entry in index:
        matchesbybook = {}
        pages=index[entry]
        for bookpage in pages:
            book,page = bookpage.split(":")
            page = int(page)
            try:
                matchesbybook[book].append(page)
            except KeyError:
                matchesbybook[book] = [page]
        for book in matchesbybook:
            sortedreduced=[]
            matchesbybook[book].sort()
            matchesbybook[book] = numreduce(matchesbybook[book])
            # Return to 1:66, 2:57 format
            
        index[entry] = []
        for book in matchesbybook:
            for page in matchesbybook[book]:
                index[entry].append(book + ":" + page)

    return index

# Take a list of numbers and reduce them into hyphenated ranges for
# sequential values.
def numreduce(data):
    str_list = []
    for k, g in groupby(enumerate(data), lambda (i,x):i-x):
       ilist = map(itemgetter(1), g)
       #print ilist
       if len(ilist) > 1:
          str_list.append('%d-%d' % (ilist[0], ilist[-1]))
       else:
          str_list.append('%d' % ilist[0])
    return str_list 


def indexsort(string):
    book,page = string.split(":")
    page = re.sub('-.*', "", page)
    return int(book)*(int(page)+1000)

def showconcordancehits(concordancehits, concordance):
    # Count the number of hits in the concordancehits dictionary
    # where val != None
    nohitcount=0
    for key in concordancehits:
        if concordancehits[key] == None:
           nohitcount+=1
    if nohitcount == 0:
        print "All entries in the concordance file produced matches."
        return
    else:
        print "The following entries in the concordance file did not produce matches:"
        for key in concordancehits:
            if concordancehits[key] == None:
                if concordance[key] == None:
                    print "\t" + key
                else:
                    print "\t" + key + " : " + concordance[key]


def usage(status=0):
    print "pptxindex v1.0.1"
    print "Usage: pptxindex.py -c <CONCORDANCE> [-o WORDFILE] [-i WORDFILE] [PPTX FILES]"
    print "                          [-h] [-t]"
    print "     -c <CONCORDANCE>    Specify the concordance filename"
    print "     -o <WORDFILE>       Specify the MS Word index output filename"
    print "     -i <WORDFILE>       Specify the MS Word template file to base index on"
    print "     -t                  Test and validate concordance file syntax, then exit"
    print "     -v                  Verbose output (including 0-hit concordance entries)"
    print "     -h                  This usage information"
    sys.exit(status)

if __name__ == "__main__":

    concordancefile = None
    indexoutputfile = None
    testandexit     = None
    templatefile    = None
    verbose         = False

    if len(sys.argv) == 1: usage(0)

    opts = getopt.getopt(sys.argv[1:],"i:c:o:htv")
    
    for opt,optarg in opts[0]:
        if opt == "-c":
            concordancefile = optarg
        elif opt == "-o":
            indexoutputfile = optarg
        elif opt == "-i":
            templatefile = optarg
        elif opt == "-t":
            testandexit = True
        elif opt == "-v":
            verbose = True
        elif opt == "-h":
            usage()
    
    if not concordancefile:
        print "Error: concordance file not specified"
        usage()

    if not indexoutputfile:
        indexoutputfile = concordancefile + ".docx"

    # Check all the expressions in the concordance file
    if (checkconcordance(concordancefile) != 0):
        sys.stderr.write("Please correct the errors in the concordance file and try again.\n")
        sys.exit(-1)

    if testandexit:
        print("No errors in the concordance file.")
        sys.exit(0)

    # Read concordance file and build the dictionary
    concordance = {}
    for line in open(concordancefile):
        if line[0] == "#" or line == "^$": continue
        try:
            key,val = line.strip().split(";")
            concordance[key] = val
        except ValueError:
            concordance[line.strip()] = None
    

    # Handle globbing for pptx filenames on Windows
    pptxfiles = []
    for filemask in opts[1:][0]:
        pptxfiles += glob.glob(filemask)
    if len(pptxfiles) == 0:
        sys.stderr.write("No matching PPTX files found.\n")
        sys.exit(1)
    pptxfiles.sort()
    if verbose:
        print("Processing PPTX files: %s")%' '.join(os.path.basename(x) for x in pptxfiles)

    print("Extracting content from PPTX files.")
    wordsbypage = {}
    booknum=1
    for pptxfile in pptxfiles:
        try:
            prs = Presentation(pptxfile)
        except Exception:
            sys.stderr.write("Invalid PPTX file: " + pptxfile + "\n")
            sys.exit(-1)
        
        wordsbypage = parseslidetext(prs,wordsbypage,booknum)
        try:
            wordsbypage = parseslidenotes(pptxfile,wordsbypage,booknum)
        except:
            print "Unexpected error:", sys.exc_info()[0]
            sys.exit(-1)
            
        booknum+=1

    # Next, iterate through the concordance dictionary, searching for and recording
    # matches for each entry.
    print("Searching for matches with the concordance file.")
    index = {}
    concordancehits = {}
    for key in concordance:
        pages = [] # list of page numbers
        for bookpagenum in wordsbypage:
            # To track hits with concordance entries, mark hits for this
            # entry to None by default.
            concordancehits[key] = None

            # These are the variables intended to be accessible by the author in the concordance file
            cspage = wordsbypage[bookpagenum]
            page = wordsbypage[bookpagenum].lower()
            booknum,pagenum = bookpagenum.split(":")

            # Process the concordance file entry.  If it is None, then use 
            # the key as the search string
            if concordance[key] == None:
                if (key.lower() in page):
                        pages.append(bookpagenum)
            # Else, evaluate the right-side of the concordance entry as a Python expression
            elif eval(concordance[key]):
                pages.append(bookpagenum)

            # If the concordance entry generated some matches, add it to the index list
            if pages != []:
                index[key] = pages
                concordancehits[key] = True

    if verbose:
        showconcordancehits(concordancehits, concordance)

    # Reduce index entries "1:1,1:2,1:3" to 1:1-3"
    print("Creating index reference ranges.")
    index = indexreduce(index)

    # Sort the reduced index entries numerically
    for page in index:
        index[page] = sorted(index[page], key=indexsort)

    # With index list created, make the Word document
    print("Creating index document.")
    document = Document(templatefile)
    if templatefile != None:
        document.add_page_break()
    
    document.add_heading('Index', level=1)
    table = document.add_table(rows=0, cols=2, style="Light Shading")
    for entry in sorted(index.keys(), key=str.lower):
        row_cells = table.add_row().cells
        row_cells[0].text = entry
        row_cells[1].text = ", ".join(index[entry])

    document.save(indexoutputfile)
    print("Done.")
