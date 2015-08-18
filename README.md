pptxindex
======

Using a concordance file of keywords or Python expressions, build a Microsoft Word index document from the slides and notes of one or more PowerPoint pptx files.

## Linux or OS X Installation

On Linux or Mac OS X systems, install the `python-docx` package using `pip`:

```
$ pip install python-docx
```

## Windows Installation

Windows users will need to install Python before using pptxindex.  Download and install Python 2.7 from https://www.python.org/downloads/windows.  After installation, add `C:\Python27` and `C:\Python27\scripts` to your PATH environment variable to make it easier to run from the command line.

After installing Python 2.7, run the `pip` utility in `C:\Python27\scripts` to install the python-docx dependencies:

```
C:\Python27\scripts> pip install python-docx
```

## Usage

```
$ python pptxindex.py
Usage: pptxindex.py -c <CONCORDANCE> [-o WORDFILE] [-i WORDFILE] [PPTX FILES]
                          [-h] [-t]
     -c <CONCORDANCE>    Specify the concordance filename
     -o <WORDFILE>       Specify the MS Word index output filename
     -i <WORDFILE>       Specify the MS Word template file to base index on
     -t                  Test and validate concordance file syntax, then exit
     -v                  Verbose output (including 0-hit concordance entries)
     -h                  This usage information
```

### Example usage:

Check the concordance file for validity:
```
$ pptxindex.py -c concordance.txt -t  
No errors in the concordance file.
$ ./pptxindex.py -c concordance-err.txt  -t
Error processing concordance file line 9: EOL while scanning string literal (<string>, line 1)
Please correct the errors in the concordance file and try again.
```

Generate an index for a single PowerPoint file, writing the Word index file to concordance.txt.docx:
```
$ ./pptxindex.py -c concordance.txt Sec575_1_A09_RSTR.pptx
Extracting content from PPTX files.
Searching for matches with the concordance file.
Creating index reference ranges.
Creating index document.
Done.
$ ls -l concordance.txt.docx
-rw-r--r--  1 jwright  staff  41536 Aug  6 08:28 concordance.txt.docx
```

Generate an index for multiple PowerPoint files, reporting on concordance entries that did not produce matches:
```
$ ./pptxindex.py -c concordance.txt -v ~/Documents/Sec575_*.pptx
Processing PPTX files: Sec575_1_A09_RSTR.pptx Sec575_2_A09_RSTR.pptx Sec575_3_A09_RSTR.pptx Sec575_4_A09_RSTR.pptx Sec575_5_A09_RSTR.pptx Sec575_6_A09_RSTR.pptx Sec575_Handout1.pptx
Extracting content from PPTX files.
Searching for matches with the concordance file.
The following entries in the concordance file did not produce matches:
    RecoverAndroidPin.py Tool : "recoverandroidpin" in page
    Recovery Mode
    iOS Application Folders : "ios" in page and ("application folder" in page or "/Applications" in page)
Creating index reference ranges.
Creating index document.
Done.
```

Generate a Word index file called "Sec575_A09_Index.docx", using the supplied template file as the first page of the output docx file:
```
$ ./pptxindex.py -c concordance.txt -o Sec575_A09_Index.docx -i Template.docx ~/Documents/Sec575_*.pptx
Extracting content from PPTX files.
Searching for matches with the concordance file.
Creating index reference ranges.
Creating index document.
Done.
$ ls -l Sec575_A09_Index.docx
-rw-r--r--  1 jwright  staff  34634 Aug  6 08:43 Sec575_A09_Index.docx
```

## Building a Concordance

The concordance file is the file by which you specify the entries that will be used to generate your index.  Your concordance file can be cery simple, matching a list of keywords to use when populating the index document, or it can be complex using Pythonic expressions.

The concordance file is a simple text file, in either Windows or Unix format.  Lines beginning with a "#" are ignored as comments.  Blank lines are also ignored.

### Simple Concordance Entry Matching

Your concordance file can contain a simple list of keywords to use when generating the index.  Each line (ignoring comments and blank lines) in the concordance file is used to produce the entry in the index, and as the search term to match in the PowerPoint files.

```
$ cat concordance.txt
Apple Configurator
iTunes
Cain & Abel
Wireshark
```

Here, each of the 4 entries in the concordance file will be used as both the search terms and the index entry produced in the output Word document.

### Customized Concordance Entry Matching

The concordance file can also use Pythonic expressions to determine if the page of content will produce a hit for a given search term, using a more elegant index entry description than what is actually used for the search term.  Consider the following example:

```
$ cat concordance.txt
SQL Injection;"sql injection" in page
SQL Injection, Testing Risk;"sql injection" in page and "testing risk" in page
```

Here, two concordance entries are used to index content associated with SQL injection.  The content to the left of the semi-colon is the string used in the index output.  The content to the right of the semi-colon is a Python expression returning a Boolean expression.

The Python expressions in the concordance can be almost anything that is valid Python, referring to one of six internal variables:

1. `page` - The `page` variable contains all the content from an individual page in one or more PowerPoint documents, converted to all lowercase characters.
2. `cspage` - The `cspage` variable is similar to `page`, except that it is case sentitive.
3. `pagenum` - The `pagenum` variable is the current page number being evaluated in the PowerPoint file.
4. `booknum` - The `booknum` variable is the current book number (see "On Page Numbering" below) in the PowerPoint file.
5. `wordlist` - The `wordlist` variable is a collection of each individual word on the page with punctuation and other special characters removed.  Use `wordlist` when you want to match a specific word but not when that word might appear *inside* of another word (e.g. match "APT", but not "captive")
6. `cswordlist` - The `cswordlist` variable is similar to `wordlist`, except that it is case sensitive.  This is particularly useful for matching acronyms.

In the first example from the concordance file, the string "SQL Injection" will be added to the index if the Python expression to the right of the semi-colon evaluates to True.  Here, the expression will be True if the string "sql injection" is present in any given PowerPoint page.  This functionality is identical to a concordance entry of simply "SQL Injection" with no expression following a semi-colon.

In the second example from the concordance file, the string "SQL Injection, Testing Risk" will similarly be added to the index if the Python expression to the right of the semi-colon evaluates to True.  Here, the expression will be True if the string "sql injection" is present in any given PowerPoint page, and if "testing risk" is also present in the same page.

**Note**: When creating concordance rules and referring to the `page` string variable, your search keyword should be specified lowercase. The concordance entry `SQL Injection;"SQL injection" or "SQLi"` will not return any hits since the strings in the `page` variable are always lowercase.

### Advanced Concordance Entry Matching

Since the concordance expressions are any valid Python expression, you can be creative in expressing when `pptxindex.py` will return a match for an indexed entry.

```
$ cat concordance.txt
Client Side Injection Attacks; ("client side injection" in page or "csi" in page) and "attack" in page
```

Here, the index entry is "Client Side Injection Attacks", which will trigger anytime the string "attack" is in the page, and when "client side injection" or "csi" is in the page as well.

```
$ cat concordance.txt
Burp Suite;page.count("burp suite") > 1
```

If I feel that the number of page hits are too frequent to be useful in the index (e.g. the indexed list of matches is more than what would be useful as a reference source), then I would consider only registering a match when the number of instances of the search term is greater than 1.  In the previous example, the index entry for "Burp Suite" will only match a page when "Burp Suite" occurs more than once on a page (in the notes of the slide bullets), eliminating most of the casual references to Burp Suite throughout the PowerPoint files.

```
$ cat concordance.txt
Penetration Testing, Mobile;("pentest" in page or "pen test" in page or "penetration test" in page) and (booknum == 2 or booknum == 3)
```

In this example, the index entry for "Penetration Testing, Mobile" will only appear when the strings "pentest", "pen test" or "penetration testing" appear, and the book is the 2nd or 3rd PowerPoint file in the list of files.

```
$ cat concordance.txt
Open Handset Alliance (OHA);"open handset alliance" in page or "OHA" in cswords
```

In this example, the string "open handset alliance" can appear anywhere in the page to register an index hit, or the letters "OHA" must appear as a word in a sentence of bullet (but not matching "aloha", for example).

## On Page Numbering

Pptxindex accepts one or more PowerPoint files to use for building the content to evaluate for indexing. When multiple PowerPoint files are specified, it creates the index using the notation *BookNumber:PageNumber*, or for a range of pages, *BookNumber:StartPageNumber-EndPageNumber* (e.x. "1:6", "4-125", "3:20-23").

The book number is taken from the PowerPoint filename order, alphanumerically sorted.  Consider the following PowerPoint filenames:

```
$ ls *.pptx
Sec575_1_A09.pptx
Sec575_2_A09.pptx
Sec575_3_A09.pptx
```

Here, the Sec575_1_A09.pptx will be marked as book number 1 (e.g. "1:6" for hits on page 6).  The file Sec575_2_A09.pptx would be page 2, etc.  This assumes you don't name your files like this:

```
$ ls *.pptx
Sec575_MobArch_1_A09.pptx
Sec575_MobPentest_2_A09.pptx
Sec575_MobDevRecommend_3_A09.pptx
```

In this naming convention, the file "Sec575_MobDevRecommend_3_A09.pptx" would be alphanumerically sorted as book 2, and Sec575_MobPentest_2_A09.pptx would be book 3.  *Don't name your files this way.*


## Questions, Comments, Concerns?

Open a ticket, or drop me a note: jwright@hasborg.com.

-Josh

