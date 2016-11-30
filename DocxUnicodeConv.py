#! /usr/local/bin/python
#
# The concept behind this code is taken from page
# http://stackoverflow.com/questions/17850227/text-replace-in-docx-and-save-the-changed-file-with-python-docx

import os
import sys
import getopt
import zipfile

from docx import Document
from BaminiDict import bamini_dict
from AdhawinTamilDict import adhawintamil_dict 

english_fonts = ["Times New Roman"]

def usage():
    print ""
    print "This tool converts word document (.docx) file with Bamini font to"
    print "Unicode font without changing the format of the document"
    print ""
    print "usage: ./Bamini2UniConv -i <input file>"
    print ""

def print_run(run):
    print "font = "+str(run.font.name)
    print "text = "+run.text

def convert_bamini(paragraph):
    text = paragraph.encode("utf-8")
    for key in bamini_dict.keys():
        text = text.replace(str(key), str(bamini_dict.get(key)))
    print "CONVERTED (Bamini)!!"
    return text.decode("utf-8")    

def convert_adhawintamil(paragraph):
    text = paragraph.encode("utf-8")
    for key in adhawintamil_dict.keys():
        text = text.replace(str(key), str(adhawintamil_dict.get(key)))
    print "CONVERTED (Adhawin-Tamil)!!"
    return text.decode("utf-8")    

def main():
    try:
        opts, args = getopt.getopt(sys.argv[1:], "i:o:", ["help", "input"])
    except getopt.GetoptError as err:
        # print help information and exit:
        print str(err)  # will print something like "option -a not recognized"
        usage()
        sys.exit(2)
    
    infile = outfile = None;
    for o, a in opts:
        if o in ("-i", "--input"):
            infile = a
            outfile = infile+".mod.docx"
        elif o in ("-h", "--help"):
            usage()
            sys.exit()
        else:
            assert False, "unhandled option"
    
    if infile == None or outfile == None:
        print ""
        print "Error: Not all expected arguments are passed!!"
        print ""
        usage()
        sys.exit(2)

    document = Document(infile)
    for p in document.paragraphs:
        for run in p.runs:
            print "\n<run>"
            if isinstance(run.font.name, basestring):
                if run.font.name in english_fonts:
                    print_run(run)
                elif run.font.name == "Adhawin-Tamil":
                    print_run(run)
                    run.text = convert_adhawintamil(run.text);
                else:
                    print_run(run)
                    run.text = convert_bamini(run.text)
            elif run.font.name == None and run.style.name == "Default Paragraph Font":
                print_run(run)
                run.text = convert_bamini(run.text)
            else:
                print_run(run)
            print "</run>"
            
            
    document.save(infile+"-modified.docx")


if __name__ == "__main__": 
    main()