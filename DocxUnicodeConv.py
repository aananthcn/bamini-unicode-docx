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
from AmudhamDict import amudham_dict
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
    print "style.font = "+str(run.style.font.name)
    print "text = "+run.text

def convert_amudham(wg):
    text = wg.encode("utf-8")
    for key in amudham_dict.keys():
        text = text.replace(str(key), str(amudham_dict.get(key)))
    print "Amudham: "+text
    return text.decode("utf-8")

def convert_bamini(wg):
    text = wg.encode("utf-8")
    for key in bamini_dict.keys():
        text = text.replace(str(key), str(bamini_dict.get(key)))
    print "Bamini: "+text
    return text.decode("utf-8")

def convert_adhawintamil(wg):
    text = wg.encode("utf-8")
    for key in adhawintamil_dict.keys():
        text = text.replace(str(key), str(adhawintamil_dict.get(key)))
    print "Adhawin-Tamil: "+text
    return text.decode("utf-8")


# arg1: run object , arg2: paragraph font
def convert_runfont(run, p_font):

    r_font = run.font.name
    print_run(run)

    if r_font == "Amudham" or (r_font == None and p_font == "Amudham"):
        run.text = convert_amudham(run.text)
    elif r_font == "Adhawin-Tamil" or (r_font == None and p_font == "Adhawin-Tamil"):
        run.text = convert_adhawintamil(run.text)
    elif r_font == "Bamini" or (r_font == None and p_font == "Bamini"):
        run.text = convert_bamini(run.text)
    elif r_font in english_fonts:
        print "No conversion for "+str(r_font)
    else:
        print "UNKOWN FONT"
        run.text = convert_bamini(run.text)


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
        # Collect paragraph information
        print "\n<para>"
        paragraph_font = str(p.style.font.name)
        print "paragraph font: "+paragraph_font
        print "paragraph text: "+p.text

        # Collect run information in this paragraph
        runs = [r for r in p.runs]
        runs_len = len(runs)
        print runs_len

        # if no runs, skip conversion
        if runs_len == 0:
            print "</para>"
            continue

        # Convert single-run paragragh
        if runs_len == 1:
            print "\n<run>"
            convert_runfont(runs[0], paragraph_font)
            print "</run>"
            print "</para>"
            continue

        # Convert multi-run paragraph
        i = ref = 0
        #for run in p.runs:
        for i in range(runs_len):
            print "\n<run>"
            print runs[i].text
            print runs[i].font.name

            # take a copy of i as reference
            ref = i

            # Cat all run.text as long as they have same font
            while i+1 < runs_len and runs[i].font.name == runs[i+1].font.name:
                runs[ref].text += runs[i+1].text
                runs[i+1].clear() # clear the content of this run
                i += 1

            # Convert the concatenated run
            convert_runfont(runs[ref], paragraph_font)
            print "</run>"
        print "</para>"
            
            
    document.save(infile+"-modified.docx")


if __name__ == "__main__": 
    main()
