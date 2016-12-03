#! /usr/local/bin/python


import os
import sys
import getopt
import zipfile



# Functions....
def usage():
    print ""
    print "This tool extracts the xml file from a word document passed"
    print "as argument"
    print ""
    print "usage: ./XmlExtract.py -i <input file>"
    print ""


def extract_xml(infile):
    doc = zipfile.ZipFile(infile)

    with open(doc.extract("word/document.xml", "/tmp/")) as tempfile:
        tempstr = tempfile.read()

    with open(infile+".xml", "w+") as tempfile:
        tempfile.write(tempstr)

    doc.close()
    tempfile.close()



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
        elif o in ("-h", "--help"):
            usage()
            sys.exit()
        else:
            assert False, "unhandled option"
    
    if infile == None:
        print ""
        print "Error: Not all expected arguments are passed!!"
        print ""
        usage()
        sys.exit(2)

    extract_xml(infile)

if __name__ == "__main__": 
    main()
