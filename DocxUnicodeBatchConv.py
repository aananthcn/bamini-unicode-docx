#! /usr/local/bin/python


import os
import sys
import getopt

from DocxUnicodeConv import DocxUnicodeConv

# Functions....
def usage():
    print ""
    print "This tool invokes DocxUnicodeConv.py in batch mode based on"
    print "a file list passed as argument"
    print ""
    print "usage: ./DocxUnicodeBatchConv.py -i <batch file>"
    print ""


def extract_batchfile(infile):
    with open(infile) as batchfile:
        flist = batchfile.read().splitlines()

    batchfile.close()
    return flist



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

    flist = extract_batchfile(infile)
    for f in flist:
        path = '/'.join(f.split('/')[0:-1])
        DocxUnicodeConv(f, path)


if __name__ == "__main__": 
    main()
