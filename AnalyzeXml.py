#! /usr/local/bin/python
# -*- coding: utf-8 -*-
#

import os
import sys
import getopt
import xml.etree.ElementTree as ET

from XmlExtract import extract_xml
from DocxUnicodeConv import convert_bamini


# Globals...

pending_tbi_vowel = None

# Functions....
def usage():
    print ""
    print "This tool extracts the xml file from a word document passed"
    print "as argument"
    print ""
    print "usage: ./XmlExtract.py -i <input file>"
    print ""


def ReplaceTextInXml(tb_obj):
    global pending_tbi_vowel
    tb_text =  tb_obj.text
    print tb_text

    if tb_text.encode("utf-8") == '¿':
        pending_tbi_vowel = tb_text
        tb_obj.text = ""
    elif pending_tbi_vowel != None:
        tb_text = tb_text[:1]+pending_tbi_vowel+tb_text[1:]
        tb_obj.text = convert_bamini(tb_text)
        pending_tbi_vowel = None
    else:
        tb_obj.text = convert_bamini(tb_text)


def ParseAndReplaceTextInXml(filepath):
    xmlfile = extract_xml(filepath)
    tree = ET.parse(xmlfile)
    root = tree.getroot()
    for i in range(len(root[0])):
        print "i = "+str(i)+" tag: "+root[0][i].tag
        for j in range(len(root[0][i])):
            print "j = "+str(j)+" tag: "+root[0][i][j].tag
            for k in range(len(root[0][i][j])):
                print "k = "+str(k)+" tag: "+root[0][i][j][k].tag
                for l in range(len(root[0][i][j][k])):
                    print "l = "+str(l)+" tag: "+root[0][i][j][k][l].tag
                    if root[0][i][j][k][l].tag.split('}')[-1] == 't':
                        ReplaceTextInXml(root[0][i][j][k][l])
                    for m in range(len(root[0][i][j][k][l])):
                        print "m = "+str(m)+" tag: "+root[0][i][j][k][l][m].tag
                        if root[0][i][j][k][l][m].tag.split('}')[-1] == 't':
                            ReplaceTextInXml(root[0][i][j][k][l][m])
                        for n in range(len(root[0][i][j][k][l][m])):
                            print "n = "+str(n)+" tag: "+root[0][i][j][k][l][m][n].tag
                            if root[0][i][j][k][l][m][n].tag.split('}')[-1] == 't':
                                ReplaceTextInXml(root[0][i][j][k][l][m][n])
                            for o in range(len(root[0][i][j][k][l][m])):
                                print "o = "+str(o)+" tag: "+root[0][i][j][k][l][m][n][o].tag
                                if root[0][i][j][k][l][m][n][o].tag.split('}')[-1] == 't':
                                    ReplaceTextInXml(root[0][i][j][k][l][m][n][o])
                                for p in range(len(root[0][i][j][k][l][m][n][o])):
                                    print "p = "+str(p)+" tag: "+root[0][i][j][k][l][m][n][o][p].tag
                                    if root[0][i][j][k][l][m][n][o][p].tag.split('}')[-1] == 't':
                                        ReplaceTextInXml(root[0][i][j][k][l][m][n][o][p])
                                    for q in range(len(root[0][i][j][k][l][m][n][o][p])):
                                        print "q = "+str(q)+" tag: "+root[0][i][j][k][l][m][n][o][p][q].tag
                                        if root[0][i][j][k][l][m][n][o][p][q].tag.split('}')[-1] == 't':
                                            ReplaceTextInXml(root[0][i][j][k][l][m][n][o][p][q])
                                        for r in range(len(root[0][i][j][k][l][m][n][o][p][q])):
                                            print "r = "+str(r)+" tag: "+root[0][i][j][k][l][m][n][o][p][q][r].tag
                                            if root[0][i][j][k][l][m][n][o][p][q][r].tag.split('}')[-1] == 't':
                                                ReplaceTextInXml(root[0][i][j][k][l][m][n][o][p][q][r])
                                            for s in range(len(root[0][i][j][k][l][m][n][o][p][q][r])):
                                                print "s = "+str(s)+" tag: "+root[0][i][j][k][l][m][n][o][p][q][r][s].tag
                                                if root[0][i][j][k][l][m][n][o][p][q][r][s].tag.split('}')[-1] == 't':
                                                    ReplaceTextInXml(root[0][i][j][k][l][m][n][o][p][q][r][s])
                                                for t in range(len(root[0][i][j][k][l][m][n][o][p][q][r][s])):
                                                    print "t = "+str(t)+" tag: "+root[0][i][j][k][l][m][n][o][p][q][r][s][t].tag
                                                    if root[0][i][j][k][l][m][n][o][p][q][r][s][t].tag.split('}')[-1] == 't':
                                                        ReplaceTextInXml(root[0][i][j][k][l][m][n][o][p][q][r][s][t])
                                                    for u in range(len(root[0][i][j][k][l][m][n][o][p][q][r][s][t])):
                                                        print "u = "+str(t)+" tag: "+root[0][i][j][k][l][m][n][o][p][q][r][s][t][u].tag
                                                        if root[0][i][j][k][l][m][n][o][p][q][r][s][t][u].tag.split('}')[-1] == 't':
                                                            ReplaceTextInXml(root[0][i][j][k][l][m][n][o][p][q][r][s][t][u])
    
                                                            
                    
    
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

    ParseAndReplaceTextInXml(infile)

if __name__ == "__main__": 
    main()
