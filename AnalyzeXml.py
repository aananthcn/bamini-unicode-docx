#! /usr/local/bin/python
# -*- coding: utf-8 -*-
#

import xml.etree.ElementTree as ET

def ParseAndReplaceTextInXml(filepath)
tree = ET.parse(filepath)
root = tree.getroot()
for i in range(len(root[0])):
    for j in range(len(root[0][i])):
        for k in range(len(root[0][i][j])):
            for l in range(len(root[0][i][j][k])):
                if root[0][i][j][k][l].tag.split('}')[-1] == 'Fallback':
                    for m in range(len(root[0][i][j][k][l])):
                        for n in range(len(root[0][i][j][k][l][m])):
                            for o in range(len(root[0][i][j][k][l][m][n])):
                                for p in range(len(root[0][i][j][k][l][m][n][o])):
                                    if root[0][i][j][k][l][m][n][o][p].tag.split('}')[-1] == 'textbox':
                                        for q in range(len(root[0][i][j][k][l][m][n][o][p])):
                                            for r in range(len(root[0][i][j][k][l][m][n][o][p][q])):
                                                for s in range(len(root[0][i][j][k][l][m][n][o][p][q][r])):
                                                    for t in range(len(root[0][i][j][k][l][m][n][o][p][q][r][s])):
                                                        if root[0][i][j][k][l][m][n][o][p][q][r][s][t].tag.split('}')[-1] == 't':
                                                            print "o = " + str(o)
                                                            print "i = " + str(i)
                                                            print "n = " + str(n)
                                                            print "r = " + str(r)
                                                            print root[0][i][j][k][l][m][n][o][p][q][r][s][t].text
                                                            root[0][i][j][k][l][m][n][o][p][q][r][s][t].text = "Aananth C N"
tree .write(filepath+"-modified.xml")

                                                            
                    
    
