#! /usr/bin/python

from BaminiDict import bamini_dict

text = "vd;W Ngrp mopAk;&gt; Gfo;&gt; nghUs;&gt; NjLtjw;F my;y vd;W njsptu Gupe;Jnfhz;L&gt; elf;f&gt; ,iwNfl;L&gt; mo Ntz;Lk;"
copy = text
for key in bamini_dict.keys():
    copy = copy.replace(str(key), str(bamini_dict.get(str(key))))

print "Original: "+text
print "Modified: "+copy

