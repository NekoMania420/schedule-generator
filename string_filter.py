#!/usr/bin/env python
# -*- coding: UTF-8 -*-

from bs4 import BeautifulSoup
import codecs

def string_filter(html):
    item_list = []

    soup = BeautifulSoup(html.read())
    for each in soup.find_all("td"):
        if not each.string in ["\n", None]:
            item_list.append(each.string.strip())

    with codecs.open("filtered.txt", "w", "utf-8-sig") as f:
        for i, each in enumerate(item_list):
            if i%8 in [2, 5, 6, 7]:
                f.write("%s\n" % each)

        f.close()
