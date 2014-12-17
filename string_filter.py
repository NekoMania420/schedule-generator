#!/usr/bin/env python
# -*- coding: UTF-8 -*-

import sys
sys.path.append("bs4")

from bs4 import BeautifulSoup
import codecs


def string_filter(html):
    """Filter HTML code to use with gen.py."""

    item_list = []

    soup = BeautifulSoup(html.read())
    for each in soup.find_all("td")[33:]:
        item_list.append(each.text)

    with codecs.open("filtered.txt", "w", "utf-8-sig") as f:
        for i, each in enumerate(item_list):
            if i%16 in [4, 8, 10, 12, 14]:
                f.write("%s\n" % each)

        f.close()