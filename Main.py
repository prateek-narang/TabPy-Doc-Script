# -*- coding: utf-8 -*-
"""
Created on Wed Dec 13 15:28:17 2023

@author: PrateekNarang
"""

from TableauDocExtract import TableauDocument
import json

def main():
    with open("./config.json") as f:
        conf = f.read()
        conf = json.loads(conf)
        fpath = conf['tableau_path']

    obj = TableauDocument(fpath)
    obj.output_to_excel()
    obj.generate_thumbnails()

if __name__=="__main__":
    main()

