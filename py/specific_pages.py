from pprint import pprint

import pymupdf
import pandas as pd
import openpyxl


res = []

doc = pymupdf.open(f"files/input/file.pdf") #  open document
doc.select([118, 119, 120, 121])  # select pages
doc.save("files/input/changed.pdf")

doc = pymupdf.open(f"files/input/changed.pdf") #  open document
for i, page in enumerate(doc):
    tabs = page.find_tables(strategy="lines_strict") #  locate and extract any tables on page
    s = tabs[0].to_markdown()
    print(type(s))
    with open(f"files/output/{i} changed.md", "w", encoding="UTF-8") as f:
        f.write(s)
