import os
import subprocess
import docx2csv
from docs.python.ExtractOpenXMLAndy import convert


# DEFINING FOLDER PATHS
fdr_rt_path = ''.join([os.path.realpath(__file__), "\.."])
pdfpath = ''.join([fdr_rt_path, "\\resources\pdfs"])
results_path = ''.join([fdr_rt_path, "\\resources\\results\\"])
files = os.listdir(pdfpath)


for file in files:
    if file.find('~$') == -1:

    #SET PATH VARIABLES
        pdffilepath = ''.join([pdfpath, "\\", file])
        dumppath = ''.join([fdr_rt_path, "\\resources\\dump"])
        newxmltxtfile = ''.join([fdr_rt_path, "\\resources\\xmltexts\\" + file[:-4]+".xml"])
        # try:
    #   CODE TO EXTRACT DOCX TO CSV TABLES
        convert('XML', pdffilepath, newxmltxtfile)
        # except:
        #     print ("EXTRACTION not complete for " + file+ ". Attempting CSV creation at Dump folder.")


    # DELETING DUMP FOLDER DOCX FILE
