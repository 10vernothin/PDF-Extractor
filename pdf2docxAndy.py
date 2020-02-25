import os
import subprocess
import python_packages.docx2csv as docx2csv

# THESE CONSTANTS WILL NEED TO BE EDITED ON A COMPUTER-BY-COMPUTER BASIS
mswordpath = '"C:\Program Files (x86)\Microsoft Office\\root\Office16\WINWORD.exe"'  # Location of WINWORD
startuppath = '"C:\\Users\AL_is\AppData\Roaming\Microsoft\Word\Startup\"'  # Location of WINWORD STARTUP PATH
default_save_folder = 'C:\\Users\AL_is\OneDrive\Documents'  # Location of THE DEFAULT FOLDER THAT MICROSOFT WORD SAVES YOUR FILES IN

# DEFINING FOLDER PATHS
fdr_rt_path = ''.join([os.path.realpath(__file__), "\.."])
pdfpath = ''.join([fdr_rt_path, "\\resources\pdfs"])
original_macro_path = ''.join([fdr_rt_path, "\\resources\macros\PdfPromptDisable.dotm"])
results_path = ''.join([fdr_rt_path, "\\resources\\results\\"])
files = os.listdir(pdfpath)

# COPYING THE TEMPLATE FILE INTO THE TRUSTED FOLDER
subprocess.call("xcopy {0} {1} /Y".format(original_macro_path, startuppath))
new_macro_path = ''.join([startuppath[:-1], '\\PdfPromptDisable.dotm"'])

# DISABLING PDF CONVERSION PROMPT
subprocess.call("{0} /mDisableAllWarnings /mStartTimer -q".format(mswordpath, pdfpath), shell=True)

# DEFINING DEFAULT SAVE FILE PATH
defaultfilepath = ''.join(['"', default_save_folder, "\\temporary_pdf_docx.docx", '"'])


for file in files:
    if file.find('~$') == -1:

    #SET PATH VARIABLES
        oldfilepath = ''.join([pdfpath, "\\", file])
        dumppath = ''.join([fdr_rt_path, "\\resources\\dump"])
        newfolderpath = ''.join([results_path, "ResFolder_", file[:-4]])
        newfilepath = ''.join([newfolderpath, "\\", "Results_", file[:-4], ".docx"])
        defaultdumpfilepath = ''.join([fdr_rt_path, "\\resources\\dump\\temporary_pdf_docx.docx"])
        dumpfilepath = ''.join([defaultdumpfilepath, "\..\\","Results_", file[:-4], ".docx"])

        try:

    # OPENING PDF IN WORD THEN SAVING IT AS DOCX (Will not work with some pdfs) (IN WINDOWS MODE)
            subprocess.call("{0} /mSaveDoc /mSetInvisible /mStartTimer -q /f {1}".format(mswordpath, oldfilepath), shell=True)

            try:

        # CREATING A FILE-SPECIFIC FOLDER IF NONE EXISTS
                subprocess.call("mkdir {}".format(newfolderpath), shell= True)

            except:
                    print ("FOLDER ALREADY EXISTS")

    # COPYING FILE TO DUMPPATH FOR EXCEPTION HANDLING
            subprocess.call("copy /y {0} {1}".format(
                defaultfilepath,
                dumppath), shell = True)
            subprocess.call("move /Y {0} {1}".format(
                defaultdumpfilepath,
                dumpfilepath), shell = True)

    # MOVING FILE TO THE NEW FOLDER IN /results/
            subprocess.call("move /Y {0} {1}".format(
                defaultfilepath,
                newfilepath), shell = True)

        except:
            print ("WARNING: " + file +" NOT TRANSFERRED.")
            continue

        try:

    # CODE TO EXTRACT DOCX TO CSV TABLES
            docx2csv.extract(newfilepath,'csv', True)
            docx2csv.extract(newfilepath,'xls', True)

        except:
             print ("EXTRACTION not complete for " + file+ ". Attempting CSV creation at Dump folder.")

             try:

        # ATTEMPTTING TO CONVERT THE FILE IN THE DUMP FOLDER
                docx2csv.extract(dumpfilepath,'csv', True)
                docx2csv.extract(dumpfilepath,'xls', True)
                print("EXTRACTION COMPLETE FOR" + file + ", CHECK /dump FOLDER FOR THE CORRESPONDING CSV TABLES")

             except:

        # SKIPPING FILE WHEN ALL FAILS
                print("ERROR: DUMP FILE EXTRACTION FAILED. SKIPPING FILE.")
                continue

    # DELETING DUMP FOLDER DOCX FILE
        subprocess.call("del {0}".format(dumpfilepath), shell = True)

# DISABLING PDF CONVERSION PROMPT
subprocess.call("{0} /mEnableAllWarnings /mStartTimer -q".format(mswordpath, pdfpath), shell=True)

# DELETING THE TEMPLATE FILE FROM THE TRUSTED FOLDER
subprocess.call("del {0}".format(new_macro_path), shell = True)