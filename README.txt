IMPORTANT!!!

BEFORE EXECUTING, For Both Excel and Word:

1. Go to Options -> Trust Center
2. Click on the Trust Center button
3. Go to Macro Settings tab, and select "Enable all macros"

4. Find your Windows Word executable path and copy the value to $mswordpath in pdf2docxAndy,py. 
   It should be at or near:
   C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.exe 
   (Alternatively, You can "open file location" on Word shortcuts until you get the path)

5. Find your default save path for documents and copy it to $default_save_path in pdf2docxAndy.py. 
   For most people, it should be at:
   C:\\Users\USERNAME\Documents

6. Finally, place the pdf files in /resources/pdfs folder	
	


Disclaimer: 

Everything in the python_packages folder is not mine.

Execution: run /docs/python/pdf2docxAndy.py