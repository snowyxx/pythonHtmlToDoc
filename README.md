# pythonHtmlToDoc
user path to convert html to docx, and some other features.

I exported my site from https://www.helpdocsonline.com.

But it does not come with a index page and navigator to let me view it as a site offline.

And I need a docx file include all content of these html files.


So this tool try to: 

1. add a user guide template with index and navigator to exported html files. 
2. convert html files to a single docx file by an exact order. 

Prerequest: 

1. python 2.7 
2. Windows OS 
3. pywin32 
4. Office Word 

Install pywin32 

1. download from https://sourceforge.net/projects/pywin32/ 
or 2. pip install pypiwin32 

How to use:
 
1. login into https://www.helpdocsonline.com/login 
2. settings -- back up site, to get a zip file. 
3. unzip exported file to a folder. 
4. unzip attached pythonHtmlToDoc.zip 
6. run command : python helpIQ.py \<you unzip help folder\> 
