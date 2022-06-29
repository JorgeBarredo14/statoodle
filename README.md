![alt text](https://github.com/JorgeBarredo14/statoodle/statoodle.png?raw=true)

Hello!
This is the repository for Statoodle - The external Learning Analytics and statistics tool for Moodle!

The process for creating your own executable in Windows 10 is as follows!!

1) Download all the files in the same folder.

2) Be sure you installed PyInstaller previously! If you didn't, do this:
   cd Users\#YourUser\AppData\Local\Programs\Python\Python37\Scripts or move to the folder where Python is installed
   pip install pyinstaller
  
3) Then, execute the following commands in order to generate your executable:
   cd Users\#YourUser\AppData\Local\Programs\Python\Python37\Scripts or move to the folder where Python is installed
   pyinstaller ThePathRouteWhereYouDownloadedTheFiles\statoodle.py -n="Statoodle" --onefile --noconsole --icon "icon.ico" --add-data "*.png;." --hidden-import=’matplotlib’;’seaborn’;’pandas’;’fpdf’

READY, STEADY, GO! Enjoy the tool!
