# email-attachements-downloader-script
A Python script which scans and downloads attachements from MS Outlook.

Description:
If you regularly receive letters from one addressee that contain some files (for example, daily, weekly or quarterly reports or uploads from the database), then it can be very useful for you to keep them in a structured form, for example, hierarchically by date of receipt. At the same time, looking for a letter, downloading files, then creating a folder for them and putting them in it is a very laborious routine.
This script offers to automate the process.

Note:
To solve the described problem, it is very convenient to create a sub-folder in Outlook and configure a filter by which letters from the desired addressee fall into it.
The script is configured to search for fresh incoming letters from just such a sub-folder (although this can be easily changed by indicating the main folder of incoming letters to it).

Directory structure of regularly downloaded files:
Suppose the files you get have the same name every time (this happens often).
So your directory structure become looking like this:
2020-01-03-downloaded-report
  report.docx
  archive.zip
2020-01-02-downloaded-report
  report.docx
  archive.zip
2020-01-01-downloaded-report
  report.docx
  archive.zip
report.docx  <--- files, received today are downloaded in your main downloads directory,
archive.zip  <--- therefore they are easy to distinguish

How script works:
Script algorithm can be divided into two steps.
Step 1. Create directory for yesterday's files and move them into it
Step 2. Connect to MS Outlook, scan letters in sub folder, and if find letter with predefined theme, download all attachments to folder /get_path.
As a result, you get a directory from folders structurally named by download date and the files downloaded today at the very bottom.
