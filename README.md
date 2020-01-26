# unzip_and_ship
I wrote this code to help a colleague in automating one of his tasks. Everyday he downloads a file from company server which contains information about payment handlers, then downloads project information for all the current projects. Due to some reasons the system does not output all the information about a project in one single excel file, instead it outputs a zip file with an unknown number of files. The files are named randomly.  He appends those project files together, then appends each project file into one single file. He looks up handler information from the handlers file, assigns statuses(Paid or Pending) and then proceeds to email the handlers. I designed an RPA process on UiPath to download all the files, and this python code does the rest except the email part. I will be writing a code for email soon.
The process goes as follows:
1. Download a handlers file from company server (Accomplished by UiPath)
2. Download files for all the current projects (Accomplished by UiPath)
3. Find latest handlers file in downloads and read it
4. The projects files are not available as one single file for each project, instead each project has a .zip file with an unknown number of files in it named randomly so the code finds latest .zip file for each project, unzips it into relevant folders and then reads all the files in those folders irrespective of number or name of files
5. All those files are appended and handler information is looked up from handlers file
6. A 'Status' column is added and statuses are assigned according to Bill Payment Percentage
7. All this information is written into a single excel file
8. This file is later used for further analysis and email intimation to the handlers, however that part is still done manually

