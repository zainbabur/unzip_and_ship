# unzip_and_ship
I wrote this code to help a colleague in automating one of his tasks. Everyday he downloads a file from company server which contains information about payment handlers, then downloads project information for all the current projects. Due to some reasons the system does not output all the information about a project in one single excel file, instead it outputs a zip file with an unknown number of files. The files are named randomly.  He appends those project files together, then appends each project file into one single file. He looks up handler information from the handlers file, assigns statuses(Paid or Pending) and then proceeds to email the handlers. I designed an RPA process on UiPath to download all the files, and this python code does the rest except the email part. I will be writing a code for email soon.
