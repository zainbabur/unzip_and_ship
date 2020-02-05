import shutil   #for file operations
import glob #to get latest file
import pandas as pd #for working on excel files
import os   #for directory operations

def get_latest_file(path):
    '''
    This function takes a path(directory + filename), makes a list of such filenames present in the
    directory and returns the latest w.r.t creation time'''
    list_of_files = glob.glob(path)
    latest_file = max(list_of_files, key=os.path.getctime)
    return latest_file


downloads_path = 'C:\\Users\\Zain Babur\\Downloads\\'
dest_path = 'C:\\Users\\Zain Babur\\Documents\\'

latest_file = get_latest_file(downloads_path+'Handler_Details_*.xls') #Get the latest Handlers information

#Reading Handlers information
handlers = pd.read_excel(latest_file, ignore_index=True, usecols=['Project No.','Shipment No.', 'Current Handler'])
#Making a unique value for lookup
handlers['Unique No.'] = handlers['Project No.'].astype(str) + handlers['Shipment No.'].astype(str)
#Dropping columns no longer needed
handlers.drop(columns=['Project No.','Shipment No.'], inplace=True)

#Getting the open project codes from a file
projects = pd.read_csv(dest_path+'Projects.csv')
projectList = projects['Project Code'].tolist()

#Making an empty dataframe
final_df = pd.DataFrame(Columns=['Project No.', 'Project Name', 'Client Name','Supplier Name', 'Shipment No.','Bill Payment Percentage','Status'])

for number in projectList:
    try: #if the files are in .zip format
        #Deleting the directory previously made to ensure latest data is saved
        dirpath = dest_path+str(number)
        if os.path.exists(dirpath):
            shutil.rmtree(dirpath)

        #finding latest project zip file and unzipping it 
        latest_file = get_latest_file(downloads_path+str(number)+'_Details_*.zip')
        shutil.unpack_archive(latest_file, dest_path+str(number), 'zip')
    except: #if the file is a single excel file
        latest_file = get_latest_file(downloads_path+str(number)+'_Details_*.xls')
        shutil.copyfile(latest_file, dest_path+str(number)+'\\'+str(number)+'.xls')
        
    #Making a list of files present in the archive we unzipped
    list_of_files = glob.glob(dest_path+str(number)+'\\*.xls')
    #Making an empty dataframe
    project_df = pd.DataFrame(Columns=['Project No.', 'Project Name', 'Client Name','Supplier Name', 'Shipment No.','Bill Payment Percentage','Status'])
    #Appending all the files of one project into one project df
    for i in list_of_files:
        temp_df = pd.read_excel(i, ignore_index = True)
        #Bill Payment Percentage is written as a string with a % sign, so we remove it and convert it to float
        temp_df['Bill Payment Percentage'] = temp_df['Bill Payment Percentage'].str.replace('%','').astype(float)/100 
        project_df = project_df.append(temp_df, ignore_index=True,sort=False)
    final_df = final_df.apppend(project_df, ignore_index=True,sort=False)
#Making a unique number for lookup
final_df['Unique No.'] = final_df['Project No.'].astype(str) + final_df['Shipment No.'].astype(str)
#Looking up handlers into projects information
final_df = final_df.merge(handlers, on='Unique No.', how='left')
#Assigning status for easy filtering
final_df.loc[final_df['Bill Payment Percentage'] == 1, 'Status'] = 'Paid'
final_df.loc[final_df['Bill Payment Percentage'] <= 1, 'Status'] = 'Pending'
#Writing into a file
final_df.to_excel(dest_path+'Payment Status.xlsx', index=False)
