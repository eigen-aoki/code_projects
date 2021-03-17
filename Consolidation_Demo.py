import os
import pandas as pd
import openpyxl
from tqdm import tqdm

#set CWD
os.chdir(r'C:\Users\Eigen Aoki\Documents\Scripts\Consolidation')

#Read Consolidated Sheet Template as Dictionary
ws_dict = pd.read_excel('Consolidation.xlsx')

#convert sheet to datafarme. this is the master sheet I want to append to
mod_df = ws_dict
mod_list = [mod_df] #create list to be used later

#grab the workbooks/sheets from data folder
os.chdir('.\Data') #go to the data folder
print(os.getcwd()) #verify I'm in the right directory

#list of files
file_list = os.listdir()
print(file_list) #verify that the files are what I want

append_list = [] #create empty list for appending

#run a loop for all the files in the data folder
for files in tqdm(range(len(file_list))):
    try:
        file_name = file_list[files] #grab the stirng file name
        data = pd.read_excel(file_name, header = 0) #read in excel file as panda df
        data = data.reset_index(drop = True) #drop index since we dont need
        data['Source'] = file_name
        append_list.append(data) #append the dataframe to empty append list
        time.sleep(0.1)
    except:
        pass

append_list = mod_list + append_list #combine both lists with main sheet first

#concatenate into new pandas df
print('Will start appending now')
appended_df = pd.concat(append_list, axis=0, ignore_index=True, sort=False)
print('Finished appending')


#save to new excel file
print('Will start exporting to excel')
appended_df.to_excel('..\Consolidated.xlsx', sheet_name = 'test', index = False)
print('Finished')
