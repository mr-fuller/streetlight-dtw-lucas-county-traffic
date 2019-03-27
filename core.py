import zipfile,fnmatch,os, csv, openpyxl, pandas as pd, xlsxwriter

rootPath = 'C:\\Users\\fullerm\\OneDrive - Toledo Metropolitan Area Council of Governments\\Documents\\Streetlight\\DTW_airport_traffic'
# unzip files
trip_list = []

os.chdir(rootPath)
extract_location = ''
pattern = '*.zip'
print('Unzipping files...')
for root, dirs, files in os.walk(rootPath):
    for filename in fnmatch.filter(files, pattern):
        
        year = filename[-23:-19]
        
        month = filename[-18:-16]
        slid = filename[-15:-11]
        if filename[0:3] == 'DTW':
            direction = 'southbound'
        else:
            direction = 'northbound'
        print(year, month, slid, direction)
        
        extract_location = os.path.join(root, os.path.splitext(filename)[0])
        
        with zipfile.ZipFile(os.path.join(root, filename),'r') as zip_ref:
            zip_ref.extractall(extract_location)
        
        os.chdir(extract_location)
        os.chdir(os.listdir(extract_location)[0])
        
        for item in os.listdir(os.getcwd()):
            counter = 0            
            if item[-15:] == f'{slid}_od_all.csv':
                trip_counts_df = pd.read_csv(os.path.abspath(item))
                
                trip_count=trip_counts_df.iloc[0,14]
                
                trip_list.append([year, month, direction, trip_count])
print('\bDone')         

#gather/spread the df for ease of chart creation
df=pd.DataFrame(trip_list,columns=['Year', 'Month','Direction','Trips'])
print('Creating Pivot Table (Converting Long to Wide)...')
pt = pd.pivot_table(df, index=['Year','Month'],columns=['Direction'])
print('\bDone')
print(f'Saving to {rootPath}...')
# df = pd.concat(trip_list,pd.Series(trip_counts))
pt.to_excel(f'{rootPath}/trip_counts_pt.xlsx',engine='xlsxwriter')
print('\bDone')
# create excel linechart?