import pandas as pd
import re

def cleanName(name):
    name = str(name).lower()
    name = re.sub(r'^adgeek_', '', name)
    name = re.sub(r'\(.*?\)', '', name)
    suffixes = [
        r'_全屏',
        r'_蓋版.*',
        r'_蓋板.*',
        r'蓋版.*',
        r'蓋板.*',
        r'_impressive.*',
        r'_置底.*',
        r'_滿版.*',
        r'_native.*',
        r'_320h.*',
        r'_interstitial.*',
        r'_文中.*',
        r'_靜態.*',
        r'客製.*'
    ]
    for suffix in suffixes:
        name = re.sub(suffix, '', name)
    name = re.sub(r'[_\s]+', '', name)
    return name.strip()

excelDf = pd.read_excel('files/AdisonAdgeek.xlsx')
excelSelected = excelDf[['Name', 'Request']]

csvDf = pd.read_csv('files/LookerAdgeek.csv')
csvSelected = csvDf[['Placement', 'Requests']]

excelSelected = excelSelected.copy()
csvSelected = csvSelected.copy()

excelSelected['cleanName'] = excelSelected['Name'].apply(cleanName)
csvSelected['cleanName'] = csvSelected['Placement'].apply(cleanName)

matchedCsvNames = set()

excelData = {row['cleanName']: {'name': row['Name'], 'request': row['Request']} 
             for _, row in excelSelected.iterrows()}

csvGrouped = csvSelected.groupby('cleanName').agg({
    'Placement': list,
    'Requests': list
}).reset_index()

matchedData = []
summaryData = []

for _, excelRow in excelSelected.iterrows():
    excelCleanName = excelRow['cleanName']
    matchingCsv = csvGrouped[csvGrouped['cleanName'] == excelCleanName]
    
    if not matchingCsv.empty:
        placements = matchingCsv.iloc[0]['Placement']
        requests = matchingCsv.iloc[0]['Requests']
        matchedCsvNames.add(excelCleanName)
        
        totalCsvRequests = sum(requests)
        
        summaryData.append({
            'excelName': excelRow['Name'],
            'excelRequest': excelRow['Request'],
            'totalCsvRequests': totalCsvRequests,
            'difference': totalCsvRequests - excelRow['Request'] if pd.notna(excelRow['Request']) else None,
            'numberOfPlacements': len(placements),
            'cleanName': excelCleanName
        })
        
        for placement, request in zip(placements, requests):
            matchedData.append({
                'excelName': excelRow['Name'],
                'csvPlacement': placement,
                'csvRequests': request,
                'excelRequest': excelRow['Request'],
                'cleanName': excelCleanName
            })
    else:
        summaryData.append({
            'excelName': excelRow['Name'],
            'excelRequest': excelRow['Request'],
            'totalCsvRequests': 0,
            'difference': 0 - excelRow['Request'] if pd.notna(excelRow['Request']) else None,
            'numberOfPlacements': 0,
            'cleanName': excelCleanName
        })
        
        matchedData.append({
            'excelName': excelRow['Name'],
            'csvPlacement': None,
            'csvRequests': None,
            'excelRequest': excelRow['Request'],
            'cleanName': excelCleanName
        })

unmatchedCsv = csvSelected[~csvSelected['cleanName'].isin(matchedCsvNames)]
unmatchedData = []

unmatchedGrouped = unmatchedCsv.groupby('cleanName').agg({
    'Placement': list,
    'Requests': list
}).reset_index()

for _, row in unmatchedGrouped.iterrows():
    cleanName = row['cleanName']
    placements = row['Placement']
    requests = row['Requests']
    totalRequests = sum(requests)
    
    summaryData.append({
        'excelName': None,
        'excelRequest': None,
        'totalCsvRequests': totalRequests,
        'difference': None,
        'numberOfPlacements': len(placements),
        'cleanName': cleanName
    })
    
    for placement, request in zip(placements, requests):
        unmatchedData.append({
            'placement': placement,
            'csvRequests': request,
            'excelRequest': None,
            'cleanName': cleanName
        })

resultDf = pd.DataFrame(matchedData)
unmatchedDf = pd.DataFrame(unmatchedData)
summaryDf = pd.DataFrame(summaryData)

resultDf = resultDf.sort_values(['excelName', 'csvRequests'], na_position='last')
unmatchedDf = unmatchedDf.sort_values('placement')
summaryDf = summaryDf.sort_values('excelName', na_position='last')

print("\nSample of matches (grouped by Excel Name):")
for name in resultDf['excelName'].unique()[:5]:
    group = resultDf[resultDf['excelName'] == name]
    print(f"\nExcel Name: {name}")
    if group['csvPlacement'].isna().all():
        print(f"  No matching CSV entries (Excel Request: {group['excelRequest'].iloc[0]})")
    else:
        for _, row in group.iterrows():
            if pd.notna(row['csvPlacement']):
                print(f"  Placement: {row['csvPlacement']}")
                print(f"  CSV Requests: {row['csvRequests']}")
                print(f"  Excel Request: {row['excelRequest']}")

print("\nSample of Request Summaries:")
print(summaryDf.head())

resultDf.to_csv('output/matched_data.csv', index=False)
unmatchedDf.to_csv('output/unmatched_csv_data.csv', index=False)
summaryDf.to_csv('output/request_summary.csv', index=False)

print("\nFiles exported:")
print("- matched_data.csv: Detailed matches between Excel and CSV data")
print("- unmatched_csv_data.csv: CSV entries without Excel matches")
print("- request_summary.csv: Summary of Request totals and differences")