import pdfplumber
import pandas as pd 
import numpy as np 
import re 
import glob
import os
from tqdm.auto import tqdm 


sample_result = {}


for current_pdf_file in tqdm(glob.glob(r"C:\Users\Terence\OneDrive\Work\Python Scripts\oilreportPDFreader\PDF\*.pdf")):
    with pdfplumber.open(current_pdf_file) as my_pdf:
        filename = os.path.basename(current_pdf_file)
        
        try:
            p0 = my_pdf.pages[0]
            text = p0.extract_text().splitlines()
            rating = str([s for s in text if "Oil/Unit Rating" in s]).split()[-3]
            
            sample_result[filename] = rating 

        except Exception as e: 
            #print(getattr(e, 'message', repr(e)))
            #print(getattr(e, 'message', str(e)))
            sample_result[filename] = getattr(e, 'message', repr(e))


'''
keys_list = list(sample_result)
values = sample_result.values() 
values_list = list(values)
print(keys_list[0])
print(values_list[0])
'''

df = pd.DataFrame.from_dict(sample_result, orient='index', columns=['SampleResult']).rename_axis('FileName').reset_index()
print(df.head())
df.to_excel('OilSampleResults.xlsx', index=False)
