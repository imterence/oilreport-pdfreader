import pdfplumber
import pandas as pd 
import numpy as np 
import glob
import os
from tqdm.auto import tqdm 
import xlsxwriter


workbook = xlsxwriter.Workbook('OilSampleResults.xlsx')
sheet = workbook.add_worksheet()
sample_result = {}


for current_pdf_file in tqdm(glob.glob(r"C:\Users\Terence\OneDrive\Work\Python Scripts\oilreportPDFreader\PDF\*.pdf")):
    with pdfplumber.open(current_pdf_file) as my_pdf:
        filename = os.path.basename(current_pdf_file)        
        
        try:
            p0 = my_pdf.pages[0]
            text = p0.extract_text().splitlines()
                        
            equip_no = str([s for s in text if "Equip.no.:" in s]).split()[2]            
            #Handle non-WO_NO values. 
            if str([s for s in text if "Work Order Number" in s]).split()[-2].isdigit():
                wo_no = str([s for s in text if "Work Order Number" in s]).split()[-2]
            else:
                wo_no = "WO not given" 
            sample_no = str([s for s in text if "Sample Number" in s]).split()[-2]            
            rating = str([s for s in text if "Oil/Unit Rating" in s]).split()[-3]
            
            #list of values
            sample_data = []
            sample_data.append(equip_no)
            sample_data.append(wo_no)
            sample_data.append(sample_no) 
            sample_data.append(rating)             
                        
            sample_result[filename] = sample_data

        except Exception as e: 
            #print(getattr(e, 'message', repr(e)))
            #print(getattr(e, 'message', str(e)))
            sample_result[filename] = ["Wrong Format"]
            

sheet.write(0,0,"FILENAME")
sheet.write(0,1,"EQUIP_NO")
sheet.write(0,2,"WO_NO")
sheet.write(0,3,"SAMPLE_NO")
sheet.write(0,4,"RATING")

row = 1
for key in sample_result.keys():
    sheet.write(row,0,key)
    sheet.write_row(row, 1, sample_result[key])
    row+=1

workbook.close()

