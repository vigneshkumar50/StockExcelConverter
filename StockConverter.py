# -*- coding: utf-8 -*-
"""
Created on Thu Jul 23 09:07:06 2020

@author: Vigneshkumar Ganapathy
"""
import configparser
import pandas as pd
import xlsxwriter
import logging

# Logger level setup configuration
logging.getLogger().setLevel(logging.INFO)
try:
    logging.info("Fetching data values from the stock download template")
    source = pd.read_excel('Input/stock_download.xlsx')
except Exception as error:
    logging.info("The expected download template is not available in the mentioned location")
    print(error)
    exit(0)

# Read the property file.
logging.info("Fetching column values from the property file")
config = configparser.ConfigParser()
config.read('Properties/config_file.properties')
# Read the download Excel file.
# Create the new upload file and sheet      
workbook = xlsxwriter.Workbook('Output/stock_upload.xlsx', options={'nan_inf_to_errors': True})
worksheet = workbook.add_worksheet("ProductSheet")

# Add header value for upload file.    
logging.info("Creating a column header for the stock upload template file")
for key, val in config.items('UploadFileSection'):
    worksheet.write(0, int(key), val)


# Update logic for upload file    
def update_method(row_count, index_count, upload_index, download_index):
    if pd.notnull(source[config['DownloadFileSection'][download_index]][index_count]):
        worksheet.write(row_count, upload_index, source[config['DownloadFileSection'][download_index]][index_count])
    else:
        worksheet.write(row_count, upload_index, '')
    return


# Get the Handling unit details from download Excel file and update in expected upload file.
def handling_unit_details(row_count, index_count):
    worksheet.write(row_count, 0, config['DefaultValueSection']['hand_unit_pos_type'])
    update_method(row_count, index_count, 16, '11')
    update_method(row_count, index_count, 20, '9')
    update_method(row_count, index_count, 22, '12')
    update_method(row_count, index_count, 24, '12')
    worksheet.write(row_count, 21, config['DefaultValueSection']['ext_no'])
    worksheet.write(row_count, 25, config['DefaultValueSection']['hand_unit_row'])
    return


# Get the product value from download Excel file and update in expected upload file.
def product_details(row_count, index_count):
    worksheet.write(row_count, 0, config['DefaultValueSection']['product_pos_type'])
    update_method(row_count, index_count, 1, '0')
    update_method(row_count, index_count, 2, '1')
    update_method(row_count, index_count, 4, '2')
    update_method(row_count, index_count, 5, '3')
    update_method(row_count, index_count, 10, '5')
    update_method(row_count, index_count, 12, '6')
    update_method(row_count, index_count, 13, '7')
    update_method(row_count, index_count, 14, '8')
    update_method(row_count, index_count, 15, '9')
    update_method(row_count, index_count, 16, '11')
    update_method(row_count, index_count, 20, '9')
    update_method(row_count, index_count, 22, '12')
    update_method(row_count, index_count, 24, '12')
    update_method(row_count, index_count, 61, '24')
    update_method(row_count, index_count, 62, '25')
    worksheet.write(row_count, 3, config['DefaultValueSection']['owner_role'])
    worksheet.write(row_count, 11, config['DefaultValueSection']['entitled_role'])
    worksheet.write(row_count, 21, config['DefaultValueSection']['ext_no'])
    worksheet.write(row_count, 25, config['DefaultValueSection']['product_row'])
    return;


try:
    row = 1
    logging.info("Creating the stock upload template is in in progress. Please wait for a few seconds...")
    for count in source.index:
        if pd.notnull(source[config['DownloadFileSection']['0']][count]):
            handling_unit_details(row, count)
            row += 1
            product_details(row, count)
            row += 1
        else:
            break
    workbook.close()
    logging.info("The stock upload template is created successfully!!!")

except Exception as err:
    logging.info("The stock upload template is getting failed")
    print(err)
