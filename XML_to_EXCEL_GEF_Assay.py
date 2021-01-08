
#!/usr/bin/python

import traceback, sys
import xml.dom.minidom
import numpy as np
import xlsxwriter
import math

def append_value(dict_obj, key, value):
    try:
        if key in dict_obj:
            dict_obj[key].append(value)
        else:
            dict_obj[key] = [value]
    except:
        traceback.print_exc(file=sys.stderr)

        
def read_xml_file(filename):
    try:
        doc = xml.dom.minidom.parse(filename)
        well_tag = doc.getElementsByTagName('well')

        print('number of wells read:  ', len(well_tag))

        Organized_values_dict = dict() # Dictionary to store values

        well_val = np.zeros(0)
        time_val = np.zeros(0)
        i = 0
        
        for item in well_tag:
            wellname  = item.attributes['wellName'].value
            Fluorescence_Data = item.getElementsByTagName('rawData')
            timedata = item.getElementsByTagName('timeData')
            well_val = Fluorescence_Data[0].firstChild.nodeValue
            time_val = timedata[0].firstChild.nodeValue
            
            # Join the time series values at one place
            i = i+1
            if i == 1:
                wellname_i = wellname
                Organized_values_dict['Time'] = [time_val]
            if i > 1 and wellname_i == wellname:
                Organized_values_dict['Time'].append(time_val)
            
            append_value(Organized_values_dict, wellname, well_val) # call function
                
        
        return Organized_values_dict
    except:
        traceback.print_exc(file=sys.stderr)

def create_output(filename):
    try:
        Organized_values_dict = read_xml_file(filename) # call function
        excelname = filename[:-4] + '.xlsx' # make the filename end with .xlsx
        
        row = 0
        col = 0
        workbook = xlsxwriter.Workbook(excelname)
        worksheet = workbook.add_worksheet()
        for key in Organized_values_dict.keys():
            row = 0
            col = col + 1
            worksheet.write(row, col, key)
            for item in Organized_values_dict[key]:
                item1 = np.char.split(item)
                item2 = np.array(item1, ndmin=1)
                for val in item2[0]:
                    row += 1
                    worksheet.write(row, col, float(val))
        workbook.close()
    except:
        traceback.print_exc(file=sys.stderr)

if __name__ == '__main__':
    filename = input('Enter the .xml filename (e.g. filename.xml)')
    create_output(filename)

    
