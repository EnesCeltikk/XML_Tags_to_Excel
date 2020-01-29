import pandas as pd
import xml.etree.ElementTree as ET
import os
import xlrd
from collections import Counter
import re
import xlsxwriter

folder_path=('') #Enter here your XML files folder
dosyalar=[]
#if os.path.exists('output.xlsx'):
#    os.remove('output.xlsx')

workbook=xlsxwriter.Workbook('output.xlsx')
worksheet=workbook.add_worksheet()
workbook.close()



for dosya in os.listdir(folder_path):
    if '.DS_Store' not in dosya: #If your OS is not macOS, you can delete that IF statement
        dosyalar.append(dosya)
dosyalar=sorted(dosyalar, key=lambda x:int(os.path.splitext(x)[0]))

for i in range(len(dosyalar)):
    tree = ET.parse('/Users/enes/Desktop/Xml/xml_files/'+str(i)+'.xml')
    root = tree.getroot()
    lkey=[]

    def rij(elem,level,tags,rtag,mtag,keys,rootkey,data,lkey):
        otag=mtag
        mtag=elem.tag
        mtag=mtag[mtag.rfind('}')+1:]
        tags.append(mtag)
        if level==1:
            rtag=mtag
            if elem.keys() is not None:
                mkey=[]
                if len(elem.keys())>1:
                    for key in elem.keys():
                        if 'maindoc/UBL-Invoice-2.1.xsd' not in elem.attrib.get(key): #I'm working on XML files with UBL Standarts. That if statement is unnecessary
                            mkey.append(elem.attrib.get(key))
                            rootkey=mkey
                        else:
                            mkey.append('-')
                            rootkey=mkey
                else:
                    for key in elem.keys():
                        if 'maindoc/UBL-Invoice-2.1.xsd' not in elem.attrib.get(key): #I'm working on XML files with UBL Standarts. That if statement is unnecessary
                            rootkey=elem.attrib.get(key)
                        else:
                            mkey.append('-')
                            rootkey=mkey
        else:
            if elem.keys() is not None:
                    mkey=[]
                    lkey=[]
                    for key in elem.keys():
                        if len(elem.keys())>1:
                            if 'maindoc/UBL-Invoice-2.1.xsd' not in elem.attrib.get(key): #I'm working on XML files with UBL Standarts. That if statement is unnecessary
                                mkey.append(elem.attrib.get(key))
                                keys=(mkey)
                            else:
                                mkey.append("-")
                                keys=mkey
                        else:
                            for key in elem.keys():
                                if 'maindoc/UBL-Invoice-2.1.xsd' not in elem.attrib.get(key): #I'm working on XML files with UBL Standarts. That if statement is unnecessary
                                    keys=elem.attrib.get(key)
                                    lkey=key
                                else:
                                    keys=("-")
                                    lkey=key


        if elem.text is not None:
            if '\n    'not in elem.text:
                data.append([rtag,otag,mtag,lkey,keys,elem.text])
        else:
                data.append([rtag,otag,mtag,lkey,keys,''])

                #print(data)
        level+=1
        for chil in elem.getchildren():
                data = rij(chil, level,tags,rtag,mtag, keys,rootkey,data,lkey)
        level-=1
        mtag=elem.tag
        mtag=mtag[mtag.rfind('}')+1:]
        tags.remove(mtag)
        return data


    data = rij(root,0,[],'','', [],[],[],lkey)

    writer=pd.ExcelWriter('output.xlsx',engine='openpyxl',mode='a')

    df=pd.DataFrame(data, columns=['RootTag','Tag','Keys','Key Names','Key Values','Text',])
    df.to_excel(writer,sheet_name='Sheet'+str(i+2)) #Name of first sheet is 'Sheet2'. Because when excel file created, 'Sheet1' is created and we cant override on it
    writer.save()
