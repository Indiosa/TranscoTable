# Created on:       21.03.2022
# Author:           Marcin Konopka
# Last change on:   29.05.2022

import pandas as pd
import os
import tkinter
import pyodbc
import datetime
from tkinter.scrolledtext import ScrolledText

class Scenario:
    def __init__(self, master):
        self.master = master
        
        self.rootPath = 'C:\\Users\\marci\\Documents\\Marcin\\SGH\\' # <-- main folder with the located script
        self.transcoPath = 'C:\\Users\\marci\\Documents\\Marcin\\SGH\\TranscoTable\\' # <-- output folder for the HTML and Excel files
        
        self.versionFile = open(self.rootPath + 'versionInfo.txt','r')
        self.versionInfo = self.versionFile.read()
        self.versionFile.close()
        
        ##SQL CONNECTION
        try:
            self.sql_conn = pyodbc.connect('Driver={SQL Server};'
                              'SERVER=DELL-5570\KONOPKA;'
                              'Database=REFERENCE_BI;'
                              'Trusted_Connection=yes;')
            
            self.sqlConnectionInfo = tkinter.Label(self.master, text='Successful connection to the SQL database.', font=("Courier", 8, 'bold'), fg='green')
            self.sqlConnectionInfo.place(x=10, y=10)
        except:
            self.sqlConnectionInfo = tkinter.Label(self.master, text='Failed connection to the SQL database.', font=("Courier", 8, 'bold'), fg='red')
            self.sqlConnectionInfo.place(x=10, y=10)

        # Scenario choosing:
        self.frameGenerate = tkinter.LabelFrame(self.master, text='Choose the way you want to create the TranscoTable pages:', height=50, width=360)
        self.frameGenerate.place(x=10, y=30)
        
        self.generateOptions = tkinter.IntVar()
        self.generateOptions.set(0)

        self.rb_1 = tkinter.Radiobutton(self.frameGenerate, text='By MME', variable = self.generateOptions, command = self.MMEdataArea, value=1).grid(row=0, column=1)
        self.rb_2 = tkinter.Radiobutton(self.frameGenerate, text='By DH4', variable = self.generateOptions, command = self.DH4dataArea, value=2).grid(row=0, column=2)
        self.rb_3 = tkinter.Radiobutton(self.frameGenerate, text='By Brand', variable = self.generateOptions, command = self.DataBrand, value=3).grid(row=0, column=3)
        self.rb_4 = tkinter.Radiobutton(self.frameGenerate, text='All', variable = self.generateOptions, command = self.AllTranscoGenerator, value=4).grid(row=0, column=4)
        
        # Operation frame:
        self.frameData = tkinter.LabelFrame(self.master, height=200, width=328)
        self.frameData.place(x=10, y=74)
        
        # Action button:
        self.run_button = tkinter.Button(self.master, text='Create files', font='Arial 8 bold', height=1, command=self.getCodes)
        self.run_button.place(x=10, y=278, width=80, height=26)
        
        # Reseting button:
        self.reset_button = tkinter.Button(self.master, text='Reset', command=self.resetContent)
        self.reset_button.place(x=92, y=278, width=80, height=26)
        
        # Open path button:
        self.path_button = tkinter.Button(self.master, text="Open folder", command=self.OpenFolder)
        self.path_button.place(x=174, y=278, width=80, height=26)
        
        # Close button:
        self.close_button = tkinter.Button(self.master, text="Close", command=self.master.destroy)
        self.close_button.place(x=256,y=278, width=80, height=26)
        
        # Command text_filed:
        self.command_text = tkinter.Text(self.master, height=3, width=46, font=("Courier", 8), fg='black', bd=2)
        self.command_text.place(x=10, y=310)
        
        # Version info button:
        self.versionInfo_button = tkinter.Button(self.master, text="Version info", command=self.VersionInfoWindow)
        self.versionInfo_button.place(x=10, y=362, width=80, height=26)
        
    
    def OpenFolder(self):
        os.startfile(self.transcoPath)
        
    def resetContent(self):
        self.command_text.delete("1.0", tkinter.END)
        self.command_text.config(font=("Courier", 8), fg='black')
        self.ClearFrameData()
        self.generateOptions.set(0)
    
    def eraseContent(self):
        scenario_test = self.generateOptions.get()
        if scenario_test == 1:
            self.MME_data_area.delete("1.0", tkinter.END)
        elif scenario_test == 2:
            self.DH4_data_area.delete("1.0", tkinter.END)
        else:
            pass
    
    def VersionInfoWindow(self):
        self.VersionInfo = tkinter.Toplevel(self.master)
        self.VersionInfo.geometry("500x600")
        self.VersionInfo.title("Version information")
        self.VersionInfo.iconphoto(False, guiLogo)
        self.version_info_area = tkinter.scrolledtext.ScrolledText(self.VersionInfo, width = 58, height = 36, font = ("Courier", 10))
        self.version_info_area.grid(column = 0, pady=10, padx=10)
        self.version_info_area.insert(tkinter.INSERT, self.versionInfo)
        self.version_info_area.configure(state='disabled')

    def ClearFrameData(self):
       for widgets in self.frameData.winfo_children():
           widgets.destroy()
    
    def MMEdataArea(self):
        self.ClearFrameData()
        self.clear_button = tkinter.Button(self.frameData, text='Erase data', command=self.eraseContent, width=12)
        self.clear_button.place(x=2, y=2)
        self.MME_data_area = tkinter.scrolledtext.ScrolledText(self.frameData, width=10, height=10, font = ("Courier", 10))
        self.MME_data_area.place(x=2, y=30)
        self.MME_data_area.focus() 
    
    def DH4dataArea(self):
        self.ClearFrameData()
        self.clear_button = tkinter.Button(self.frameData, text='Erase data', command=self.eraseContent, width=12)
        self.clear_button.place(x=100, y=2)
        self.DH4_data_area = tkinter.scrolledtext.ScrolledText(self.frameData, width=10, height=10, font = ("Courier", 10))
        self.DH4_data_area.place(x=100, y=30)
        self.DH4_data_area.focus()
        
        
    def DataBrand(self):
        self.ClearFrameData()        
        self.dataBrandOptions = []
        for d in range(10):
            self.brands = tkinter.StringVar(value = "")
            self.dataBrandOptions.append(self.brands)
        
        self.cb_01 = tkinter.Checkbutton(self.frameData, text='Brand1', variable=self.dataBrandOptions[0], onvalue='B1', offvalue='').place(x=140,y=4)
        self.cb_02 = tkinter.Checkbutton(self.frameData, text='Brand2', variable=self.dataBrandOptions[1], onvalue='B2', offvalue='').place(x=140,y=24)
        self.cb_03 = tkinter.Checkbutton(self.frameData, text='Brand3', variable=self.dataBrandOptions[2], onvalue='B3', offvalue='').place(x=140,y=44)
        self.cb_04 = tkinter.Checkbutton(self.frameData, text='Brand4', variable=self.dataBrandOptions[3], onvalue='B4', offvalue='').place(x=140,y=64)
        
    def AllTranscoGenerator(self):
        self.ClearFrameData()
        self.all_label = tkinter.Label(self.frameData, text = 'All the pages will be created.')
        self.all_label.config(font=("Courier", 12))
        self.all_label.place(x=10, y=10)
        self.warningAll_label = tkinter.Label(self.frameData, text = 'Warning: This operation can take a long time due to the large number of pages to be generated at the same time!')
        self.warningAll_label.config(font=("Courier", 8), fg='red', wraplength=300, justify='left')
        self.warningAll_label.place(x=10, y=30)
        
    def getCodes(self):
        self.sql_query = ''
        start_time = datetime.datetime.now()
        self.command_text.delete("1.0", tkinter.END)
        self.command_text.config(font=("Courier", 8), fg='black')
        self.command_text.insert(tkinter.INSERT, 'Initializing...')
        self.master.update()
        
        scenario = 0
        scenario = self.generateOptions.get()
        if(scenario == 0):
            self.command_text.delete("1.0", tkinter.END)
            self.command_text.config(font=("Courier", 8), fg='black')
            self.command_text.insert(tkinter.INSERT, 'Please choose the way you want to create\nthe TranscoTable pages.')
            self.command_text.config(font=("Courier", 8), fg='red')
        else:
            #### MANAGING BY MME LIST:
            if(scenario == 1):
                self.input_data = self.MME_data_area.get('1.0', tkinter.END)
                
                if len(self.input_data) == 1:
                    self.command_text.delete("1.0", tkinter.END)
                    self.command_text.config(font=("Courier", 8), fg='blue')
                    self.command_text.insert(tkinter.INSERT, 'Please insert at least one MME code.')
                    self.sql_query = ''
                else:
                    self.input_data_list = list(set(self.input_data.split('\n')))
                    self.MMEsql = []
                    for k in self.input_data_list:
                        if k == '':
                            continue
                        self.MME_to_query = "'" + k + "'"
                        self.MMEsql.append(self.MME_to_query)
                            
                    self.MMEsql.sort()
                    self.MME_sql_query = "Select distinct DH4 \nFROM [v_Transco_Table] \nWHERE [MME] in (" + ','.join(self.MMEsql) + ");"
                    
                    self.codes_temp = pd.read_sql(self.MME_sql_query, self.sql_conn)
                    self.codes_list = self.codes_temp.stack().tolist()
                    
                    if not self.codes_list:
                        self.command_text.delete("1.0", tkinter.END)
                        self.command_text.config(font=("Courier", 8), fg='red')
                        self.command_text.insert(tkinter.INSERT, 'Entered codes do not exist in the database\nor DH4 mapping is missing.\nPlease check the entered data.')
                    else:
                        self.DH4sql = []
                        for m in self.codes_list:
                            if m == '':
                                continue
                            self.y = "'" + m + "'"
                            self.DH4sql.append(self.y)
                            
                        self.DH4sql.sort()
                        self.sql_query = "Select * \nFROM [v_Transco_Table] \nWHERE [DH4] in (" + ','.join(self.DH4sql) + ");"
                    
            #### MANAGING BY DH4 LIST:
            elif(scenario == 2):
                self.input_data = self.DH4_data_area.get('1.0', tkinter.END)

                if len(self.input_data) == 1:
                    self.command_text.delete("1.0", tkinter.END)
                    self.command_text.config(font=("Courier", 8), fg='blue')
                    self.command_text.insert(tkinter.INSERT, 'Please insert at least one DH4 code.')
                    self.sql_query = ''
                else:
                    self.input_data_list = list(set(self.input_data.split('\n')))
                    self.DH4sql = []
                    for k in self.input_data_list:
                        if k == '':
                            continue
                        self.DH4_to_query = "'" + k + "'"
                        self.DH4sql.append(self.DH4_to_query)
                            
                    self.DH4sql.sort()
                    self.sql_query = "Select * \nFROM [v_Transco_Table] \nWHERE [DH4] in (" + ','.join(self.DH4sql) + ");"
                    
            elif(scenario == 3):
                self.values = [var.get() for var in self.dataBrandOptions if var.get()]
                if (len(self.values)) == 0:
                    self.command_text.delete("1.0", tkinter.END)
                    self.command_text.config(font=("Courier", 8), fg='blue')
                    self.command_text.insert(tkinter.INSERT, 'Please choose at least one brand.')
                else:
                    self.brand_sql = []
                    for k in self.values:
                        if k == '':
                            continue
                        self.brand_to_query = "'" + k + "'"
                        self.brand_sql.append(self.brand_to_query)
                    
                    self.sql_query = "Select * \nFROM [v_Transco_Table] \nWHERE [BrandID] in (" + ','.join(self.brand_sql) + ");"
    
            elif(scenario == 4):
                self.sql_query = 'Select * FROM [v_Transco_Table]'
            
            ########################################### HTML CREATION ###################################################################
            
            if self.sql_query == '':
                pass
            else:
                self.codes = pd.read_sql(self.sql_query, self.sql_conn)
                if self.codes.empty:
                    self.command_text.delete("1.0", tkinter.END)
                    self.command_text.config(font=("Courier", 8), fg='red')
                    self.command_text.insert(tkinter.INSERT, 'The values used returned an empty results.\nPlease check the entered data.')
                else:
                    self.codes.replace([None], '---', inplace=True)
            
                    self.codesSorted = self.codes.sort_values(by=['DH4', 'Ind'])
                    self.DH4_list = []
                    self.DH4_list = sorted(set(self.codesSorted['DH4'].values.tolist()))
        
                    for count, i in enumerate(self.DH4_list, start = 1):
                        self.command_text.delete("1.0", tkinter.END)
                        self.command_text.config(font=("Courier", 8), fg='black')
                        self.command_text.insert(tkinter.INSERT, 'Creating DH4 = %s (%s of %s)' % (i, count, len(self.DH4_list)))
                        self.master.update_idletasks()
                        self.DH4TT = self.codesSorted[self.codesSorted['DH4'] == i]
                        self.brandID = ''
                        self.brandID = self.DH4TT.iloc[0,3]
                        self.rowCount = self.DH4TT.shape[0]
                        self.columnCount = len(self.DH4TT.columns)
                        self.c = [0,1]
                        for j in range(5, self.columnCount):
                            if self.DH4TT.iloc[:,j].sum() == self.rowCount * '---':
                                self.c
                            else:
                                self.c = self.c + [j]
                        self.DH4TT_final = self.DH4TT.iloc[:,self.c]
                        self.colNumber = len(self.c)
                        self.subsNumber = int((self.colNumber / 2) - 1) ## clients column number in HTML and Excel
                        self.finalColNumber = self.subsNumber + 2 ## final column numbers in HTML
                        class TranscoTable():
                            def __init__(self):
                                self.table = ''
                                self.x = 0
                                self.header = ''
                                self.data = ''
                                
                            def page(self, brandID):
                                self.table = \
'''<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" "http://www.w3.org">
<html>
    <head>
      <meta content="text/html; charset=ISO-8859-1" http-equiv="content-type">
      <title>TranscoTable for DH4 %s</title>
      <link href="%s_TranscoTable.css" rel="stylesheet" type="text/css">
    </head>
    <body>
        <table>
            <tbody>
                %s
            </tbody>
        </table>
        <br>
        <div class="TTdiv">
            Download this TranscoTable in <a title="download Transco in Excel" target="_blank" 
                                      href="%s"><Strong>Excel</strong></a> format.
        </div>
        <br>
        <p style="font-size: 14px; font-family: Calibri">
            Last update: <b>%s</b>
        </p>
    </body>
</html>
''' % (i, brandID, self.table, 'TranscoTable_' + str(i) + '.xlsx', today)
                        
                            def thead(self,headers,colNumber):
                                if colNumber > 2:
                                    self.header = \
        '''<thead>
                    <tr>
                        <th style="width: 60px; height: 32px" class="TTheader">MME</th>
                        <th style="width: 300px" class="TTheader">Commercial product name</th>
                        <th style="width: 60px" class="TTheader">%s</th>
                    </tr>
                </thead>''' % ('</th>\n\t\t\t\t\t\t<th style="width: 60px" class="TTheader">'.join(headers))
                                    self.x = len(headers)
                                else:
                                    self.header = \
        '''<thead>
                    <tr>
                        <th style="width: 60px; height: 32px" class="TTheader">MME</th>
                        <th style="width: 300px" class="TTheader">Commercial product name</th>
                        <th style="width: 60px" class="TTheader">NO LOCAL</th>
                    </tr>
                </thead>'''
                      
                            def tbody(self,data,colNumber,subsNumber,finalColNumber):
                                    if colNumber > 2:
                                        for v in data:
                                            self.data += \
                        '''\n\t\t\t\t\t<tr>
                        <td class="TTfirstColumn">%s</td>
                        <td class="TTnameColumn">%s</td>\n''' % (v[0],v[1])
                                            nn = ''
                                            for a in range(subsNumber):
                                                zz = '\t\t\t\t\t\t<td title="%s" class="TTcodification">%s</td>\n' % (v[a+finalColNumber].replace(';','&#013;'),v[a+2].replace(';','<br>'))
                                                nn += zz
                                            self.data += nn + '\t\t\t\t\t</tr>'
                                    else:
                                        for v in data:
                                            self.data += \
                                '''\n\t\t\t\t\t<tr>
                        <td class="TTfirstColumn">%s</td>
                        <td class="TTnameColumn">%s</td>
                        <td title="no local codification for this product" class="TTcodification">no local</td>\n\t\t\t\t\t</tr>''' % (v[0],v[1])
                                            self.data += '\t\t\t\t'                    
                            def TTshow(self,brandID):
                                self.table = self.header + self.data
                                self.page(brandID)
                                return self.table
                    
                        #final HTML creation
                        TT = TranscoTable()
                        TT.thead(self.DH4TT_final.iloc[[0],2:self.finalColNumber], self.colNumber)
                        TT.tbody(self.DH4TT_final.values.tolist(), self.colNumber, self.subsNumber, self.finalColNumber)
                        file = open(self.transcoPath + str(i) + ' TranscoTable.html','w')
                        file.write(TT.TTshow(self.brandID))
                        file.close()
                        ##putting dataframe into the Excel file
                        file_name = transcoPath + 'TranscoTable_' + str(i) + '.xlsx'
                        sheet_name = 'TranscoTable_' + str(i)
                        writer = pd.ExcelWriter(path = file_name, engine='xlsxwriter', mode ='w')
                        excelMainOut = self.DH4TT_final.iloc[:,0:self.finalColNumber]
                        excelMainOut.to_excel(writer, sheet_name=sheet_name, startrow = 1, index = False)
                        ##Today's date in the first row
                        workbook  = writer.book
                        worksheet = writer.sheets[sheet_name]
                        worksheet.write(0, 0, 'Last update: ' + today,
                                        workbook.add_format({'bold': True, 'size': 10}))
                        ##header and first column format:
                        header_format = workbook.add_format({'bold': True, 'fg_color': '#cccfff', 'border': 1})
                        first_col_format = workbook.add_format({'bold': True, 'fg_color': '#e6e6e6', 'border': 1,
                                                                'align': 'center', 'font_name': 'Courier New'})
                        for self.finalColNumber, value in enumerate(excelMainOut.columns.values):
                            worksheet.write(1, self.finalColNumber, value, header_format)
                        worksheet.set_column('A:A', 1, first_col_format)
                            
                        ##column width set
                        worksheet.set_column(0, 0, 6)
                        worksheet.set_column(1, 1, 60)
                        writer.save()
                        self.master.update()
                    
                    duration = datetime.datetime.now() - start_time
                    durationSeconds = duration.days * 24*60*60 +  duration.seconds
                    self.sql_query = ''
                    final_status = 'All the files have been generated!\nOperation takes %s seconds.\nTotal files generated: %s.' \
                                    % (durationSeconds, len(self.DH4_list) * 2)
                
                    self.command_text.delete("1.0", tkinter.END)
                    self.command_text.config(font=("Courier", 8), fg='green')
                    self.command_text.insert(tkinter.INSERT, final_status)
            
version = '[v 3.4]'
rootPath = 'C:\\Users\\marci\\Documents\\Marcin\\SGH\\' # <-- main folder with the located script
transcoPath = 'C:\\Users\\marci\\Documents\\Marcin\\SGH\\TranscoTable\\' # <-- output folder for the HTML and Excel files
today = datetime.date.today().strftime("%d.%m.%Y")

TTgui = tkinter.Tk()
TTgui.geometry("360x400")
TTgui.eval('tk::PlaceWindow . center')
TTgui.title("Transco Table creator " + version)
guiLogo = tkinter.PhotoImage(file = rootPath + 'TT_logo.png')
TTgui.iconphoto(False, guiLogo)

Scenario(TTgui)
TTgui.mainloop()