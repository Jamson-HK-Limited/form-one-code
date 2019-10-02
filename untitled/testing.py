import win32com.client as wincli
import os, os.path
from tkinter import *
import glob
from PyPDF2 import PdfFileWriter, PdfFileReader, PdfFileMerger
import pandas as pd
import numpy as np
import re
import win32api
import win32print
from datetime import datetime


# Global Variables
PO_Number= ""
WorkOrder_Number = ""
Finish_Date = ""

def get_main_dataframe():
    # xlsm_file = pd.ExcelFile('C:/Users/Wing Fung/Documents/Drawing/'
    #                          'JAMSON_LEUNG DWG/Carpet Status/Production Information.xlsm')
    xlsm_file = pd.ExcelFile(os.getcwd()[])
    required_df = xlsm_file.parse('Main')
    return required_df


def search_info():
    PO_Number_Info = PO_Number.get()
    WorkOrder_Number_Info = WorkOrder_Number.get()
    Finish_Date_Info = Finish_Date.get()
    req_row = main_data[main_data['PO number']==PO_Number_Info]
    if req_row.empty:
        print("Please Enter Valid PO Number")
    else:
        print(req_row)
        req_PO_num = req_row['PO number'].values[0]
        req_PO_Date = req_row['PO date'].values[0]
        req_PN = req_row['P/N'].values[0]
        req_DT = req_row['DT form number'].values[0]
        req_mat = req_row['Material'].values[0]
        req_qty = req_row['Qty'].values[0]
        req_batch = req_row['Batch number'].values[0]
        req_part_page = req_PN[:req_PN.index('-')]
        req_part_num = req_PN[req_PN.index('-') + 1:]
        if len(req_part_num) == 6:
            req_PNx = req_PN[:-3] + 'XXX'
        else:
            req_PNx = req_PN


        # Part I of Form One
        # region
        # Parts for JHL-2 Work Order Form
        # region
        ExcelApp = wincli.Dispatch("Excel.Application")
        ExcelApp.Visible = True
        JHL2_filename = 'Resources_For_Automated_Program\Forms\JHL-2 Product Work Order Form.xlsx'
        JHL_2_wb = ExcelApp.Workbooks.Open(os.path.abspath(JHL2_filename))
        JHL_2_WS = JHL_2_wb.Worksheets('Sheet1')
        JHL_2_WS.Cells(4, 7).Value = WorkOrder_Number_Info
        JHL_2_WS.Cells(4, 7).Font.Underline = False
        JHL_2_WS.Cells(5 ,7).Value = req_DT
        JHL_2_WS.Cells(9, 2).Value = "PO 0" +str(req_PO_num)
        JHL_2_WS.Cells(9, 2).VerticalAlignment = 2   # put in middle
        JHL_2_WS.Cells(9, 2).HorizontalAlignment = 3 # put in middle
        JHL_2_WS.Cells(9, 4).Value = "Carpet"
        JHL_2_WS.Cells(9, 4).VerticalAlignment = 2   # put in middle
        JHL_2_WS.Cells(9, 4).HorizontalAlignment = 3 # put in middle
        kit_mat_df =  pd.ExcelFile('Resources_For_Automated_Program\Kit_Material.xlsx').parse(
            req_part_page)
        req_kit_mat_row = pd.DataFrame()
        if req_mat != 'KIT':
            JHL_2_WS.Cells(9, 5).Value = req_mat
        else:
            req_kit_mat_row = kit_mat_df[kit_mat_df['PN']==req_PNx].T.drop('PN')
            mat_row = 9
            for row in req_kit_mat_row.iterrows():
                if row[1][0] == 1:
                    JHL_2_WS.Cells(mat_row, 5).Value = row[0]
                    JHL_2_WS.Cells(mat_row, 5).Font.Size = 12
                    JHL_2_WS.Cells(mat_row, 5).VerticalAlignment = 2  # put in middle
                    JHL_2_WS.Cells(mat_row, 5).HorizontalAlignment = 3  # put in middle
                    mat_row +=1
        JHL_2_WS.Cells(9, 6).Value = req_PN
        JHL_2_WS.Cells(9, 8).Value = req_qty
        JHL_2_WS.Cells(9, 8).VerticalAlignment = 2   # put in middle
        JHL_2_WS.Cells(9, 8).HorizontalAlignment = 3 # put in middle

        drawing_row = 9
        drawingddt_ref_df = pd.ExcelFile('Resources_For_Automated_Program\DrawingDDT_Page_Reference.xlsx').parse(
            req_part_page)
        if drawingddt_ref_df[drawingddt_ref_df['P/N']==req_PNx]['Drawings'].empty:
            print('No Required Drawing Reference Found')
        else:
            req_drawing_page = drawingddt_ref_df[drawingddt_ref_df['P/N'] == req_PNx]['Drawings'].values[0]
            req_drawing_page = [int(s) for s in req_drawing_page.split(',')]
            drawing_ver_df = pd.ExcelFile('Resources_For_Automated_Program\Drawing_Versions.xlsx').parse(
                req_part_page)
            req_drawing_ref_df = pd.DataFrame()
            for page in req_drawing_page:
                # print(drawing_ver_df[drawing_ver_df['Page']==page])
                req_drawing_ref_df = req_drawing_ref_df.append(drawing_ver_df[drawing_ver_df['Page'] == page])
            req_drawing_ref_df = req_drawing_ref_df.reset_index(drop= True)
            iterate_df = req_drawing_ref_df
            no_of_drawingpage = str(drawing_ver_df.count()['Page'])
            while not iterate_df.empty:
                page_string = ""
                cur_ver = iterate_df[iterate_df.index==0]['Rev'].values[0]
                temp_df = iterate_df[iterate_df['Rev'] == cur_ver]
                iterate_df = iterate_df[iterate_df['Rev'] != cur_ver].reset_index(drop= True)
                for row in temp_df.iterrows():
                    if page_string=="":
                        page_string = page_string + str(row[1][0])
                    else:
                        page_string = page_string + ", "+ str(row[1][0])
                if req_drawing_ref_df.count()[0] <= 20:
                    JHL_2_WS.Cells(drawing_row, 7).Value = str(req_part_page) + '               REV:' + cur_ver
                    JHL_2_WS.Cells(drawing_row, 7).Font.Size = 12
                    drawing_row += 1
                    JHL_2_WS.Cells(drawing_row, 7).Value = 'SHT '+ page_string + ' OF ' + no_of_drawingpage
                    JHL_2_WS.Cells(drawing_row, 7).Font.Size = 12
                    drawing_row += 1
                else:
                    if drawing_row == 9:
                        JHL_2_WS.Cells(drawing_row, 7).Value = str(req_part_page)
                        JHL_2_WS.Cells(drawing_row, 7).Font.Size = 12
                        drawing_row += 1
                    req_sentence = 'SHT ' + page_string + ' OF ' + no_of_drawingpage + '     REV:' + cur_ver
                    if len(req_sentence) <= 35:
                        JHL_2_WS.Cells(drawing_row, 7).Value = req_sentence
                        JHL_2_WS.Cells(drawing_row, 7).Font.Size = 12
                        drawing_row += 1
                    else:
                        JHL_2_WS.Cells(drawing_row, 7).Value = 'SHT ' + page_string
                        JHL_2_WS.Cells(drawing_row, 7).Font.Size = 12
                        drawing_row += 1
                        JHL_2_WS.Cells(drawing_row, 7).Value = 'OF ' + no_of_drawingpage + ' REV:' + cur_ver
                        JHL_2_WS.Cells(drawing_row, 7).Font.Size = 12
                        drawing_row += 1

        form_date = datetime.strftime(datetime.strptime(req_batch[1:-3], '%y%m%d'),'%d/%m/%Y')
        JHL_2_WS.Cells(22, 4).Value = form_date
        JHL_2_WS.Cells(22, 4).VerticalAlignment = 2   # put in middle
        JHL_2_WS.Cells(22, 4).HorizontalAlignment = 3 # put in middle

        JHL_2_wb.SaveAs(os.getcwd()+'\JHL-2_Temp.xlsx')
        JHL_2_wb.Close()
        ExcelApp.Visible = False
        ExcelApp.Quit
        print_file('JHL-2_Temp.xlsx')
        os.remove('JHL-2_Temp.xlsx')
        # endregion

        # Parts for printing required PO
        # region
        req_PO_num_string = req_PO_num.astype(str)[4:6]
        req_PO_Date_string = req_PO_Date.strftime('%y%m%d')
        req_POForm_Search = 'PO_' + req_PO_Date_string + "_" + req_PO_num_string + 'XX'
        POForm_filename = 'Resources_For_Automated_Program\PO\\' + req_POForm_Search +".pdf"
        # os.startfile(POForm_filename)
        print_file(POForm_filename)
        # endregion

        # Parts for printing required DDT Form
        # region
        req_DT_Search = "DDT" + req_DT[9:12]
        DDT_filename = 'Resources_For_Automated_Program\DDT Form\\' + req_DT_Search + ".pdf"

        if drawingddt_ref_df[drawingddt_ref_df['P/N']==req_PNx]['DDT Page'].empty:
            # os.startfile(DDT_filename)
            print_file(DDT_filename)
        else:
            req_ddt_page = int(drawingddt_ref_df[drawingddt_ref_df['P/N'] == req_PNx]['DDT Page'].values[0])
            temp_pdf = PdfFileReader(DDT_filename)
            if req_ddt_page== 1 or req_ddt_page== temp_pdf.getNumPages():
                SplitDDT_filename = pdf_splitter(DDT_filename,[1,temp_pdf.getNumPages()])
            else:
                SplitDDT_filename = pdf_splitter(DDT_filename, [1,req_ddt_page, temp_pdf.getNumPages()])
            print_file(SplitDDT_filename)
            os.remove(SplitDDT_filename)
        # endregion

        # Parts for printing required Drawings
        # region
        Drawing_filename = 'Resources_For_Automated_Program\Drawings\\' + req_part_page + ".pdf"
        # os.startfile(Drawing_filename)
        SplitDrawing_filename = pdf_splitter(Drawing_filename,req_drawing_page)
        print_file(SplitDrawing_filename)
        os.remove(SplitDrawing_filename)
        # endregion

        # endregion

        # Part II of Form One
        # region
        # Parts for JHL-7 Material Requisition Form
        # region
        req_mat_size =""
        # endregion

        # Parts for Carpet Materials
        # region
        # endregion

        # endregion

        # Part III of Form One
        # region
        # Parts for printing required Wire
        # region
        if req_mat != 'KIT':
            if req_mat == '004540-020164-09B':
                req_wire_Search = '線44794452'
            elif req_mat == '6560-092107-07A':
                req_wire_Search = '線4217'
            elif req_mat == 'L25WR001464400':
                req_wire_Search = '線45014502'
            elif req_mat == 'FT01948200LS701':
                req_wire_Search = '線3506'
            elif req_mat == 'FT01607252LS100':
                req_wire_Search = '線GTUVTB'
            Wire_filename = 'Resources_For_Automated_Program\Wires\\' + req_wire_Search +".pdf"
            # os.startfile(Wire_filename)
            print_file(Wire_filename)
        else:
            for row in req_kit_mat_row.iterrows():
                if row[1][0] == 1:
                    if row[0] == '004540-020164-09B':
                        req_wire_Search = '線44794452'
                    elif row[0] == '6560-092107-07A':
                        req_wire_Search = '線4217'
                    elif row[0] == 'L25WR001464400':
                        req_wire_Search = '線45014502'
                    elif row[0] == 'FT01948200LS701':
                        req_wire_Search = '線3506'
                    elif row[0] == 'FT01607252LS100':
                        req_wire_Search = '線GTUVTB'
                    Wire_filename = 'Resources_For_Automated_Program\Wires\\' + req_wire_Search + ".pdf"
                    # os.startfile(Wire_filename)
                    print_file(Wire_filename)
        # endregion

        # Parts for JHL-1 Product Material Batch Control Form
        # region
        ExcelApp = wincli.Dispatch("Excel.Application")
        ExcelApp.Visible = True
        JHL_1_filename = 'Resources_For_Automated_Program\Forms\JHL-1 Product Material Batch Control Form.xlsx'
        JHL_1_wb = ExcelApp.Workbooks.Open(os.path.abspath(JHL_1_filename))
        JHL_1_WS = JHL_1_wb.Worksheets('Sheet1')
        JHL_1_WS.Cells(5, 4).Value = "PO 0" +str(req_PO_num)
        JHL_1_WS.Cells(5, 4).VerticalAlignment = 2   # put in middle
        JHL_1_WS.Cells(5, 4).HorizontalAlignment = 3 # put in middle
        JHL_1_WS.Cells(6, 4).Value = req_DT
        JHL_1_WS.Cells(6, 4).VerticalAlignment = 2   # put in middle
        JHL_1_WS.Cells(6, 4).HorizontalAlignment = 3 # put in middle
        JHL_1_WS.Cells(7, 4).VerticalAlignment = 2  # put in middle
        JHL_1_WS.Cells(7, 4).HorizontalAlignment = 3  # put in middle
        JHL_1_WS.Cells(8, 4).Value = req_PN
        JHL_1_WS.Cells(8, 4).VerticalAlignment = 2  # put in middle
        JHL_1_WS.Cells(8, 4).HorizontalAlignment = 3  # put in middle
        JHL_1_WS.Cells(9, 4).Value = req_part_page
        JHL_1_WS.Cells(9, 4).VerticalAlignment = 2  # put in middle
        JHL_1_WS.Cells(9, 4).HorizontalAlignment = 3  # put in middle
        JHL_1_WS.Cells(5, 11).Value = req_batch
        JHL_1_WS.Cells(5, 11).VerticalAlignment = 2  # put in middle
        JHL_1_WS.Cells(5, 11).HorizontalAlignment = 3  # put in middle
        JHL_1_WS.Cells(6, 11).Value = form_date
        JHL_1_WS.Cells(6, 11).VerticalAlignment = 2  # put in middle
        JHL_1_WS.Cells(6, 11).HorizontalAlignment = 3  # put in middle
        JHL_1_WS.Cells(7, 11).Value = req_qty
        JHL_1_WS.Cells(7, 11).VerticalAlignment = 2  # put in middle
        JHL_1_WS.Cells(7, 11).HorizontalAlignment = 3  # put in middle

        # 表Part
        if req_kit_mat_row[req_kit_mat_row[0] == 1].count()[0]>=3:
            JHL_1_WS.rows(20).Insert()
            JHL_1_WS.Cells(20, 1).Value = '8.'
            JHL_1_WS.Cells(21, 1).Value = '9.'
            JHL_1_WS.Range('E20:F20').MergeCells = True
            JHL_1_WS.Range('H20:I20').MergeCells = True
            JHL_1_WS.Range('J20:K20').MergeCells = True
            JHL_1_WS.rows(26).Delete()
            JHL_1_WS.Range('B13:L21').VerticalAlignment = 2  # put in middle
            JHL_1_WS.Range('B13:L21').HorizontalAlignment = 3  # put in middle

        info_row = 13
        wires_info_df = pd.ExcelFile('Resources_For_Automated_Program\Wires_Information.xlsx').parse(
            'Sheet1')
        for row in req_kit_mat_row.iterrows():
            if row[1][0] == 1:
                JHL_1_WS.Cells(info_row, 2).Value = 'COVERING'
                JHL_1_WS.Cells(info_row,4).Value = row[0]
                ## input other info
                ####
                ###
                ####
                info_row += 1
                cur_wire_info = wires_info_df[wires_info_df['Req_Mat']==row[0]].reset_index(drop=True)
                for x in range(0,2):
                    trow = cur_wire_info.iloc[x]
                    JHL_1_WS.Cells(info_row, 2).Value = trow['Description']
                    JHL_1_WS.Cells(info_row, 4).Value = trow['Wires']
                    JHL_1_WS.Cells(info_row, 5).Value = trow['Batch']
                    JHL_1_WS.Cells(info_row, 7).Value = trow['Manufacturer']
                    if req_mat_size <=80:
                        JHL_1_WS.Cells(info_row, 8).Value = '1EA'
                    else:
                        JHL_1_WS.Cells(info_row, 8).Value = '2EA'
                    if not trow['GRN'] == trow['GRN']:
                        JHL_1_WS.Cells(info_row, 10).Value = ""
                    else:
                        JHL_1_WS.Cells(info_row, 10).Value = trow['GRN']
                    JHL_1_WS.Cells(info_row, 12).Value = trow['BurnTest_Cert']
                    info_row += 1
                    del trow
        finish_day = datetime.strftime(datetime.strptime(Finish_Date_Info, '%d/%m/%y'), '%d/%m/%Y')
        JHL_1_WS.Cells(28, 4).Value = finish_day
        JHL_1_WS.Cells(28, 4).VerticalAlignment = 2  # put in middle
        JHL_1_WS.Cells(28, 4).HorizontalAlignment = 3  # put in middle
        JHL_1_WS.Cells(28, 12).Value = finish_day
        JHL_1_WS.Cells(28, 12).VerticalAlignment = 2  # put in middle
        JHL_1_WS.Cells(28, 12).HorizontalAlignment = 3  # put in middle

        # endregion

        # endregion


def pdf_splitter(path, req_pages):
    fname = os.path.splitext(os.path.basename(path))[0]

    pdf = PdfFileReader(path)
    pdf_writer = PdfFileWriter()
    for each in req_pages:
        pdf_writer.addPage(pdf.getPage(each-1))

    output_filename = '{}_splitted.pdf'.format(
        fname)

    with open(output_filename, 'wb') as out:
        pdf_writer.write(out)

    return output_filename

# need amend
def pdf_merger(output_path, input_paths):
    merger = PdfFileMerger()
    file_handles = []

    for path in input_paths:
        merger.append(path)

    with open(output_path, 'wb') as fileobj:
        merger.write(fileobj)


def print_file(filename):
    win32api.ShellExecute(
        0,
        "printto",
        filename,
        '"%s"' % win32print.GetDefaultPrinter(),
        ".",
        0
    )


def create_userform():
    # Create userform for input
    # region
    master = Tk()
    master.geometry("250x350")
    master.title("Form One Creator")
    heading = Label(text="Form One Creator", bg="grey", fg="black", width="500", height="3")
    heading.pack()

    PO_Number_text = Label(master, text="PO Number: ", )
    PO_Number_text.place(x=15, y=70)
    WorkOrder_Number_text = Label(master, text="Work Order Number: ", )
    WorkOrder_Number_text.place(x=15, y=140)
    Finish_Date_text = Label(master, text="Finish Date: ", )
    Finish_Date_text.place(x=15, y=210)


    PO_Number = IntVar()
    WorkOrder_Number = StringVar()
    Finish_Date = StringVar()


    PO_Number_Entry = Entry(textvariable=PO_Number, width="30")
    PO_Number_Entry.place(x=15, y=100)
    WorkOrder_Number_Entry = Entry(textvariable=WorkOrder_Number, width="30")
    WorkOrder_Number_Entry.place(x=15, y=170)
    Finish_Date_Entry = Entry(textvariable=Finish_Date, width="30")
    Finish_Date_Entry.place(x=15, y=240)

    Search = Button(master, text="Search PO", width="30", height="2", command=search_info, bg="grey")
    Search.place(x=15, y=300)
    master.mainloop()
    # endregion


# Background main dataframe settings
# region
main_data = get_main_dataframe()
main_data_undone = main_data[main_data['Status'] == 0]
main_data_undone = main_data_undone.reset_index(drop= True)
main_data_done = main_data[main_data['Status'] == 1]
main_data_done = main_data_done.reset_index(drop= True)
# endregion

create_userform()


# for searching if contain file
# glob.glob(r'*\*\*.pdf')













## ------------------------------Tips of Excel stuff---------------------------------
#region
# ExcelApp = wincli.Dispatch("Excel.Application")
# ExcelApp.Visible = True

# ExcelWorkbook = ExcelApp.Workbooks.Add()
#
# ExcelWrkSht = ExcelWorkbook.Worksheets.Add()
#
# Excelrng = ExcelWrkSht.Range("A1:A10")
# Excelrng.Value = 1


##resources

## For Python
#https://stackoverflow.com/questions/19616205/running-an-excel-macro-via-python
#https://www.youtube.com/watch?v=Jd2PtDV5mL0

## For Tkinter
#https://www.python-course.eu/tkinter_entry_widgets.php
#https://www.youtube.com/watch?v=xH4JOEJ5Uc0

#tk.Label(master, text="First Name").grid(row=0)
#tk.Label(master, text="Last Name").grid(row=1)
# e1 = tk.Entry(master)
#e1.grid(row=0, column=1)
#e2.grid(row=1, column=1)

#tk.Button(master,
#          text='Quit',
#          command=master.quit).grid(row=3,
#                                    column=0,
#                                    sticky=tk.W,
#                                    pady=4)
#tk.Button(master,
#          text='Show', command=show_entry_fields).grid(row=3,
#                                                       column=1,
#                                                       sticky=tk.W,
#                                                       pady=4)

# print("First Name: %s\nLast Name: %s" % (e1.get(), e2.get()))
# https://morvanzhou.github.io/tutorials/python-basic/tkinter/2-01-label-button/
# endregion
# just copy need check later
# def find_reference_page():
#     pdfFileObj = open(r'C:\Users\Craig\RomeoAndJuliet.pdf', mode='rb')
#     pdfReader = PdfFileReader(pdfFileObj)
#     number_of_pages = pdfReader.numPages
#
#     pages_text = []
#     words_start_pos = {}
#     words = {}
#
#     searchwords = ['romeo', 'juliet']
#
#     with open('FoundWordsList.csv', 'w') as f:
#         f.write('{0},{1}\n'.format("Sheet Number", "Search Word"))
#         for word in searchwords:
#             for page in range(number_of_pages):
#                 print(page)
#                 pages_text.append(pdfReader.getPage(page).extractText())
#                 words_start_pos[page] = [dwg.start() for dwg in re.finditer(word, pages_text[page].lower())]
#                 words[page] = [pages_text[page][value:value + len(word)] for value in words_start_pos[page]]
#             for page in words:
#                 for i in range(0, len(words[page])):
#                     if str(words[page][i]) != 'nan':
#                         f.write('{0},{1}\n'.format(page + 1, words[page][i]))
#                         print(page, words[page][i])

# https://www.youtube.com/watch?v=nw5_oSz4RJk

