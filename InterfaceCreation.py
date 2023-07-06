from tkinter import *
from tkinter import filedialog
import customtkinter
import os
import pandas as pd
import openpyxl
from pyepc import SGTIN
from pyepc.exceptions import DecodingError
import datetime

global store_num
global date

global epc_directory_path
global cycle_count_paths

global app
count = 0


def store_number_verification():
    global store_num
    store_num = store_entry.get()
    try:
        int(store_num)
        print("Store Number: {}".format(store_num))
        return True
    except:
        print(":: ERROR :: Store Num is not an int!")
        return False


def date_verification():
    global date
    date = date_entry.get()
    try:
        if date == "":
            return False
        date_list = date.split(".")
        if len(date_list[0]) == 4 and isinstance(int(date_list[0]), int):
            if len(date_list[1]) == 2 and isinstance(int(date_list[1]), int):
                if len(date_list[2]) == 2 and isinstance(int(date_list[2]), int):
                    print("Date: {}".format(date))
                    return True
    except:
        print(":: ERROR :: Date input is not valid!")
        return False


def import_epc_directory():
    print("Importing EPC Directory...")
    pop_up_title = "Select EPC Directory Data (.xlsx)"
    filename = filedialog.askopenfilename(initialdir="/", title=pop_up_title,
                                          filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
    global epc_directory_path
    epc_directory_path = filename
    print(epc_directory_path)


def import_cycle_count():
    print("Importing Cycle Count...")
    pop_up_title = "Select Cycle Count Data (.txt)"
    filenames = filedialog.askopenfilenames(initialdir="/", title=pop_up_title,
                                          filetypes=(("txt files", "*.txt"), ("all files", "*.*")))
    global cycle_count_paths
    cycle_count_paths = []
    for filename in filenames:
        cycle_count_paths.append(filename)
    print(cycle_count_paths)


def decodePreparation():
    epc_list = []
    for filename in cycle_count_paths:
        f = open(filename, "r")
        lines = f.readlines()
        for x in lines:
            epc_list.append(x.split('\n')[0])
        f.close()

    epc_list_no_dupe = [*set(epc_list)]
    epc_list_df = pd.DataFrame(epc_list_no_dupe, columns=['EPCs'])

    print("Preparing to Decode...")
    return epc_list_df


def decodeCycleCount(epc_list_df):

    epc_list = []
    columns = epc_list_df.columns.tolist()

    for _, i in epc_list_df.iterrows():
        for col in columns:
            epc_list.append(i[col])

    temp_epc_list = []
    for epc in epc_list:
        temp_epc_list.append(str(epc))

    epc_list = temp_epc_list

    res = list(map(''.join, epc_list))
    epc_list = [*set(res)]

    upc_list, error_epcs, error_upcs = [], [], []

    print("Decoding...")
    for epc in epc_list:
        try:
            upc_list.append(SGTIN.decode(epc).gtin)
        except DecodingError as de:
            error_epcs.append(epc)
            error_upcs.append(de)
        except TypeError as te:
            error_epcs.append(epc)
            error_upcs.append(te)

    for epc in error_epcs:
        if epc in epc_list:
            epc_list.remove(epc)

    for upc in range(len(upc_list)):
        upc_list[upc] = upc_list[upc].lstrip('0')

    return epc_list, upc_list


def create_active_epcs():
    pd.set_option('display.float_format', lambda x: f'{x:.3f}')
    df = pd.read_excel(epc_directory_path, sheet_name=0)
    active_epc_list = df['EPC'].tolist()
    active_upc_list = df['UPC'].tolist()
    active_date_list = df['Latest Date Seen'].tolist()

    temp_epc_list = []
    for epc in active_epc_list:
        temp_epc_list.append(str(epc))

    active_epc_list = temp_epc_list

    return active_epc_list, active_upc_list, active_date_list


def update_active_epcs(epc_list, upc_list, active_epc_list, active_upc_list, active_date_list):
    global date
    date_str = str(date)
    year = int(date_str[:4])
    month = int(date_str[5:7])
    day = int(date_str[8:])
    # reformatted_date = date_str[:4] + "-" + date_str[5:7] + "-" + date_str[8:] + " 00:00:00"
    reformatted_date = datetime.datetime(year, month, day)
    print(reformatted_date)
    active_epcs_updated = epc_list
    active_upcs_updated = upc_list
    active_date_updated = []
    for i in range(len(active_epcs_updated)):
        active_date_updated.append(reformatted_date)

    active_epcs_updated.extend(active_epc_list)
    active_upcs_updated.extend(active_upc_list)
    active_date_updated.extend(active_date_list)

    upc_list_match_epcs = []
    date_list_match_epcs = []

    index_dict = {}
    for i, epc in enumerate(active_epcs_updated):
        if epc not in index_dict:
            index_dict[epc] = i
    epc_list_no_dupe = [(v,k) for k, v in index_dict.items()]
    index_list, epc_list_no_dupe = zip(*epc_list_no_dupe)
    for index in index_list:
        upc_list_match_epcs.append(active_upcs_updated[index])
        date_list_match_epcs.append(active_date_updated[index])

    active_epcs_updated = epc_list_no_dupe
    active_upcs_updated = upc_list_match_epcs
    active_date_updated = date_list_match_epcs

    temp_epc_list = []
    for epc in active_epcs_updated:
        temp_epc_list.append(str(epc))

    active_epcs_updated = temp_epc_list

    return active_epcs_updated, active_upcs_updated, active_date_updated


def create_new_epcs():
    pd.set_option('display.float_format', lambda x: f'{x:.3f}')
    df = pd.read_excel(epc_directory_path, sheet_name=1)
    new_epc_list = df['EPC'].tolist()
    new_upc_list = df['UPC'].tolist()
    new_date_list = df['Date'].tolist()
    new_status_list = df['Status'].tolist()

    temp_epc_list = []
    for epc in new_epc_list:
        temp_epc_list.append(str(epc))

    new_epc_list = temp_epc_list

    return new_epc_list, new_upc_list, new_date_list, new_status_list


def update_new_epcs(epc_list, upc_list, new_epc_list, new_upc_list, new_date_list, new_status_list):
    global date
    date_str = str(date)
    year = int(date_str[:4])
    month = int(date_str[5:7])
    day = int(date_str[8:])
    # reformatted_date = date_str[:4] + "-" + date_str[5:7] + "-" + date_str[8:] + " 00:00:00"
    reformatted_date = datetime.datetime(year, month, day)
    print(reformatted_date)

    for i in range(len(new_status_list)):
        if new_status_list[i] == "New":
            new_status_list[i] = "Existing"

    new_epcs_updated = new_epc_list
    new_epcs_updated.extend(epc_list)

    new_upcs_updated = new_upc_list
    new_upcs_updated.extend(upc_list)

    new_date_updated = new_date_list
    new_status_updated = new_status_list
    for i in range(len(new_epcs_updated)):
        new_date_updated.append(reformatted_date)
        new_status_updated.append("New")

    upc_list_match_epcs = []
    date_list_match_epcs = []
    status_list_match_epcs = []

    index_dict = {}
    for i, epc in enumerate(new_epcs_updated):
        if epc not in index_dict:
            index_dict[epc] = i
    epc_list_no_dupe = [(v, k) for k, v in index_dict.items()]
    index_list, epc_list_no_dupe = zip(*epc_list_no_dupe)
    for index in index_list:
        upc_list_match_epcs.append(new_upcs_updated[index])
        date_list_match_epcs.append(new_date_updated[index])
        status_list_match_epcs.append(new_status_updated[index])

    new_epcs_updated = epc_list_no_dupe
    new_upcs_updated = upc_list_match_epcs
    new_date_updated = date_list_match_epcs
    new_status_updated = status_list_match_epcs

    temp_epc_list = []
    for epc in new_epcs_updated:
        temp_epc_list.append(str(epc))

    new_epcs_updated = temp_epc_list

    return new_epcs_updated, new_upcs_updated, new_date_updated, new_status_updated


def export_epc_directory(active_epcs, active_upcs, active_latest_date, new_epcs, new_upcs, new_date, new_status):
    temp = active_epcs
    active_epcs = []
    for epc in temp:
        active_epcs.append(str(epc))

    pd.set_option('display.float_format', lambda x: f'{x:.3f}')
    df1 = pd.DataFrame(active_epcs, columns=['EPC'])
    df2 = pd.DataFrame(active_upcs, columns=['UPC'])
    df3 = pd.DataFrame(active_latest_date, columns=['Latest Date Seen'])
    df4 = pd.DataFrame(new_epcs, columns=['EPC'])
    df5 = pd.DataFrame(new_upcs, columns=['UPC'])
    df6 = pd.DataFrame(new_date, columns=['Date'])
    df7 = pd.DataFrame(new_status, columns=['Status'])


    cc_file_name = "Store{}_ActiveAndNewEPCs_{}.xlsx".format(store_num, date)
    path = os.path.join(os.path.expanduser("~"), "Desktop", cc_file_name)  # Saves on Desktop
    str(path)

    global count
    if count > 1:
        path2 = path.split('.')
        path2.insert(-1, ' ({0}).'.format(count - 1))  # Adds file count to filename
        path = ''.join(path2)

    pd.set_option('display.float_format', lambda x: f'{x:.3f}')
    writer = pd.ExcelWriter(path, engine='xlsxwriter')

    workbook = writer.book
    worksheet1 = workbook.add_worksheet('Active EPC Directory')
    worksheet2 = workbook.add_worksheet('New EPC Directory')

    # set the format of the 'value' column to text
    text_format = workbook.add_format({'num_format': '@'})
    worksheet1.set_column('A:A', None, text_format)
    worksheet2.set_column('A:A', None, text_format)

    number_format = workbook.add_format({'num_format': '0'})
    worksheet1.set_column('B:B', None, number_format)
    worksheet2.set_column('B:B', None, number_format)

    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss'})
    worksheet1.set_column('C:C', None, date_format)
    worksheet2.set_column('C:C', None, date_format)

    print("Creating Active and New EPCs Directory File...")

    df1.to_excel(writer, sheet_name='Active EPC Directory', startrow=0, startcol=0, index=False)
    df4.to_excel(writer, sheet_name='New EPC Directory', startrow=0, startcol=0, index=False)
    df2.to_excel(writer, sheet_name='Active EPC Directory', startrow=0, startcol=1, index=False)
    df3.to_excel(writer, sheet_name='Active EPC Directory', startrow=0, startcol=2, index=False)
    df5.to_excel(writer, sheet_name='New EPC Directory', startrow=0, startcol=1, index=False)
    df6.to_excel(writer, sheet_name='New EPC Directory', startrow=0, startcol=2, index=False)
    df7.to_excel(writer, sheet_name='New EPC Directory', startrow=0, startcol=3, index=False)

    writer.save()


def submit():
    if store_number_verification() and date_verification():
        epc_list_df = decodePreparation()

        epc_list, upc_list = decodeCycleCount(epc_list_df)
        active_epc_list, active_upc_list, active_date_list = create_active_epcs()
        active_epc_list, active_upc_list, active_date_list = update_active_epcs(epc_list, upc_list, active_epc_list,
                                                                                active_upc_list, active_date_list)
        new_epc_list, new_upc_list, new_date_list, new_status_list = create_new_epcs()
        new_epc_list, new_upc_list, new_date_list, new_status_list = update_new_epcs(epc_list, upc_list, new_epc_list,
                                                                                     new_upc_list, new_date_list,
                                                                                     new_status_list)
        export_epc_directory(active_epc_list, active_upc_list, active_date_list,
                             new_epc_list, new_upc_list, new_date_list, new_status_list)
        print("Active and New EPCs exported. Press 'Quit' to quit.")
    else:
        print("Invalid Store Number or Date!")


def quit_app():
    # enumerateTest = ["ab", "abc", "abcd", "abcde", "abcdef", "abcdefg", "abcdefgh", "abcdefghi"]
    # tester = [(i, j) for i, j in enumerate(enumerateTest)]
    # print(tester)
    print("Quitting...")
    exit(1)


class InterfaceCreation:

    def __init__(self, root, w, h):
        self.root = root
        self.width = w
        self.height = h
        self.store_num = None
        self.date_input = None


    customtkinter.set_appearance_mode("Dark")
    customtkinter.set_default_color_theme("dark-blue")
    global app
    app = customtkinter.CTk()
    app.title("Active and New EPC Generator")
    app.geometry("800x600")


    '''
    Frame Creation
    '''
    main_frame = customtkinter.CTkFrame(master=app, fg_color="transparent")
    main_frame.pack(fill="both", expand=True)

    global store_entry
    store_entry = customtkinter.CTkEntry(master=main_frame, placeholder_text="Store #")

    global date_entry
    date_entry = customtkinter.CTkEntry(master=main_frame, placeholder_text="Date (YYYY.MM.DD)")

    store_entry.pack(padx=30, pady=50)
    date_entry.pack(padx=30, pady=10)

    '''
    Button Creation
    '''
    cycle_count_button = customtkinter.CTkButton(master=main_frame, text="Cycle Count (.txt)",
                                                 command=import_cycle_count)

    epc_directory_button = customtkinter.CTkButton(master=main_frame, text="EPC Directory (.xlsx)",
                                                    command=import_epc_directory)
    submit_button = customtkinter.CTkButton(master=main_frame, text="Submit", command=submit)

    quit_button = customtkinter.CTkButton(master=main_frame, text="Quit", command=quit_app)

    cycle_count_button.pack(padx=30, pady=50)
    epc_directory_button.pack(padx=30, pady=0)
    submit_button.pack(padx=30, pady=20)
    quit_button.pack(padx=30, pady=50)

    app.mainloop()