import csv
import os
import shutil as pen_drive
import struct
from datetime import datetime
from math import floor
from threading import Timer
from time import sleep
import win32print
from csv2pdf import convert
from modbus_tk import modbus_rtu
from serial import Serial
from win32api import ShellExecute
import csv
import requests 
import mimetypes
import uuid
import json
import os
from datetime import datetime
import time
import shutil

ser = Serial(port="COM6", baudrate=115200, bytesize=8, parity='N', timeout=0, stopbits=1)
pid = modbus_rtu.RtuMaster(Serial(port="COM3", baudrate=38400, bytesize=8, parity='N', stopbits=1))
pid.set_timeout(0.2)
pid.set_verbose(True)

day_1, mon_1, year_1, min_1, hour_1, year_2, Sear_content, Sear_Type, Block_main1, Block_main = 0, 0, 0, 0, 0, b'', "", "", 0, b''
Pre_Heat_Touch, prev_Pre_Heat_Touch, Sv_flag, Sv_flag1, prev_read_flag, Preset_temp = 1, 0, 1, 1, 0, [0, 121, 232, 288, 316]
flag, flag1, a, j, Search, Save_DT, Signal, test_id, temp_l, temp_h = 0, 0, 3, 0, 0, 0, 0, 0, 0, 0
scroll_up_dw, timer1, timer2, flag_timer, flag_timer2, ready, ready2, res1, set_point, Block_temp_float, Block_temp_int, Block_int, Block_float, Transfer = 0, 0, 0, 1, 1, 0, 0, 0, 0, b'', b'', 0, 0.0, 0
res_float, prev_setpoint, timer_sec, timer_sec2, prev_scroll_dw, flag_scroll, Drop_point_int = 0.0, 0, 0, 0, 0, 0, 0
P1_FileName, P1_Drop_Temp, P2_FileName, P2_Drop_Temp, P3_FileName, P3_Drop_Temp, P4_FileName, P4_Drop_Temp, P5_FileName, P5_Drop_Temp, P6_FileName, P6_Drop_Temp = "", 0, "", 0, "", 0, "", 0, "", 0, "", 0
P1_Drop_Point, P1_Save, P2_Drop_Point, P2_Save, P3_Drop_Point, P3_Save, P4_Drop_Point, P4_Save, P5_Drop_Point, P5_Save, P6_Drop_Point, P6_Save = 0.0, 0, 0.0, 0, 0.0, 0, 0.0, 0, 0.0, 0, 0.0, 0
ready_flag, ready_flag2, initial_data, empty_search, current_len_data, old_len_data, cutoff_value, flag_scroll1, Prev_Temp_cal_Pre_Heat_Touch, Temp_cal_Pre_Heat_Touch = 0, 0, 1, 0, 0, 0, 400, 0, 0, 0
cut_off_flag, Cal_on_off_flag, Cal_on_off_flag_1, main_on_off_flag, Temp_cal_Heater_on_off, Temp_cal_enable, Temp_cal_Offset, Offset_value, Offset_value_int, pre_Offset_value, Offset_reset, Temp_cal_state = 1, 0, 1, 1, 0, 0, 0, 0.0, 0, 0, 0, 0
value1, value2, value3, value4, value5, value6, pre_heat_flag1, pre_heat_flag2, pre_set_point_flag1, stabilize_flag, stabilize_flag2, on_off_flag1 = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
temp_1, temp_2, temp_3, temp_4, pre_Offset_value_float, check_search_avail, Shutdown = 0.0, 0.0, 0.0, 0.0, 0.0, 0, 0
float_data, read_flag1, frame, data, data1, data2, data3, data4 = [0, 0, 0, 0], 0, b'', b'', b'', b'', b'', b''
printer_list1, cut_flag, sock_hr1, sock_min1, sock_sec1, sock_hr2, sock_min2, sock_sec2 = 0, 0, 0, 30, 0, 0, 30, 0
printers, Printer, printer1, print_flag, set_temp_clr, i, tail_str, tail_int = 0, 0, 0, 1, 0, b'\x01', b'\x00\xcc\x33\xc3\x3c', b'\xcc\x33\xc3\x3c'

mode = 0o666
if not os.path.isdir(r"D:\Dropping_Point"):
    make = os.path.join("D:/", "Dropping_Point")
    os.mkdir(make, mode)
if not os.path.isfile(r"D:\Dropping_Point\Imp_datas_dont_delete.CSV"):
    file1 = open(r"D:\Dropping_Point\Imp_datas_dont_delete.CSV", 'w')
    file1.write("0")
    file1.close()
if not os.path.isfile(r"D:\Dropping_Point\search.CSV"):
    file1 = open(r"D:\Dropping_Point\search.CSV", 'w')
    file1.write("Test_ID,Date,Time,File_Name, Drop_Temp, Block_Temp, Drop_Point \n")
    file1.close()
if not os.path.isfile(r"D:\Dropping_Point\DATA.CSV"):
    file1 = open(r"D:\Dropping_Point\DATA.CSV", "w")
    file1.write("")
    file1.close()

path = "D:\Dropping_Point\DATA.CSV"
file1 = open("D:\Dropping_Point\search1.CSV", 'w')
file1.write("Test_ID,Date,Time,File_Name, Drop_Temp, Block_Temp, Drop_Point \n")
file1.close()

file1 = open("D:\Dropping_Point\search2.CSV", 'w')
file1.write("Test_ID,Date,Time,File_Name, Drop_Temp, Block_Temp, Drop_Point \n")
file1.write("Test_ID,Date,Time,File_Name, Drop_Temp, Block_Temp, Drop_Point \n")
file1.close()

# Presetting works------------------------------------------------------------------------------------------------------
ser.write(b'\xaa\x42\x00\x00\x04\x00' + bytes("File Name", 'utf-8') + tail_str)

ser.write(b'\xaa\x3d\x00\x08\x00\x1A\x00\x00\xcc\x33\xc3\x3c')

ser.write(b'\xaa\x3d\x00\x08\x00\x58\x00\x1e\xcc\x33\xc3\x3c')
ser.write(b'\xaa\x3d\x00\x08\x00\x6c\x00\x1e\xcc\x33\xc3\x3c')
ser.write(b'\xaa\x3d\x00\x08\x00\x3c\x01\x90\xcc\x33\xc3\x3c')
ser.write(b'\xaa\x3d\x00\x08\x00\x38\x00\x01\xcc\x33\xc3\x3c')
ser.write(b'\xaa\x3d\x00\x08\x00\x68\x00\x01\xcc\x33\xc3\x3c')
ser.write(b'\xaa\x3d\x00\x08\x00\x8A\x00\x00\xcc\x33\xc3\x3c')
ser.write(b'\xaa\x3d\x00\x08\x00\x9A\x00\x00\xcc\x33\xc3\x3c')

file2 = open("D:\Dropping_Point\Imp_datas_dont_delete.CSV", "r")
pre_Offset_value = int(file2.read())
pre_Offset_value_float = pre_Offset_value / 10
file2.close()
print(pre_Offset_value_float)
ser.write(b'\xaa\x44\x00\x02\x00\x04' + struct.pack("!f", pre_Offset_value_float) + tail_int)

# Presetting works------------------------------------------------------------------------------------------------------

def read_row_add(f_filename, f_drop_temp, f_oven_temp, f_drop_point):
    ser.write(b'\xaa\x42\x00\x00\x1A\x00' + bytes(f_filename, 'utf-8') + tail_str)
    ser.write(b'\xaa\x3d\x00\x08\x00\x3A' + int(f_drop_temp).to_bytes(2, 'big') + tail_int)
    ser.write(b'\xaa\x3d\x00\x08\x00\x5c' + int(f_oven_temp).to_bytes(2, 'big') + tail_int)
    ser.write(b'\xaa\x3d\x00\x08\x00\x3E' + int(f_drop_point).to_bytes(2, 'big') + tail_int)


def len_Datafile():
    with open("D:\Dropping_Point\DATA.CSV", mode="r") as fn:
        fn1 = csv.reader(fn)
        len_data = len(list(fn1))
    return len_data


def printer(filename, value_1, value_2):
    o = open(filename, 'r')
    mydata = csv.reader(o)
    if value_1 == 1:
        if value_2 == 1:
            file1 = open("D:\Dropping_Point\search1.CSV", 'w')
            file1.write("Test_ID,Date,Time,File_Name, Drop_Temp, Block_Temp, Drop_Point \n")
            file1.close()
            with open("D:\Dropping_Point\search1.CSV", mode="a", newline="") as new_file:
                writer_obj = csv.writer(new_file, delimiter=",")
                for row in mydata:
                    writer_obj.writerow(row)
        else:
            file1 = open("D:\Dropping_Point\DATA_USB.CSV", 'w')
            file1.write("Test_ID,Date,Time,File_Name, Drop_Temp, Block_Temp, Drop_Point \n")
            file1.close()
            with open("D:\Dropping_Point\DATA_USB.CSV", mode="a", newline="") as new_file:
                writer_obj = csv.writer(new_file, delimiter=",")
                for row in mydata:
                    writer_obj.writerow(row)
    else:
        file1 = open("D:\Dropping_Point\search2.CSV", 'w')
        file1.write("Test_ID,Date,Time,File_Name, Drop_Temp, Block_Temp, Drop_Point \n")
        file1.write("Test_ID,Date,Time,File_Name, Drop_Temp, Block_Temp, Drop_Point \n")
        file1.close()
        with open("D:\Dropping_Point\search2.CSV", mode="a", newline="") as new_file:
            writer_obj = csv.writer(new_file, delimiter=",")
            for row in mydata:
                writer_obj.writerow(row)
    new_file.close()
    o.close()


def search_by(filename, col_number, value):
    var = str(value)
    o = open(filename, 'r')
    mydata = csv.reader(o)
    Signal1 = 0
    with open("D:\Dropping_Point\search.CSV", mode="a", newline="") as new_file:
        writer_obj = csv.writer(new_file, delimiter=",")
        for row in mydata:
            if row[col_number].lower() == var.lower():
                Signal1 = 1
                writer_obj.writerow(row)
    new_file.close()
    o.close()
    return Signal1


def table_Search(scroll_up_dw_value):
    index = 1
    addr = 1024
    signal, count = 0, 0

    val = index + scroll_up_dw_value

    with open("D:\Dropping_Point\search.CSV", mode="r") as fn:
        fn1 = csv.reader(fn)
        for index1, row in enumerate(fn1):
            if index1 == val:
                signal = 1
            if signal == 1:
                count = count + 1
                for y in range(0, 7):
                    if y != 2:
                        addr = addr + 128
                        ser.write(b'\xaa\x42\x00\x00' + addr.to_bytes(2, 'big') + bytes(row[y], 'utf-8') + tail_str)
            if count == 4:
                break
        if count != 4:
            for count in range(4):
                for z in range(1, 7):
                    addr = addr + 128
                    ser.write(b'\xaa\x42\x00\x00' + addr.to_bytes(2, 'big') + bytes(str('\0'), 'utf-8') + tail_str)


def table_Data(scroll_up_dw_value):
    addr = 1024
    signal, count = 0, 0
    val = scroll_up_dw_value

    with open("D:\Dropping_Point\DATA.CSV", mode="r") as fn:
        for index1, row in enumerate(reversed(list(csv.reader(fn)))):
            if index1 == val:
                signal = 1
            if signal == 1:
                count = count + 1
                for y in range(0, 7):
                    if y != 2:
                        addr = addr + 128
                        ser.write(b'\xaa\x42\x00\x00' + addr.to_bytes(2, 'big') + bytes(row[y], 'utf-8') + tail_str)
            if count == 4:
                break
        if count != 4:
            for count in range(4):
                for z in range(1, 7):
                    addr = addr + 128
                    ser.write(b'\xaa\x42\x00\x00' + addr.to_bytes(2, 'big') + bytes(str('\0'), 'utf-8') + tail_str)


def read_row(method, flag1):
    if method == 1:
        with open("D:\Dropping_Point\search.CSV", mode="r") as obj:
            obj1 = csv.reader(obj)
            flag2 = 0

            if flag1 != 0 or flag_scroll != 1:
                if flag1 == 0:
                    flag1 = 1

                flag1 = flag1 + scroll_up_dw
                for row in obj1:
                    if flag1 == flag2:
                        f_filename = row[3]
                        f_drop_temp = row[4]
                        f_oven_temp = row[5]
                        f_drop_point = row[6]
                        read_row_add(f_filename, f_drop_temp, f_oven_temp, f_drop_point)
                        break
                    flag2 = flag2 + 1
    else:
        with open("D:\Dropping_Point\DATA.CSV", mode="r") as obj:
            flag2 = 0

            if flag1 != 0 or flag_scroll1 != 1:

                flag1 = flag1 + scroll_up_dw - 1
                for row in reversed(list(csv.reader(obj))):
                    if flag1 == flag2:
                        f_filename = row[3]
                        f_drop_temp = row[4]
                        f_oven_temp = row[5]
                        f_drop_point = row[6]
                        read_row_add(f_filename, f_drop_temp, f_oven_temp, f_drop_point)
                        break
                    flag2 = flag2 + 1


def socking_time(Sock_hour, Sock_min, Sock_sec):
    return ((Sock_hour * 60) + Sock_min + Sock_sec)


def P_timer(p_pos):
    global value1, value2, value3, value4, value5, value6
    if p_pos == 1:
        value1 = 1
    if p_pos == 2:
        value2 = 1
    if p_pos == 3:
        value3 = 1
    if p_pos == 4:
        value4 = 1
    if p_pos == 5:
        value5 = 1
    if p_pos == 6:
        value6 = 1


# --------------------------------------------------main_Start----------------------------------------------------------
while True:
    main_add, touch_add, touch_add1, touch_data, touch_data_l, touch_data_h, ad, j = '', '', '', 0, 0, 0, '', 0
    while ser.in_waiting != 0:
        ad = ser.read()
        j = j + 1
        if j == 4:
            main_add = ad
        if main_add == b'\x08':
            if j == 6:
                touch_add = ad
            elif j == 7:
                touch_data_h = int.from_bytes(ad, "big")
            elif j == 8:
                touch_data_l = int.from_bytes(ad, "big")
                touch_data = touch_data_l | (touch_data_h << 8)
                break
        elif main_add == b'\x02':
            if j == 6:
                touch_add = ad
            elif j == 7:
                float_data[0] = int.from_bytes(ad, "big")
            elif j == 8:
                float_data[1] = int.from_bytes(ad, "big")
            elif j == 9:
                float_data[2] = int.from_bytes(ad, "big")
            elif j == 10:
                float_data[3] = int.from_bytes(ad, "big")
                break
        elif main_add == b'\x00':
            if j == 5:
                touch_add = ad
            elif j == 6:
                touch_add1 = ad
                while i != b'\x00':
                    i = ser.read()
                    if i != b'\x00':
                        frame += i
                i = b'\x01'
                break

    if main_add == b'\x08' and touch_add == b'\x54':
        read_flag1 = touch_data

    if main_add == b'\x08' and touch_add == b'\x56':
        sock_hr1 = touch_data

    elif main_add == b'\x08' and touch_add == b'\x58':
        sock_min1 = touch_data

    elif main_add == b'\x08' and touch_add == b'\x5a':
        sock_sec1 = touch_data

    if main_add == b'\x08' and touch_add == b'\x6a':
        sock_hr2 = touch_data

    elif main_add == b'\x08' and touch_add == b'\x6c':
        sock_min2 = touch_data

    elif main_add == b'\x08' and touch_add == b'\x6e':
        sock_sec2 = touch_data

    current = datetime.now()
    date = str(str(current.day) + '/' + str(current.month) + '/' + str(current.year - 2000))
    time1 = str(str(current.hour) + ':' + str(current.minute) + ':' + str(current.second))

    ser.write(b'\xaa\x3d\x00\x08\x00\x94\x00\x01\xcc\x33\xc3\x3c')  # loading sequence

    if (os.path.isfile(path) == False):
        file1 = open("D:\Dropping_Point\DATA.CSV", "w")
        file1.write("")
        file1.close()

    file2 = open("D:\Dropping_Point\Imp_datas_dont_delete.CSV", "r")
    pre_Offset_value = file2.read()
    pre_Offset_value = int(pre_Offset_value)
    file2.close()

    if value1 == 1:
        ser.write(b'\xaa\x3d\x00\x08\x00\x7e\x00\x00\xcc\x33\xc3\x3c')
        value1 = 0
    if value2 == 1:
        ser.write(b'\xaa\x3d\x00\x08\x00\x80\x00\x00\xcc\x33\xc3\x3c')
        value2 = 0
    if value3 == 1:
        ser.write(b'\xaa\x3d\x00\x08\x00\x82\x00\x00\xcc\x33\xc3\x3c')
        value3 = 0
    if value4 == 1:
        ser.write(b'\xaa\x3d\x00\x08\x00\x84\x00\x00\xcc\x33\xc3\x3c')
        value4 = 0
    if value5 == 1:
        ser.write(b'\xaa\x3d\x00\x08\x00\x86\x00\x00\xcc\x33\xc3\x3c')
        value5 = 0
    if value6 == 1:
        ser.write(b'\xaa\x3d\x00\x08\x00\x88\x00\x00\xcc\x33\xc3\x3c')
        value6 = 0
    # Power_off  PC-----------------------------------------------------------------------------------------------------
    if main_add == b'\x08' and touch_add == b'\x8c':
        pid.execute(slave=1, function_code=5, starting_address=0x0000, output_value=True)
        ser.write(b'\xaa\x70\x00\x04\xcc\x33\xc3\x3c')
        os.system("shutdown /s /t 1")

    # Power_off  PC-----------------------------------------------------------------------------------------------------
    # Main PID Controller-----------------------------------------------------------------------------------------------
    Block_main = pid.execute(slave=1, function_code=4, starting_address=0x03e8, quantity_of_x=1, data_format='',
                             returns_raw=True)
    Block_main1 = int.from_bytes(Block_main)
    Block_int = int(Block_main1 / 10)
    Block_float = Block_main1 / 10
    Block_float1 = (floor(Block_float * 10) % 10)

    Block_temp_int = Block_int.to_bytes(2, 'big')
    Block_temp_float = Block_float1.to_bytes(2, 'big')

    ser.write(b'\xaa\x3d\x00\x08\x00\x4E' + Block_temp_int + tail_int)
    ser.write(b'\xaa\x3d\x00\x08\x00\x50' + Block_temp_float + tail_int)
    ser.write(b'\xaa\x3d\x00\x08\x00\x7A' + Block_temp_int + tail_int)  # Temp_cal_block_temp_print
    ser.write(b'\xaa\x3d\x00\x08\x00\x7C' + Block_temp_float + tail_int)

    if main_add == b'\x08' and touch_add == b'\x3c':
        cutoff_value = touch_data

    if cutoff_value < Block_int and cut_off_flag == 1:
        pid.execute(slave=1, function_code=5, starting_address=0x0000, output_value=True)
        ser.write(b'\xaa\x7A\x05\x05\x05\x05\x32\xcc\x33\xc3\x3c')  # buzzer call
        ser.write(b'\xaa\x3d\x00\x08\x00\x98\x00\x01\xcc\x33\xc3\x3c')
        sleep(2)
        ser.write(b'\xaa\x3d\x00\x08\x00\x98\x00\x00\xcc\x33\xc3\x3c')
        flag, Cal_on_off_flag, cut_off_flag = 0, 0, 0
    elif cutoff_value > Block_int:
        cut_off_flag = 1

    # Main PID Controller-----------------------------------END---------------------------------------------------------
    # Set_main_Block_Temp-----------------------------------------------------------------------------------------------
    # Preset_Heating_ON_OFF---------------------------------------------------------------------------------------------

    if main_add == b'\x08' and touch_add == b'\x36' and Cal_on_off_flag != 1:
        if cut_off_flag == 0:
            pid.execute(slave=1, function_code=5, starting_address=0x0000, output_value=True)
            ser.write(b'\xaa\x7A\x05\x05\x05\x05\x32\xcc\x33\xc3\x3c')  # buzzer call
            ser.write(b'\xaa\x3d\x00\x08\x00\x98\x00\x01\xcc\x33\xc3\x3c')
            sleep(2)
            ser.write(b'\xaa\x3d\x00\x08\x00\x98\x00\x00\xcc\x33\xc3\x3c')
        else:
            if flag:
                flag = 0
            else:
                flag = 1

    if (flag == 1 and cut_off_flag != 0) and Sv_flag == 1 and Cal_on_off_flag != 1:
        ser.write(b'\xaa\x3d\x00\x08\x00\x34\x00\x01\xcc\x33\xc3\x3c')
        pid.execute(slave=1, function_code=5, starting_address=0x0000, output_value=False)
        pid.execute(slave=1, function_code=5, starting_address=0x0001, output_value=True)
        Sv_flag = 0
    elif Cal_on_off_flag != 1 and flag != 1:
        Sv_flag = 1
        on_off_flag1 = 1
        ser.write(b'\xaa\x3d\x00\x08\x00\x34\x00\x00\xcc\x33\xc3\x3c')
        pid.execute(slave=1, function_code=5, starting_address=0x0001, output_value=False)
        pid.execute(slave=1, function_code=5, starting_address=0x0000, output_value=True)
    elif Cal_on_off_flag == 1:
        flag = 0

    if main_add == b'\x08' and touch_add == b'\x32'and Cal_on_off_flag != 1:
        set_point = touch_data
        ser.write(b'\xaa\x3d\x00\x08\x00\x1A' + set_point.to_bytes(2, 'big') + tail_int)
        set_point = set_point * 10

    if prev_setpoint != set_point and Cal_on_off_flag != 1:
        set_temp_clr = 1
        prev_setpoint = set_point
        pre_set_point_flag1, prev_Pre_Heat_Touch, Pre_Heat_Touch, Prev_Temp_cal_Pre_Heat_Touch, Temp_cal_Pre_Heat_Touch = 1, 0, 0, 0, 0

        if flag:
            Sv_flag = 1
        pid.execute(slave=1, function_code=5, starting_address=0x0001, output_value=False)
        pid.execute(slave=1, function_code=6, starting_address=0x0000, output_value=set_point)
        ser.write(b'\xaa\x3d\x00\x08\x00\x2a\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x2c\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x2e\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x30\x00\x00\xcc\x33\xc3\x3c')

    if main_add == b'\x08' and touch_add == b'\x38' and Cal_on_off_flag != 1:
        Pre_Heat_Touch = touch_data

    if prev_Pre_Heat_Touch != Pre_Heat_Touch and Cal_on_off_flag != 1:
        set_temp_clr, Prev_Temp_cal_Pre_Heat_Touch, Temp_cal_Pre_Heat_Touch, pre_heat_flag1, prev_Pre_Heat_Touch = 1, 0, 0, 1, Pre_Heat_Touch
        if flag:
            Sv_flag = 1
        if Pre_Heat_Touch == 1:
            ser.write(b'\xaa\x3d\x00\x08\x00\x2a\x00\x01\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x2c\x00\x00\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x2e\x00\x00\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x30\x00\x00\xcc\x33\xc3\x3c')
            pid.execute(slave=1, function_code=5, starting_address=0x0001, output_value=False)
            pid.execute(slave=1, function_code=6, starting_address=0x0000, output_value=1210)

        elif Pre_Heat_Touch == 2:
            ser.write(b'\xaa\x3d\x00\x08\x00\x2a\x00\x00\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x2c\x00\x01\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x2e\x00\x00\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x30\x00\x00\xcc\x33\xc3\x3c')
            pid.execute(slave=1, function_code=5, starting_address=0x0001, output_value=False)
            pid.execute(slave=1, function_code=6, starting_address=0x0000, output_value=2320)

        elif Pre_Heat_Touch == 3:
            ser.write(b'\xaa\x3d\x00\x08\x00\x2a\x00\x00\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x2c\x00\x00\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x2e\x00\x01\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x30\x00\x00\xcc\x33\xc3\x3c')
            pid.execute(slave=1, function_code=5, starting_address=0x0001, output_value=False)
            pid.execute(slave=1, function_code=6, starting_address=0x0000, output_value=2880)

        elif Pre_Heat_Touch == 4:
            ser.write(b'\xaa\x3d\x00\x08\x00\x2a\x00\x00\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x2c\x00\x00\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x2e\x00\x00\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x30\x00\x01\xcc\x33\xc3\x3c')
            pid.execute(slave=1, function_code=5, starting_address=0x0001, output_value=False)
            pid.execute(slave=1, function_code=6, starting_address=0x0000, output_value=3160)

    if set_temp_clr == 1 and Cal_on_off_flag != 1:
        ser.write(b'\xaa\x3d\x00\x08\x00\x5E\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x60\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x62\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x64\x00\x00\xcc\x33\xc3\x3c')
    elif set_temp_clr == 2 and flag != 1:
        ser.write(b'\xaa\x3d\x00\x08\x00\x1A\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x32\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x2a\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x2c\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x2e\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x30\x00\x00\xcc\x33\xc3\x3c')
    # Set_main_Block_Temp--------------------------------------------END------------------------------------------------
    # Preset_Heating_ON_OFF----------------------------------------END--------------------------------------------------
    # Search Transfer---------------------------------------------------------------------------------------------------
    if main_add == b'\x00' and touch_add == b'\x03' and touch_add1 == b'\x80':
        Sear_content = frame.decode()
        frame = b''
    if main_add == b'\x00' and touch_add == b'\x04' and touch_add1 == b'\x00':
        Sear_Type = frame.decode()
        frame = b''
    if main_add == b'\x08' and touch_add == b'\x40':
        scroll_up_dw = touch_data
    if main_add == b'\x08' and touch_add == b'\x24':
        prev_read_flag = 0
        initial_data = 0
        scroll_up_dw = 0
        ser.write(b'\xaa\x3d\x00\x08\x00\x40\x00\x00\xcc\x33\xc3\x3c')  # Scroll to be Zero
        read_flag1 = 0

        if initial_data == 1:
            flag_scroll1 = 0

        match Sear_Type:
            case 'File Name':
                a = 3
            case 'Drop Temp':
                a = 4
            case 'Block Temp':
                a = 5
            case 'Drop Point':
                a = 6
        file1 = open("D:\Dropping_Point\search.CSV", 'w')
        file1.write("Test_ID,Date,Time,File_Name, Drop_Temp, Block_Temp, Drop_Point \n")
        file1.close()

        Signal = search_by(path, a, Sear_content)
        flag_scroll = 0
        if Signal == 0:
            empty_search = 0
        if Sear_content == '.':
            empty_search, flag_scroll1 = 1, 0

    if (empty_search == 1 or initial_data == 1) and len_Datafile() != 0:
        flag1 = 1
        check_search_avail = 1
        if flag_scroll1 == 0:
            table_Data(scroll_up_dw)

        if prev_scroll_dw != scroll_up_dw:
            prev_scroll_dw = scroll_up_dw
            read_flag1 = 0
            flag_scroll1 = 1
            table_Data(scroll_up_dw)
        read_row(2, read_flag1)
    else:
        check_search_avail = 0
        flag1 = 0

    if Signal == 1:
        empty_search = 0
        if flag_scroll == 0:
            table_Search(scroll_up_dw)
        if prev_scroll_dw != scroll_up_dw:
            prev_scroll_dw = scroll_up_dw
            read_flag1 = 0
            flag_scroll = 1
            table_Search(scroll_up_dw)
        read_row(1, read_flag1)

    elif Signal != 1 and flag1 == 0:
        ser.write(b'\xaa\x42\x00\x00\x1A\x00' + bytes(str('\0'), 'utf-8') + tail_str)
        ser.write(b'\xaa\x3d\x00\x08\x00\x3A\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x5c\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x3E\x00\x00\xcc\x33\xc3\x3c')
        addr_n = 1024
        for count in range(4):
            for z in range(1, 7):
                addr_n = addr_n + 128
                ser.write(b'\xaa\x42\x00\x00' + addr_n.to_bytes(2, 'big') + bytes(str('\0'), 'utf-8') + tail_str)

    # Search Transfer---------------------------------------------------------------------------------------------------
    # USB Transfer------------------------------------------------------------------------------------------------------

    if main_add == b'\x08' and touch_add == b'\x28':
        Filename = "DATA" + '_' + str(current.day) + '-' + str(current.month) + '-' + str(current.year) + '-' + str(
            current.hour) + '-' + str(current.minute) + '-' + str(current.second)
        if os.path.isdir('E:'):
            if check_search_avail == 0:
                pen_drive.copyfile(src=r"D:\Dropping_Point\search.CSV", dst="E:/" + Filename + ".CSV")
            else:
                printer(r"D:\Dropping_Point\DATA.csv", 1, 2)
                pen_drive.copyfile(src=r"D:\Dropping_Point\DATA_USB.CSV", dst="E:/" + Filename + ".CSV")
            ser.write(b'\xaa\x3d\x00\x08\x00\x92\x00\x01\xcc\x33\xc3\x3c')
            sleep(2)
            ser.write(b'\xaa\x3d\x00\x08\x00\x92\x00\x00\xcc\x33\xc3\x3c')
        else:
            print("USB_not_avail")
            ser.write(b'\xaa\x3d\x00\x08\x00\x90\x00\x01\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x7A\x05\x05\x00\x05\x00\xcc\x33\xc3\x3c')  # buzzer call
            sleep(2)
            ser.write(b'\xaa\x3d\x00\x08\x00\x90\x00\x00\xcc\x33\xc3\x3c')

    # USB Transfer------------------------------------------------------------------------------------------------------
    # Printer-----------------------------------------------------------------------------------------------------------
    current_printer = win32print.GetDefaultPrinter()
    handle = win32print.OpenPrinter(current_printer)
    attributes = win32print.GetPrinter(handle)[13]
    if attributes == 2624 and print_flag == 1:
        print("printer connected")
        ser.write(b'\xaa\x3d\x00\x08\x00\x96\x00\x01\xcc\x33\xc3\x3c')
        sleep(1)
        ser.write(b'\xaa\x3d\x00\x08\x00\x96\x00\x00\xcc\x33\xc3\x3c')
        print_flag = 0
    elif attributes == 3648 and print_flag == 0:
        print_flag = 1
        print("printer Not connected")

    if main_add == b'\x08' and touch_add == b'\x26':
        if attributes == 2624:
            # Change Printer Page Size
            handle = win32print.OpenPrinter(current_printer, {'DesiredAccess': win32print.PRINTER_ALL_ACCESS})
            proper = win32print.GetPrinter(handle, 2)
            dev_mode = proper['pDevMode']
            dev_mode.PaperSize = 9  # 9 -- A4, 11 -- A5, ...
            win32print.SetPrinter(handle, 2, proper, 0)
            # Change Printer Page Size---------End
            if check_search_avail == 0:
                print("search")
                printer(r"D:\Dropping_Point\search.csv", 1, 1)
                convert(r"D:\Dropping_Point\search1.csv", r"D:\Dropping_Point\print_data.pdf")
            else:
                print("default")
                printer(r"D:\Dropping_Point\DATA.csv", 2, 1)
                convert(r"D:\Dropping_Point\search2.csv", r"D:\Dropping_Point\print_data.pdf")

            GHOSTSCRIPT_PATH = "D:\\GHOSTSCRIPT\\bin\\gswin32.exe"
            GSPRINT_PATH = "D:\\GSPRINT\\gsprint.exe"
            ShellExecute(0, 'open', GSPRINT_PATH,
                         '-ghostscript "' + GHOSTSCRIPT_PATH + '" -printer "' + current_printer +
                         '" "D:\Dropping_Point\Print_data.pdf"',
                         '.', 0)
        else:
            print("printer Not connected")
            ser.write(b'\xaa\x3d\x00\x08\x00\x8e\x00\x01\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x7A\x05\x05\x00\x05\x00\xcc\x33\xc3\x3c')  # buzzer call
            sleep(2)
            ser.write(b'\xaa\x3d\x00\x08\x00\x8e\x00\x00\xcc\x33\xc3\x3c')

    # Printer-----------------------------------------------------------------------------------------------------------
    # Transfer----------------------------------------------END---------------------------------------------------------
    # Date_Time_set-----------------------------------------------------------------------------------------------------
    if main_add == b'\x08' and touch_add == b'\x44':
        day_1 = touch_data
    elif main_add == b'\x08' and touch_add == b'\x46':
        mon_1 = touch_data
    elif main_add == b'\x08' and touch_add == b'\x48':
        year_1 = touch_data
    elif main_add == b'\x08' and touch_add == b'\x4a':
        hour_1 = touch_data
    elif main_add == b'\x08' and touch_add == b'\x4c':
        min_1 = touch_data

    if main_add == b'\x08' and touch_add == b'\x52' and day_1 != 0 and mon_1 != 0 and year_1 != 0:
        if (year_1 >= 2000) and (year_1 <= 2099):
            year_1 = year_1 - 2000
            year_2 = year_1.to_bytes(1, 'big')
        elif (year_1 >= 0) and (year_1 <= 99):
            year_2 = year_1.to_bytes(1, 'big')
        else:
            ser.write(b'\xaa\x3d\x00\x08\x00\x48\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x9c' + year_2 + mon_1.to_bytes(1, 'big') + day_1.to_bytes(1, 'big') + hour_1.to_bytes(1,
                                                                                                               'big') + min_1.to_bytes(
            1, 'big') + b'\x00\xcc\x33\xc3\x3c')

    # Date_Time_set----------------------------------------------END----------------------------------------------------

    # Set_Block_Temp_Status---------------------------------------------------------------------------------------------
    if flag == 1 and cut_off_flag == 1 and Cal_on_off_flag != 1:
        Sv_value = pid.execute(slave=1, function_code=3, starting_address=0x0000, quantity_of_x=1, data_format='',
                               returns_raw=True)
        Sv_value = int.from_bytes(Sv_value, byteorder='big', signed=False)
        Sv_value = int(Sv_value / 10)
        if Sv_value > Block_int and ready != 1 and stabilize_flag != 1:
            ser.write(b'\xaa\x42\x00\x00\x1a\x80' + bytes("Heating", 'utf-8') + tail_str)
            flag_timer, ready = 1, 0
        elif (Sv_value <= Block_int and ready != 1) or stabilize_flag == 1:
            stabilize_flag = 1
            ser.write(b'\xaa\x42\x00\x00\x1a\x80' + bytes("Stabilizing", 'utf-8') + tail_str)
            if flag_timer == 1:
                timer1 = ((current.hour * 60) + current.minute + socking_time(sock_hr1, sock_min1, sock_sec1))
                timer_sec = current.second
                flag_timer, ready_flag = 0, 1

            if (current.hour * 60) + current.minute >= timer1:
                if timer_sec <= current.second:
                    ready, stabilize_flag = 1, 0
        if pre_set_point_flag1 == 1 or pre_heat_flag1 == 1:
            ready, pre_set_point_flag1, pre_heat_flag1, stabilize_flag, flag_timer, on_off_flag1 = 0, 0, 0, 0, 1, 0

        if ready == 1:
            ser.write(b'\xaa\x42\x00\x00\x1a\x80' + bytes("Ready", 'utf-8') + tail_str)
            if ready_flag == 1:
                ready_flag = 0
                ser.write(b'\xaa\x7A\x05\x05\x05\x05\x32\xcc\x33\xc3\x3c')  # buzzer call
                ser.write(b'\xaa\x3d\x00\x08\x00\x70\x00\x01\xcc\x33\xc3\x3c')
                sleep(2)
                ser.write(b'\xaa\x3d\x00\x08\x00\x70\x00\x00\xcc\x33\xc3\x3c')
    else:
        ser.write(b'\xaa\x42\x00\x00\x1a\x80' + bytes("Heater OFF", 'utf-8') + tail_str)
        ready, pre_set_point_flag1, pre_heat_flag1, stabilize_flag, flag_timer, on_off_flag1 = 0, 0, 0, 0, 1, 0

    # Set_Block_Temp_Status----------------------------------------------END--------------------------------------------
    # Temp_calibration----------------------------------------Start-----------------------------------------------------
    # Temp_cal_Preset_Heating_ON_OFF------------------------------------------------------------------------------------
    if main_add == b'\x08' and touch_add == b'\x72' and flag != 1:
        if cut_off_flag == 0:
            pid.execute(slave=1, function_code=5, starting_address=0x0000, output_value=True)
            ser.write(b'\xaa\x7A\x05\x05\x05\x05\x32\xcc\x33\xc3\x3c')  # buzzer call
            ser.write(b'\xaa\x3d\x00\x08\x00\x98\x00\x01\xcc\x33\xc3\x3c')
            sleep(2)
            ser.write(b'\xaa\x3d\x00\x08\x00\x98\x00\x00\xcc\x33\xc3\x3c')
        else:
            if Cal_on_off_flag:
                Cal_on_off_flag = 0
            else:
                Cal_on_off_flag = 1

    if (Cal_on_off_flag == 1 and cut_off_flag != 0) and Sv_flag1 == 1 and flag != 1:
        ser.write(b'\xaa\x3d\x00\x08\x00\x74\x00\x01\xcc\x33\xc3\x3c')
        pid.execute(slave=1, function_code=5, starting_address=0x0000, output_value=False)
        pid.execute(slave=1, function_code=5, starting_address=0x0001, output_value=True)
        Sv_flag1 = 0
    elif flag != 1 and Cal_on_off_flag != 1:
        Sv_flag1 = 1
        ser.write(b'\xaa\x3d\x00\x08\x00\x74\x00\x00\xcc\x33\xc3\x3c')
        pid.execute(slave=1, function_code=5, starting_address=0x0001, output_value=False)
        pid.execute(slave=1, function_code=5, starting_address=0x0000, output_value=True)
    elif flag == 1:
        Cal_on_off_flag = 0

    if main_add == b'\x08' and touch_add == b'\x68' and flag != 1:
        Temp_cal_Pre_Heat_Touch = touch_data

    if Temp_cal_Pre_Heat_Touch != Prev_Temp_cal_Pre_Heat_Touch and flag != 1:
        pre_heat_flag2, Sv_flag1, set_temp_clr, Pre_Heat_Touch, prev_Pre_Heat_Touch, set_point, prev_setpoint = 1, 1, 2, 0, 0, 0, 0
        Prev_Temp_cal_Pre_Heat_Touch = Temp_cal_Pre_Heat_Touch
        if Temp_cal_Pre_Heat_Touch == 1:
            ser.write(b'\xaa\x3d\x00\x08\x00\x5E\x00\x01\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x60\x00\x00\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x62\x00\x00\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x64\x00\x00\xcc\x33\xc3\x3c')
            pid.execute(slave=1, function_code=5, starting_address=0x0001, output_value=False)
            pid.execute(slave=1, function_code=6, starting_address=0x0000, output_value=1210)

        elif Temp_cal_Pre_Heat_Touch == 2:
            ser.write(b'\xaa\x3d\x00\x08\x00\x5E\x00\x00\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x60\x00\x01\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x62\x00\x00\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x64\x00\x00\xcc\x33\xc3\x3c')
            pid.execute(slave=1, function_code=5, starting_address=0x0001, output_value=False)
            pid.execute(slave=1, function_code=6, starting_address=0x0000, output_value=2320)

        elif Temp_cal_Pre_Heat_Touch == 3:
            ser.write(b'\xaa\x3d\x00\x08\x00\x5E\x00\x00\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x60\x00\x00\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x62\x00\x01\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x64\x00\x00\xcc\x33\xc3\x3c')
            pid.execute(slave=1, function_code=5, starting_address=0x0001, output_value=False)
            pid.execute(slave=1, function_code=6, starting_address=0x0000, output_value=2880)

        elif Temp_cal_Pre_Heat_Touch == 4:
            ser.write(b'\xaa\x3d\x00\x08\x00\x5E\x00\x00\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x60\x00\x00\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x62\x00\x00\xcc\x33\xc3\x3c')
            ser.write(b'\xaa\x3d\x00\x08\x00\x64\x00\x01\xcc\x33\xc3\x3c')
            pid.execute(slave=1, function_code=5, starting_address=0x0001, output_value=False)
            pid.execute(slave=1, function_code=6, starting_address=0x0000, output_value=3160)
    # Temp_cal_Set_Block_Temp--------------------------------------------END--------------------------------------------

    # Temp_cal_Preset_Heating_ON_OFF----------------------------------------END-----------------------------------------
    # Temp_cal_Offset_set----------------------------------------Start--------------------------------------------------
    if main_add == b'\x08' and touch_add == b'\x78':
        pid.execute(slave=1, function_code=5, starting_address=0x0001, output_value=False)
        pid.execute(slave=1, function_code=6, starting_address=0x009e, output_value=0)

        file1 = open("D:\Dropping_Point\Imp_datas_dont_delete.CSV", "w")
        Offset_value = file1.write(str(0))
        file1.close()
        pre_Offset_value_float = 0.0
        Temp_cal_Offset = 1

    if main_add == b'\x02' and touch_add == b'\x04':
        b = bytes(float_data)
        Offset_value = struct.unpack('>f', b)
        Offset_value = round(Offset_value[0], 2)
        Offset_value_int = int(Offset_value * 10)

    if main_add == b'\x08' and touch_add == b'\x66':
        file1 = open("D:\Dropping_Point\Imp_datas_dont_delete.CSV", "w")
        Offset_value = file1.write(str(Offset_value_int + pre_Offset_value))
        file1.close()

        Offset_value_int = Offset_value_int + pre_Offset_value
        pid.execute(slave=1, function_code=5, starting_address=0x0001, output_value=False)
        pid.execute(slave=1, function_code=6, starting_address=0x009e, output_value=Offset_value_int)
    elif Temp_cal_Offset == 1:
        Temp_cal_Offset = 0
        ser.write(b'\xaa\x44\x00\x02\x00\x04' + struct.pack("!f", pre_Offset_value_float) + tail_int)

    if Cal_on_off_flag == 1 and cut_off_flag == 1 and flag != 1:

        Sv_value = pid.execute(slave=1, function_code=3, starting_address=0x0000, quantity_of_x=1, data_format='',
                               returns_raw=True)
        Sv_value = int.from_bytes(Sv_value, byteorder='big', signed=False)
        Sv_value = int(Sv_value / 10)

        if Sv_value > Block_int and ready2 != 1 and stabilize_flag2 != 1:
            ser.write(b'\xaa\x42\x00\x00\x1B\x00' + bytes("Heating", 'utf-8') + tail_str)
            flag_timer2, ready2 = 1, 0
        elif (Sv_value <= Block_int and ready2 != 1) or stabilize_flag2 == 1:
            stabilize_flag = 1
            ser.write(b'\xaa\x42\x00\x00\x1B\x00' + bytes("Stabilizing", 'utf-8') + tail_str)
            if flag_timer2 == 1:
                timer2 = ((current.hour * 60) + current.minute + socking_time(sock_hr2, sock_min2, sock_sec2))
                timer_sec2 = current.second
                flag_timer2, ready_flag2 = 0, 1

            if (current.hour * 60) + current.minute >= timer2:
                if timer_sec2 <= current.second:
                    ready2 = 1
        if pre_heat_flag2 == 1:
            ready2, pre_heat_flag2, stabilize_flag2, flag_timer2 = 0, 0, 0, 1

        if ready2 == 1:
            ser.write(b'\xaa\x42\x00\x00\x1B\x00' + bytes("Ready", 'utf-8') + tail_str)
            if ready_flag2 == 1:
                ready_flag2 = 0
                ser.write(b'\xaa\x7A\x05\x05\x05\x05\x32\xcc\x33\xc3\x3c')  # buzzer call
                ser.write(b'\xaa\x3d\x00\x08\x00\x70\x00\x01\xcc\x33\xc3\x3c')
                sleep(2)
                ser.write(b'\xaa\x3d\x00\x08\x00\x70\x00\x00\xcc\x33\xc3\x3c')
    else:
        ser.write(b'\xaa\x42\x00\x00\x1B\x00' + bytes("Heater OFF", 'utf-8') + tail_str)
        pre_heat_flag2, stabilize_flag2, flag_timer2, ready2 = 0, 0, 1, 0

    # Temp_cal_Offset_set----------------------------------------End----------------------------------------------------
    # Temp_calibration--------------------------------------------------------------------------------------------------

    # Saving Datas -----------------------------------------------------------------------------------------------------
    if ready == 1:
        # P1 -----------------------------------------------------------------------------------------------------------
        if main_add == b'\x00' and touch_add == b'\x00' and touch_add1 == b'\x80':
            P1_FileName = frame.decode()
            frame = b''
        elif main_add == b'\x00' and touch_add == b'\x01' and touch_add1 == b'\x00':
            P2_FileName = frame.decode()
            frame = b''
        elif main_add == b'\x00' and touch_add == b'\x01' and touch_add1 == b'\x80':
            P3_FileName = frame.decode()
            frame = b''
        elif main_add == b'\x00' and touch_add == b'\x02' and touch_add1 == b'\x00':
            P4_FileName = frame.decode()
            frame = b''
        elif main_add == b'\x00' and touch_add == b'\x02' and touch_add1 == b'\x80':
            P5_FileName = frame.decode()
            frame = b''
        elif main_add == b'\x00' and touch_add == b'\x03' and touch_add1 == b'\x00':
            P6_FileName = frame.decode()
            frame = b''

        if main_add == b'\x08' and touch_add == b'\x00':
            P1_Drop_Temp = touch_data
        elif main_add == b'\x08' and touch_add == b'\x04':
            P2_Drop_Temp = touch_data
        elif main_add == b'\x08' and touch_add == b'\x08':
            P3_Drop_Temp = touch_data
        elif main_add == b'\x08' and touch_add == b'\x0C':
            P4_Drop_Temp = touch_data
        elif main_add == b'\x08' and touch_add == b'\x10':
            P5_Drop_Temp = touch_data
        elif main_add == b'\x08' and touch_add == b'\x14':
            P6_Drop_Temp = touch_data

        if P1_Drop_Temp != 0:
            P1_Drop_Point = P1_Drop_Temp + ((Preset_temp[Pre_Heat_Touch] - P1_Drop_Temp) / 3)
            ser.write(b'\xaa\x3d\x00\x08\x00\x02' + int(P1_Drop_Point).to_bytes(2, 'big') + tail_int)
        if P2_Drop_Temp != 0:
            P2_Drop_Point = P2_Drop_Temp + ((Preset_temp[Pre_Heat_Touch] - P2_Drop_Temp) / 3)
            ser.write(b'\xaa\x3d\x00\x08\x00\x06' + int(P2_Drop_Point).to_bytes(2, 'big') + tail_int)
        if P3_Drop_Temp != 0:
            P3_Drop_Point = P3_Drop_Temp + ((Preset_temp[Pre_Heat_Touch] - P3_Drop_Temp) / 3)
            ser.write(b'\xaa\x3d\x00\x08\x00\x0A' + int(P3_Drop_Point).to_bytes(2, 'big') + tail_int)
        if P4_Drop_Temp != 0:
            P4_Drop_Point = P4_Drop_Temp + ((Preset_temp[Pre_Heat_Touch] - P4_Drop_Temp) / 3)
            ser.write(b'\xaa\x3d\x00\x08\x00\x0E' + int(P4_Drop_Point).to_bytes(2, 'big') + tail_int)
        if P5_Drop_Temp != 0:
            P5_Drop_Point = P5_Drop_Temp + ((Preset_temp[Pre_Heat_Touch] - P5_Drop_Temp) / 3)
            ser.write(b'\xaa\x3d\x00\x08\x00\x12' + int(P5_Drop_Point).to_bytes(2, 'big') + tail_int)
        if P6_Drop_Temp != 0:
            P6_Drop_Point = P6_Drop_Temp + ((Preset_temp[Pre_Heat_Touch] - P6_Drop_Temp) / 3)
            ser.write(b'\xaa\x3d\x00\x08\x00\x16' + int(P6_Drop_Point).to_bytes(2, 'big') + tail_int)

        if main_add == b'\x08' and touch_add == b'\x18':
            P1_Save = touch_data

        if P1_Save == 1:
            P1_Save = 0
            ser.write(b'\xaa\x3d\x00\x08\x00\x7E\x00\x01\xcc\x33\xc3\x3c')
            timer_p1 = Timer(30.0, P_timer, {1})
            timer_p1.start()

            length = csv.reader(open("D:\Dropping_Point\DATA.CSV"))
            test_id = len(list(length))
            file1 = open("D:\Dropping_Point\DATA.CSV", "a")
            if (file1):
                file1.seek(0, 2)
                file1.write(str(test_id))
                file1.write(",")
                file1.write(date)
                file1.write(",")
                file1.write(time1)
                file1.write(",")
                file1.write(P1_FileName)
                file1.write(",")
                file1.write(str(P1_Drop_Temp))
                file1.write(",")
                file1.write(str(round(Block_float)))
                file1.write(",")
                file1.write(str(round(P1_Drop_Point)))
                file1.write("\n")
                file1.close()

        elif P1_Save == 2:
            P1_Save = 0
            ser.write(b'\xaa\x3d\x00\x08\x00\x80\x00\x01\xcc\x33\xc3\x3c')
            timer_p2 = Timer(30.0, P_timer, {2})
            timer_p2.start()

            length = csv.reader(open("D:\Dropping_Point\DATA.CSV"))
            test_id = len(list(length))
            file1 = open("D:\Dropping_Point\DATA.CSV", "a")
            if (file1):
                file1.seek(0, 2)
                file1.write(str(test_id))
                file1.write(",")
                file1.write(date)
                file1.write(",")
                file1.write(time1)
                file1.write(",")
                file1.write(P2_FileName)
                file1.write(",")
                file1.write(str(P2_Drop_Temp))
                file1.write(",")
                file1.write(str(round(Block_float)))
                file1.write(",")
                file1.write(str(round(P2_Drop_Point)))
                file1.write("\n")
                file1.close()

        elif P1_Save == 3:
            P1_Save = 0
            ser.write(b'\xaa\x3d\x00\x08\x00\x82\x00\x01\xcc\x33\xc3\x3c')
            timer_p3 = Timer(30.0, P_timer, {3})
            timer_p3.start()

            length = csv.reader(open("D:\Dropping_Point\DATA.CSV"))
            test_id = len(list(length))
            file1 = open("D:\Dropping_Point\DATA.CSV", "a")
            if (file1):
                file1.seek(0, 2)
                file1.write(str(test_id))
                file1.write(",")
                file1.write(date)
                file1.write(",")
                file1.write(time1)
                file1.write(",")
                file1.write(P3_FileName)
                file1.write(",")
                file1.write(str(P3_Drop_Temp))
                file1.write(",")
                file1.write(str(round(Block_float)))
                file1.write(",")
                file1.write(str(round(P3_Drop_Point)))
                file1.write("\n")
                file1.close()

        elif P1_Save == 4:
            P1_Save = 0
            ser.write(b'\xaa\x3d\x00\x08\x00\x84\x00\x01\xcc\x33\xc3\x3c')
            timer_p4 = Timer(30.0, P_timer, {4})
            timer_p4.start()

            length = csv.reader(open("D:\Dropping_Point\DATA.CSV"))
            test_id = len(list(length))
            file1 = open("D:\Dropping_Point\DATA.CSV", "a")
            if (file1):
                file1.seek(0, 2)
                file1.write(str(test_id))
                file1.write(",")
                file1.write(date)
                file1.write(",")
                file1.write(time1)
                file1.write(",")
                file1.write(P4_FileName)
                file1.write(",")
                file1.write(str(P4_Drop_Temp))
                file1.write(",")
                file1.write(str(round(Block_float)))
                file1.write(",")
                file1.write(str(round(P4_Drop_Point)))
                file1.write("\n")
                file1.close()

        elif P1_Save == 5:
            P1_Save = 0
            ser.write(b'\xaa\x3d\x00\x08\x00\x86\x00\x01\xcc\x33\xc3\x3c')
            timer_p5 = Timer(30.0, P_timer, {5})
            timer_p5.start()

            length = csv.reader(open("D:\Dropping_Point\DATA.CSV"))
            test_id = len(list(length))
            file1 = open("D:\Dropping_Point\DATA.CSV", "a")
            if (file1):
                file1.seek(0, 2)
                file1.write(str(test_id))
                file1.write(",")
                file1.write(date)
                file1.write(",")
                file1.write(time1)
                file1.write(",")
                file1.write(P5_FileName)
                file1.write(",")
                file1.write(str(P5_Drop_Temp))
                file1.write(",")
                file1.write(str(round(Block_float)))
                file1.write(",")
                file1.write(str(round(P5_Drop_Point)))
                file1.write("\n")
                file1.close()

        elif P1_Save == 6:
            P1_Save = 0
            ser.write(b'\xaa\x3d\x00\x08\x00\x88\x00\x01\xcc\x33\xc3\x3c')
            timer_p6 = Timer(30.0, P_timer, {6})
            timer_p6.start()

            length = csv.reader(open("D:\Dropping_Point\DATA.CSV"))
            test_id = len(list(length))
            file1 = open("D:\Dropping_Point\DATA.CSV", "a")
            if (file1):
                file1.seek(0, 2)
                file1.write(str(test_id))
                file1.write(",")
                file1.write(date)
                file1.write(",")
                file1.write(time1)
                file1.write(",")
                file1.write(P6_FileName)
                file1.write(",")
                file1.write(str(P6_Drop_Temp))
                file1.write(",")
                file1.write(str(round(Block_float)))
                file1.write(",")
                file1.write(str(round(P6_Drop_Point)))
                file1.write("\n")
                file1.close()
    else:
        ser.write(b'\xaa\x42\x00\x00\x00\x80' + bytes(str('\0'), 'utf-8') + tail_str)
        ser.write(b'\xaa\x42\x00\x00\x01\x00' + bytes(str('\0'), 'utf-8') + tail_str)
        ser.write(b'\xaa\x42\x00\x00\x01\x80' + bytes(str('\0'), 'utf-8') + tail_str)
        ser.write(b'\xaa\x42\x00\x00\x02\x00' + bytes(str('\0'), 'utf-8') + tail_str)
        ser.write(b'\xaa\x42\x00\x00\x02\x80' + bytes(str('\0'), 'utf-8') + tail_str)
        ser.write(b'\xaa\x42\x00\x00\x03\x00' + bytes(str('\0'), 'utf-8') + tail_str)

        ser.write(b'\xaa\x3d\x00\x08\x00\x00\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x04\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x08\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x0c\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x10\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x14\x00\x00\xcc\x33\xc3\x3c')

        ser.write(b'\xaa\x3d\x00\x08\x00\x02\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x06\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x0A\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x0E\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x12\x00\x00\xcc\x33\xc3\x3c')
        ser.write(b'\xaa\x3d\x00\x08\x00\x16\x00\x00\xcc\x33\xc3\x3c')
        # Saving Datas -------------------------------------------------------------------------------------------------
        # Settings page enable\disable ---------------------------------------------------------------------------------

        if main_add == b'\x08' and touch_add == b'\x9a' and flag != 1:
            ser.write(b'\xaa\x70\x00\x03\xcc\x33\xc3\x3c')

        # Settings page enable\disable ---------------------------------------------------------------------------------
        # Graph_Home----------------------------------------------------------------------------------------------------
        if Block_int >= 400:
            Block_int = 400
        d3 = Block_int >> 8
        d4 = Block_int & 0x00ff
        ser.write(b'\xaa\x4E\x00\x06\x00\x00\x00\x68' + d3.to_bytes(1, 'big') + d4.to_bytes(1, 'big') + tail_int)
    # Graph_Home----------------------------------------------END-------------------------------------------------------


import requests
import mimetypes
import uuid
import json
import os
from datetime import datetime
import time
import shutil


class Client:
    def __init__(self, server_url):
        self.is_connected = False
        self.server_url = server_url
        self.token = None

    def login(self, username, password):
        response = requests.post(
            f"{self.server_url}/login", json={'username': username, 'password': password})
        if response.status_code == 200:
            self.token = response.json().get('token')
            return True
        else:
            return False

    def save_json_data(self, json_data, dest_path):
        try:
            json_file_name = f"{os.path.splitext(dest_path)[0]}.json"
            with open(json_file_name, 'w') as json_file:
                json.dump(json_data, json_file)
            print(f"JSON data saved to {json_file_name}")
        except Exception as e:
            print(f"Error saving JSON data: {e}")

    def move_file(self, file_path, file_name, dest_folder, json_data):
        print('Move file')
        try:
            dest_path = os.path.join(dest_folder, file_name)
            shutil.move(file_path, dest_path)
            # you might want to save jdon_data in seperate file which will be used when syncing data later
            self.save_json_data(json_data, dest_path)
        except Exception as e:
            self.logger.log_error(f"Error moving file: {e}")

    def is_alive(self):
        response = requests.get(f"{self.server_url}/is_alive")
        return response.status_code == 200

    def retry_upload(self, file_path, max_retries, data):
        try:
            if self.is_connected == True:
                for retry_count in range(max_retries):
                    if self.upload_file(file_path, data) == 1:
                        return 1
                    time.sleep(5)
                return -1
            else:
                return 0

        except Exception as e:
            print(f"Error during upload retry: {str(e)}")

    def upload_file(self, file_name, data):
        if self.is_connected == True:
            if self.is_alive() == False:
                print('upload_file: Server is Not Reachable')
                return 0  # Server not alive

        try:
            content_type, _ = mimetypes.guess_type(file_name)
            data["Content-Type"] = content_type
            data = json.loads(json.dumps(data))

            new_file_name = file_name
            with open(new_file_name, 'rb') as file:
                file_content = file.read()

            _, file_extension = os.path.splitext(new_file_name)
            content_type, _ = mimetypes.guess_type(new_file_name)

            if content_type is None:
                content_type = 'application/octet-stream'

            boundary = str(uuid.uuid4())

            headers = {
                "Authorization": f"Bearer {self.token}",
                'Content-Type': f"multipart/form-data; boundary={boundary}",
            }

            only_file_name = os.path.basename(new_file_name)
            str_payload = (
                f"--{boundary}\r\n"
                f"Content-Disposition: form-data; name=\"data\"\r\n"
                "\r\n"
                f"{data}\r\n"
                f"--{boundary}\r\n"
                f"Content-Disposition: form-data; name=\"file\"; filename=\"{only_file_name}\"\r\n"
                f"Content-Type: {content_type}\r\n"
                "\r\n"
            )
            payload = str_payload.encode('utf-8')
            payload += file_content
            payload += f"\r\n--{boundary}--\r\n".encode('utf-8')

            print(
                f'upload_file: file: {file_name} data: {data}')

            response = requests.post(
                f"{self.server_url}/attachments", data=payload, headers=headers)

            if response.status_code == 200:
                return 1  # Success
            elif response.status_code in [312, 313, 314]:
                error_messages = {
                    312: 'JWT expired',
                    313: 'Malformed JWT',
                    314: 'Incorrect JWT'
                }
                error_message = error_messages.get(
                    response.status_code, 'JWT Error')
                print(f'upload_file: {error_message}')
                if not self.login('admin', 'password'):  # Example credentials
                    return False
                return self.upload_file(file_name, data)
            elif response.status_code in [411, 412, 414]:
                error_messages = {
                    411: 'File Size > 10MB',
                    412: 'File Size == 0(no bytes received)',
                    414: 'File extension unknown or not allowed'
                }
                error_message = error_messages.get(
                    response.status_code, 'File Error')
                print(f'upload_file: {error_message}')
                return -2  # Failure
            elif response.status_code in [413, 511, 512, 513]:
                error_messages = {
                    413: 'No data (fields) received with files',
                    511: 'Sample ID not found in Data',
                    512: 'Logic ID not found in Data',
                    513: 'Device name not found in Data'
                }
                error_message = error_messages.get(
                    response.status_code, 'Other Error')
                print(f'upload_file: {error_message}')
                return -1  # Failure
            else:
                print(
                    'upload_file: ' + f"Upload error: {response.status_code} response:{response}")
                return -1  # Failure
        except Exception as e:
            print(
                'upload_file: ' + f"Upload error: {e}")
            return -1  # Error
        
    def upload_csv_file(self, file_path, file_name):
        json_data = {
            "device_id": 'D0001',
            "device_type": 'Drop Point Tester',
            "sample_code": 'S0001',
            "upload_timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        # print(f'json_data:{json_data}')
        if self.is_connected == True:
            result = self.upload_file(file_path, json_data)
            if result == 1:
                self.delete_file()
            elif result == -1:
                retry_result = self.retry_upload(
                    file_path, 3, json_data)
                if retry_result == 1:
                    self.delete_file()
                elif retry_result == -1:
                    # Error so move to error folder or sync folder
                    self.move_file(file_path, file_name, 'errors', json_data)
                elif retry_result == 0:
                    # Wait for 5 seconds before retrying
                    time.sleep(5)
            elif result == 0:
                # Wait for 5 seconds before retrying
                time.sleep(5)
        else:  # self.is_connected == True
            print('disconnected, so do auto sync')
            # Error so move to error folder or sync folder
            self.move_file(file_path, file_name, 'sync', json_data)

if __name__ == "__main__":
    server_url = "http://192.168.6.35:5000"
    client = Client(server_url)
    cleint.upload_csv_file("D:\Dropping_Point\DATA.CSV","DATA.CSV")