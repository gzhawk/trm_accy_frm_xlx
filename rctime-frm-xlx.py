"""
It's collecting the B5K's re-convergent data, based on below rules:
1. calcute fix type from 9 to none 9
2. after fix type from 9 to none 9 and goes back to 9, start re-convergent calculating
2.1 when it goes back to 9, skip 10 ahead data, then start collect 10 numbers of verage accy_1 data as comparision number
2.2 check the 10 numbers of accy_2 data once meet accy_2 less than previous accy_1
"""
xver = '0.1'
import os
import openpyxl

#xlx_path = 'C:\\Work\\Tools\\XLX2NMEA\\'
xlx_path = '/Users/Hawk/Downloads/'
xlx_name = '22020831_sv_example'
xlx_tail = '.xlsx'
output_tail = '.txt'

xlx_wb = openpyxl.load_workbook(xlx_path+xlx_name+xlx_tail)
xlx_sht = xlx_wb[xlx_wb.sheetnames[0]]

x_gpstime = 3
x_time    = 7
x_fix     = 18
x_lon     = 9
x_lat     = 10
x_posacc  = 29
x_gps_max = 86400
c_count        = 0
c_trig         = 9
c_flag         = 0
c_gps          = 0
c_skip         = 10
c_anum         = 10
c_acc_1        = 0
c_acc_2        = 0

# Open a file for exclusive creation. If the file already exists, the operation fails.
print('\n\nRe-convergent collector version:',xver)
with open(xlx_path+xlx_name+output_tail, 'x') as c_log:
    for row_1 in range(2, xlx_sht.max_row):
        if c_trig != xlx_sht.cell(row_1, x_fix).value:
            if c_flag == 0: # Never goes into fix yet
                continue
            elif c_flag == 1:# Fix lost start 
                if c_gps == 0:
                    c_count += 1
                    c_gps = xlx_sht.cell(row_1, x_gpstime).value
                    c_msg = str(c_count).zfill(3)
                    c_msg += ' LON'
                    c_msg += str(xlx_sht.cell(row_1, x_lon).value)
                    c_msg += ',LAT'
                    c_msg += str(xlx_sht.cell(row_1, x_lat).value)
                    c_msg += '('
                    c_msg += str(xlx_sht.cell(row_1, x_time).value)
                    c_msg += ')'
                    c_log.write(str(c_msg))
                else:
                    continue
            else: # Fix lost again before finished the convergent calculate
                c_flag = 0
                c_gps = 0
                c_msg += '...) Fix lost, process abort\n'
                print(c_msg)
                c_log.write(str(c_msg))
        else:
            if c_flag == 0: # From no fix goes into fix
                c_flag = 1
                continue
            elif c_flag == 1: # Fix lost end
                if c_gps != 0:
                    if xlx_sht.cell(row_1, x_gpstime).value < c_gps:
                        c_gps = x_gps_max - c_gps
                        c_gps += xlx_sht.cell(row_1, x_gpstime).value
                    else:
                        c_gps = xlx_sht.cell(row_1, x_gpstime).value - c_gps
                    c_msg += '~LON'
                    c_msg += str(xlx_sht.cell(row_1, x_lon).value)
                    c_msg += ',LAT'
                    c_msg += str(xlx_sht.cell(row_1, x_lat).value)
                    c_msg += '('
                    c_msg += str(xlx_sht.cell(row_1, x_time).value)
                    c_msg += ')\n    RTX lost(s):'
                    c_msg += str(c_gps)
                    c_msg += '\n'
                    print(c_msg)
                    c_log.write(str(c_msg))
                    # Start the re-convergent calculation
                    c_acc_1 = 0
                    c_flag = 2 
                    c_gps = xlx_sht.cell(row_1, x_gpstime).value
                    # TBD: not handling row_1 < (c_skip + c_anum)
                    row_2 = row_1 - c_skip - c_anum 

                    # Average of c_anum numbers of data after skip c_skip numbers of data 
                    for row_2 in range(row_2, row_2+c_anum):
                        c_acc_1 +=  xlx_sht.cell(row_2, x_posacc).value
                    c_acc_1 /= c_anum

                    c_msg = '    Accy ('
                    c_msg += str(c_acc_1) 
                    c_msg += ' to '
            elif c_flag == 2: # Fix back, re-convergent calculation
                # Start to get average of c_anum number of data
                if c_acc_1 >= xlx_sht.cell(row_1, x_posacc).value:

                    # Average of c_anum numbers of data
                    c_acc_2 = 0
                    row_2 = row_1
                    for row_2 in range(row_2, row_2+c_anum):
                        c_acc_2 +=  xlx_sht.cell(row_2, x_posacc).value
                    c_acc_2 /= c_anum

                    # Check if the average lower than expected
                    if c_acc_1 < c_acc_2:
                        continue
                    # TBD: not handling c_gps == 0
                    if xlx_sht.cell(row_1, x_gpstime).value < c_gps:
                        c_gps = x_gps_max - c_gps
                        c_gps += xlx_sht.cell(row_1, x_gpstime).value
                    else:
                        c_gps = xlx_sht.cell(row_1, x_gpstime).value - c_gps
                    c_msg += str(c_acc_2)
                    c_msg += ') ' 
                    c_msg += str(c_gps)
                    c_msg += '\n'
                    print(c_msg)
                    c_log.write(str(c_msg))
                    c_flag = 0
                    c_gps = 0
            else:
                c_flag = 0
                c_gps = 0
                c_msg += '\n----Flag error, process abort\n'
                print(c_msg)
                c_log.write(str(c_msg))
print('Log path:',xlx_path+xlx_name+output_tail)
