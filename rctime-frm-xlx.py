"""
It's collecting the B5K's re-convergent data, based on below rules:

1. record the period which fix type from RTX(9) to none RTX

2. after fix type changed and goes back to RTX, start re-convergent calculating
2.1 when it goes back to RTX, skip 10 ahead data (from previous before lost RTX accuracy)
2.2 then start collect 10 numbers of average accy_1 data as comparision base number
2.3 check the 10 numbers of average accy_2 data beginning with RTX recovered point
2.4 once meet (accy_2 < accy_1), consider this preriod as re-convergent time

for example:

time tick: 1,2,3,...,99
RTX lost at tick 25, return to RTX at tick 30
accy_1 will be average accuracy number from tick 5~14
accy_2 will be average accuracy number from tick 31~40
"""
xver = '0.2'
import os
import openpyxl

#xlx_path = 'C:\\Work\\Tools\\XLX2NMEA\\'
xlx_path = '/Users/Hawk/Downloads/'
xlx_name = '22020831_sv_example'
xlx_tail = '.xlsx'
output_tail = '.txt'

xlx_wb = openpyxl.load_workbook(xlx_path+xlx_name+xlx_tail)
xlx_sht = xlx_wb[xlx_wb.sheetnames[0]]

x_gpstime   = 3
x_time      = 7
x_fix       = 18
x_lon       = 9
x_lat       = 10
x_posacc    = 29
x_gps_max   = 86400
c_count     = 0
c_trig      = 9
c_flag      = 0
c_gps       = 0
c_skip      = 10
c_anum      = 10
c_acc_1     = 0
c_acc_2     = 0
c_row_fix   = 0

print('\n\nRe-convergent time version:',xver)
# Open a file for exclusive creation. If the file already exists, the operation fails.
#with open(xlx_path+xlx_name+output_tail, 'x') as c_log:

# Opens a file for writing. Creates a new file if it does not exist or truncates the file if it exists.
with open(xlx_path+xlx_name+output_tail, 'w') as c_log:
    for row_1 in range(2, xlx_sht.max_row):
        if c_trig != xlx_sht.cell(row_1, x_fix).value:
            if c_flag == 0: # Never goes into fix yet
                continue
            elif c_flag == 1:# Fix lost start 
                if c_gps == 0:
                    c_count += 1
                    c_gps = xlx_sht.cell(row_1, x_gpstime).value
                    c_msg = '\n'
                    c_msg += str(c_count).zfill(3)
                    c_msg += ' (LON,LAT)Time'
                    c_msg += '\n    ('
                    c_msg += str(xlx_sht.cell(row_1, x_lon).value)
                    c_msg += ','
                    c_msg += str(xlx_sht.cell(row_1, x_lat).value)
                    c_msg += ')'
                    c_msg += str(xlx_sht.cell(row_1, x_time).value)
                    # Keep it as for later calculation
                    c_row_fix = row_1
                else:
                    continue
            else: # Fix lost again before finished the convergent calculate
                c_flag = 0
                c_gps = 0
                c_msg += '...) Fix lost, process abort'
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
                    c_msg += '~('
                    c_msg += str(xlx_sht.cell(row_1, x_lon).value)
                    c_msg += ','
                    c_msg += str(xlx_sht.cell(row_1, x_lat).value)
                    c_msg += ')'
                    c_msg += str(xlx_sht.cell(row_1, x_time).value)
                    c_msg += '\n    RTX lost(s):'
                    c_msg += str(format(c_gps,'.2f'))
                    print(c_msg)
                    c_log.write(str(c_msg))
                    # Start the re-convergent calculation
                    c_acc_1 = 0
                    c_flag = 2 
                    c_gps = xlx_sht.cell(row_1, x_gpstime).value
                    # TBD: not handling c_row_fix < (c_skip + c_anum)
                    if c_row_fix == 0 or c_row_fix < (c_skip + c_anum):
                        c_flag = 0
                        c_gps = 0
                        c_msg += '\n----Pre-fix row error, process abort'
                        print(c_msg)
                        c_log.write(str(c_msg))
                        continue
                    row_2 = c_row_fix - c_skip - c_anum 

                    # Average of c_anum numbers of data after skip c_skip numbers of data 
                    for row_2 in range(row_2, row_2+c_anum):
                        c_acc_1 +=  xlx_sht.cell(row_2, x_posacc).value
                    c_acc_1 /= c_anum

                    c_msg = '\n    Accy ('
                    c_msg += str(format(c_acc_1,'.2f')) 
                    c_msg += ' to '
            elif c_flag == 2: # Fix back, start re-convergent calculation
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
                    c_msg += str(format(c_acc_2,'.2f'))
                    c_msg += ') ' 
                    c_msg += str(format(c_gps,'.2f'))
                    print(c_msg)
                    c_log.write(str(c_msg))
                    c_flag = 0
                    c_gps = 0
            else:
                c_flag = 0
                c_gps = 0
                c_msg += '\n----Flag error, process abort'
                print(c_msg)
                c_log.write(str(c_msg))
print('Log path:',xlx_path+xlx_name+output_tail)
