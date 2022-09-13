"""
It's collecting the B5K's re-convergent data, based on below rules:
1. record the period which fix type from RTX(9) to none RTX as RTX Lost Period

2. when acitve "average" value for re-convergent calculation (Avg_o_List == 0)
2.1 when it goes back to RTX from lost, skip c_skip number ahead data (from previous lost RTX position)
2.2 then start collect continuous c_anum numbers of average "Pos Accy (m)" value (accy_1) as comparision base number
2.3 check continuous c_anum numbers of average "Pos Accy (m)" value (accy_2) beginning with RTX recovered point
2.4 once accy_2 <= accy_1, consider this preriod as re-convergent time

3. when acitve "list" value for re-convergent calculation (Avg_o_List == 1)
3.1 when it goes back to RTX from lost, only compare the "Pos Accy (m)" value with the numbers in x_list
"""

xver = '0.3'
import os
import openpyxl

xlx_path = 'C:\\Work\\Tools\\rctime_from_xlx\\'
#xlx_path = '/Users/Hawk/Downloads/'
xlx_name = 'sv_example_22020831'
xlx_tail = '.xlsx'
output_tail = '.csv'
Avg_o_List  = 0 # 0: Avg, 1: List
x_list      = [0.1, 0.05, 0.02] # TBD: it has to be only 3, and has to be left one larger than right one

xlx_wb = openpyxl.load_workbook(xlx_path+xlx_name+xlx_tail)
xlx_sht = xlx_wb[xlx_wb.sheetnames[0]]

x_gpstime   = 3
x_time      = 7
x_fix       = 18
x_lon       = 9
x_lat       = 10
x_posacc    = 29
x_gps_max   = 86400
x_mark      = [0, 0, 0]
x_gps       = [0, 0, 0]
c_count     = 0
c_trig      = 9
c_flag      = 0
c_gps       = 0
c_skip      = 10
c_anum      = 10
c_acc_1     = 0
c_acc_2     = 0
c_row_fix   = 0

if Avg_o_List == 1:
    output_file = xlx_path+xlx_name+'_list'+output_tail
else:
    output_file = xlx_path+xlx_name+'_avg'+output_tail

# x: Open a file for exclusive creation. If the file already exists, the operation fails.
# w: Open a file for writing. Creates a new file if it does not exist or truncates the file if it exists.
with open(output_file, 'w') as c_log:
    c_msg = '\nRe-convergent time tool version:' + xver
    c_msg += '\n'
    c_msg += '\n err1 - fix lost again'
    c_msg += '\n err2 - avg: previous invalid previous row'
    c_msg += '\n err3 - avg: previous row too short'
    c_msg += '\n err4 - flag error'
    c_msg += '\n'
    
    if Avg_o_List == 0:
        c_msg += '\nIndex,StartLon,StartLat,StartTime,EndLon,EndLat,EndTime,RTXLostPeriod(s),StartAccy,EndAccy,ReCnvtPeriod(s)'
    else:
        c_msg += '\nIndex,StartLon,StartLat,StartTime,EndLon,EndLat,EndTime,RTXLostPeriod(s),'+str(x_list[0])+'M(s),'+str(x_list[1])+'M(s),'+str(x_list[2])+'M(s)'

    print(c_msg)
    c_log.write(str(c_msg))
    for row_1 in range(2, xlx_sht.max_row):
        if c_trig != xlx_sht.cell(row_1, x_fix).value:
            if c_flag == 0: # Never goes into fix yet
                continue
            elif c_flag == 1:# Fix lost start 
                if c_gps == 0:
                    c_count += 1
                    c_gps = xlx_sht.cell(row_1, x_gpstime).value
                    c_msg = '\n'
                    c_msg += str(c_count).zfill(3) # Index
                    c_msg += ','
                    c_msg += str(xlx_sht.cell(row_1, x_lon).value) # StartLon
                    c_msg += ','
                    c_msg += str(xlx_sht.cell(row_1, x_lat).value) # StartLat
                    c_msg += ','
                    c_msg += str(xlx_sht.cell(row_1, x_time).value) # StartTime
                    # Keep item index for later calculation
                    c_row_fix = row_1
                else:
                    continue
            # Fix lost again before finished the convergent calculate
            else: 
                if Avg_o_List == 1:
                    x_mark[0] = 0 
                    x_mark[1] = 0 
                    x_mark[2] = 0
                c_flag = 0
                c_gps = 0
                c_msg += ',err1' # Fix lost, process abort
                print(c_msg)
                c_log.write(str(c_msg))
                continue
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
                    c_msg += ','
                    c_msg += str(xlx_sht.cell(row_1, x_lon).value) # EndLon
                    c_msg += ','
                    c_msg += str(xlx_sht.cell(row_1, x_lat).value) # EndLat
                    c_msg += ','
                    c_msg += str(xlx_sht.cell(row_1, x_time).value) # EndTime
                    c_msg += ','
                    c_msg += str(format(c_gps,'.2f')) # RTXLostPeriod(s)
                    # Start the re-convergent calculation
                    c_acc_1 = 0
                    c_flag = 2 
                    # Keep it for later calculation
                    c_gps = xlx_sht.cell(row_1, x_gpstime).value
                    if Avg_o_List == 1:
                        x_gps[0] = xlx_sht.cell(row_1, x_gpstime).value
                        x_gps[1] = xlx_sht.cell(row_1, x_gpstime).value
                        x_gps[2] = xlx_sht.cell(row_1, x_gpstime).value
                    
                    if Avg_o_List == 1:
                        continue

                    # Average of c_anum numbers of data after skip c_skip numbers of data 
                    if c_row_fix == 0:
                        c_flag = 0
                        c_gps = 0
                        c_msg += ',err2' # fix row error, process abort
                        print(c_msg)
                        c_log.write(str(c_msg))
                        continue
                    elif c_row_fix < (c_skip + c_anum):
                        c_flag = 0
                        c_gps = 0
                        c_msg += ',err3' # fix row skip error, process abort
                        print(c_msg)
                        c_log.write(str(c_msg))
                        continue
                    row_2 = c_row_fix - c_skip - c_anum 
                    for row_2 in range(row_2, row_2+c_anum):
                        c_acc_1 +=  xlx_sht.cell(row_2, x_posacc).value
                    c_acc_1 /= c_anum
                    c_msg += ','
                    c_msg += str(format(c_acc_1,'.2f')) # StartAccy
            elif c_flag == 2: # Fix back, start re-convergent calculation
                if Avg_o_List == 1:
                    for i in range(0, 2):
                        if x_mark[i] == 0:
                            if x_list[i] >= xlx_sht.cell(row_1, x_posacc).value:
                                c_gps = x_gps[i]
                                # TBD: not handling c_gps == 0
                                if xlx_sht.cell(row_1, x_gpstime).value < c_gps:
                                     c_gps = x_gps_max - c_gps
                                     c_gps += xlx_sht.cell(row_1, x_gpstime).value
                                else:
                                     c_gps = xlx_sht.cell(row_1, x_gpstime).value - c_gps
                                c_msg += ','
                                c_msg += str(format(c_gps,'.2f')) 
                                x_gps[i]  = c_gps
                                x_mark[i] = 1
                    if x_mark[0] !=0 and x_mark[1] !=0 and x_mark[2] != 0:
                        x_mark[0] = 0 
                        x_mark[1] = 0 
                        x_mark[2] = 0
                    else:
                        continue
                else:
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

                        c_msg += ','
                        c_msg += str(format(c_acc_2,'.2f')) # EndAccy
                        c_msg += ',' 
                        c_msg += str(format(c_gps,'.2f')) # ReCnvtPeriod(s)
                    else:
                        continue
                print(c_msg)
                c_log.write(str(c_msg))
                c_flag = 0
                c_gps = 0
            else:
                c_flag = 0
                c_gps = 0
                c_msg += '\nerr4' # Flag error, process abort
                print(c_msg)
                c_log.write(str(c_msg))
c_msg = '\nLog path:' + output_file
print(c_msg)
