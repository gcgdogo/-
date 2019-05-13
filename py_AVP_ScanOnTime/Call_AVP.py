import os
import time

def Call_AVP(df_TargetList):
    print(df_TargetList)

    # File_ScanList = open('ScanList.txt',"w+")
    # File_ScanList.writelines(df_TargetList['Scan_Target'] + "\n")
    # File_ScanList.close

    Str_ScanList=""

    for Scan_Target in df_TargetList['Scan_Target']:
        Str_ScanList = Str_ScanList + " \"" + Scan_Target + "\""


    #单文件限时一分钟，跳过过大的压缩包
    Str_Command = 'avp.com SCAN -e:60 ' + Str_ScanList

    print("Shell: " + Str_Command)
    os.system(Str_Command)