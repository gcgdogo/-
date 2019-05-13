import pandas
import Get_PathSize


Preset_Excel_FileName="Scan_History.xlsx"
Preset_Excel_SheetName="Scan_History"

Preset_MaxFiles = 50000
Preset_MaxSize = 10000 * 1000000

df_Scan_History = pandas.read_excel(Preset_Excel_FileName,Preset_Excel_SheetName)
Scan_StartTime = pandas.Timestamp.now()

def Target_Filter(Target_List):

    print('开始筛选目录')

    df_TargetPath_History = df_Scan_History[['Scan_Target','Scan_FinishTime']].groupby('Scan_Target').max()

    #TargetLastTime = pandas.Timestamp('2000-1-1')
#
#    if len(df_TargetPath_History)>=1 :
#        TargetLastTime = df_TargetPath_History['Scan_FinishTime'].max()

    #获取TargetList的时间
    df_TargetList = pandas.DataFrame({'Scan_Target':Target_List})
    df_TargetList = df_TargetList.join(df_TargetPath_History,on='Scan_Target')
    df_TargetList['Scan_FinishTime'].fillna(pandas.Timestamp('2000-1-1'),inplace=True)
    df_TargetList.rename(columns = {'Scan_FinishTime':'Scan_LastTime'},inplace=True)
    
    #重新排序
    df_TargetList.sort_values(['Scan_LastTime','Scan_Target'],inplace=True)
    df_TargetList.reset_index(drop=True,inplace=True)

    df_TargetList['Target_Files']=0
    df_TargetList['Target_Size']=0

    for Target_Path in df_TargetList['Scan_Target']:

        print('    检测目录： {}'.format(Target_Path))

        (Target_Files , Target_Size) = Get_PathSize.getdirsize(Target_Path)
        if (
            df_TargetList['Target_Files'].sum()==0 
        or 
            (
                df_TargetList['Target_Files'].sum() +Target_Files < Preset_MaxFiles and 
                df_TargetList['Target_Size'].sum() + Target_Size < Preset_MaxSize
            )
        ) :
            df_TargetList.loc[df_TargetList['Scan_Target']==Target_Path,'Target_Files'] = Target_Files
            df_TargetList.loc[df_TargetList['Scan_Target']==Target_Path,'Target_Size'] = Target_Size
        else:
            break
    
    return df_TargetList.loc[
        df_TargetList['Target_Files'] > 0
    ]



def Save_History(df_Scaned):
    df_Scaned = df_Scaned[['Scan_Target','Target_Files','Target_Size']]
    df_Scaned['Scan_StartTime'] = Scan_StartTime
    df_Scaned['Scan_FinishTime'] = pandas.Timestamp.now()

    df_Scan_History_After = df_Scan_History.append(df_Scaned,ignore_index=True,sort=False)

    print( df_Scan_History_After)

    df_Scan_History_After.to_excel(
        Preset_Excel_FileName,
        Preset_Excel_SheetName,
        index=False
    )


if __name__ == "__main__":
    #print([Scan_StartTime,Preset_MaxFiles,Preset_MaxSize])
    #print(df_Scan_History)
    print(
        Target_Filter(['D:\\Code Icon x16_x64.ico', 'D:\\Config.Msi', 'D:\\DriverBackup', 'D:\\drivers2.ico', 'D:\\eMule', 'D:\\Excel', 'D:\\feiq', 'D:\\FFOutput', 'D:\\Hard disk.ico', 'D:\\MediaID.bin', 'D:\\Microsoft Visual Basic'])
    )