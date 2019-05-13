import Get_TargetList
import TargetFilter_and_MannageHistory
import Call_AVP

#获取筛选后文件列表
df_TargetList = TargetFilter_and_MannageHistory.Target_Filter(
    Get_TargetList.get_AllTarget_FromDiskList()
)

Call_AVP.Call_AVP(df_TargetList)

print("扫描完成：存储历史记录")

TargetFilter_and_MannageHistory.Save_History(df_TargetList)