import string
import os

def get_disklist():
    disk_list = []
    for c in string.ascii_uppercase:
        disk = c+':\\'
        if os.path.isdir(disk):
            disk_list.append(disk)
    return disk_list

def get_TagetInDisk(Target_Path):
    Target_List=[]
    for file in os.listdir(Target_Path):
        Target_List.append(Target_Path + file)
    return Target_List

def get_AllTarget_FromDiskList():
    DiskList=get_disklist()
    Target_List=[]

    print('获取磁盘一级目录列表')
    
    for Target_Disk in DiskList:
        Target_List.extend(get_TagetInDisk(Target_Disk))
    return Target_List

if __name__ == '__main__':
    print(get_AllTarget_FromDiskList())