import os
from os.path import join,getsize
def getdirsize(dir):
    size_sum = 0
    file_count = 0

    if os.path.isdir(dir):
        for walker in os.walk(dir):

            root=walker[0]
            files=walker[2]

            file_count += len(files)
            size_sum += sum([getsize(join(root,name))for name in files])
    
    if os.path.isfile(dir) :
        file_count = 1
        size_sum = getsize(dir)
    
    if file_count == 0 :
        file_count = 1

    return (file_count,size_sum)

if __name__ == "__main__":
    print(getdirsize('D:\\'))