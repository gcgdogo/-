import requests
import time

t_start = time.time()
print("start_time : {}".format(t_start))
r_code = 0
max_seconds = 30
print("最大等待 {} 秒".format(max_seconds))
while( r_code != 200 ):
    if (time.time() - t_start) > max_seconds:
        print("等待超时")
        break
    try:
        r_code = requests.get("http://localhost:8888/tree").status_code
    except:
        pass

print("time_elasped : {}".format(time.time() - t_start))