import json
import os

# PATH = r"\\10.154.61.233\运行管理部\兴效能"  # r"."

PATH = os.path.join(os.environ["USERPROFILE"], "Desktop")

sync_json = f"{PATH}/兴效能测试/sync.json"

moni_json = f"{PATH}/兴效能测试/moni.json"

with open(sync_json) as f:
    data = json.load(f)

for k, v in data.items():
    print(k, ":", v)


car = {
    1: 1,
    2: 2,
    3: 3,
}

car.pop()
print(car)
