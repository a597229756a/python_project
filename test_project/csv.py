import requests
import sys

print(sys.path)

url = "https://i0.hdslb.com/bfs/sycp/creative_img/202506/38a9d850eb90c4cfe8eb0976dab5fa1c.jpg"
header = {
    "User-Agent": "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:135.0) Gecko/20100101 Firefox/135.0"
}
res = requests.get(url=url, headers=header)

