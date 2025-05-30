FROM python:3.12.3

WORKDIR /usr/src/app

COPY . .

# pip freeze > requirements.txt 将包写入到文件中
# pip install -r requirements.txt 安装文件中的包
RUN pip install --no-cache-dir -r requirements.txt

CMD ["python","./main.py"]


# 构建一个app的镜像
# docker build -t app_image .

# 以app镜像为容器运行
# docker run -d -p 5000:5000 --name app_container app_image

# 进入容器内部进行操作 
# docker exec -it app_container /bin/bash

# 停止并删除容器
# docker stop app_container && docker rm app_container
 
# 删除镜像
# docker rmi app_image