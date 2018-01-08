# ftp_msg_json

python上传excel表格数据到ftp服务器

----------------------------------------------------------------------------------------------------------------------------------------

# 主要内容
1、python读取mysql数据

2、python解析json数据

3、python订做excel表格模板样式

4、python写入excel表格数据到指定data目录

5、python连接ftp服务器

6、python上传本地excel表格数据到ftp文件夹

7、linux crontab -e 定时任务

----------------------------------------------------------------------------------------------------------------------------------------

# 定时任务

每天8:30 执行解析json脚本

30 8 * * * /usr/local/bin/python /home/laidefa/msg_json/code/request_msg_json.py  >>/home/laidefa/msg_json/log/myjob1.txt 

每天8:40 执行上传本地excel到ftp服务器指定文件夹目录

40 8 * * * /usr/local/bin/python /home/laidefa/msg_json/code/ftp_uploadfile.py  >>/home/laidefa/msg_json/log/myjob2.txt 


----------------------------------------------------------------------------------------------------------------------------------------
# 效果展示
![python上传excel表格数据到ftp服务器](https://github.com/laidefa/ftp_msg_json/raw/master/msg_json/resource/1.png)

![python上传excel表格数据到ftp服务器](https://github.com/laidefa/ftp_msg_json/raw/master/msg_json/resource/2.png)





----------------------------------------------------------------------------------------------------------------------------------------
# 联系我

微信：laidefa

CSDN博客： http://blog.csdn.net/u013421629?viewmode=contents




