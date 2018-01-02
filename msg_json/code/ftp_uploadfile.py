#coding: utf-8
import os
from ftplib import FTP
import datetime
today=datetime.date.today()
yesterday = today - datetime.timedelta(days=1)
print yesterday
import time
time1=time.time()

def ftpconnect(host, username, password):
    ftp = FTP()
    #ftp.set_debuglevel(2)         #打开调试级别2，显示详细信息
    ftp.connect(host, 21)          #连接
    ftp.login(username, password)  #登录，如果匿名登录则用空串代替即可
    return ftp

def downloadfile(ftp, remotepath, localpath):
    bufsize = 1024                #设置缓冲块大小
    fp = open(localpath,'wb')     #以写模式在本地打开文件
    ftp.retrbinary('RETR ' + remotepath, fp.write, bufsize) #接收服务器上文件并写入本地文件
    ftp.set_debuglevel(0)         #关闭调试
    fp.close()                    #关闭文件

def uploadfile(ftp, remotepath, localpath):
    bufsize = 1024
    fp = open(localpath, 'rb')
    ftp.storbinary('STOR '+ remotepath , fp, bufsize) #上传文件
    ftp.set_debuglevel(0)
    fp.close()

# 使用os模块walk函数，搜索出某目录下的全部excel文件
######################获取同一个文件夹下的所有excel文件名#######################
def getFileName(filepath):
    file_list = []
    for root, dirs, files in os.walk(filepath):
        for filespath in files:
            # print(os.path.join(root, filespath))
            file_list.append(os.path.join(root, filespath))

    return file_list

if __name__ == "__main__":

	#依次填入ftp连接的 ip、账号、密码
    ftp = ftpconnect("XXXXXXX", "XXXX", "XXXXX")
    ftp.cwd('微农贷')
    listpath = ftp.nlst()  # 获得目录列表
    if str(yesterday) in listpath:
        print "目录已创建，无需再建....."
    else:
        ftp.mkd(str(yesterday))


    ftp.cwd(str(yesterday))
    #########设置本地读取文件路径##############
    filepath='/home/laidefa/msg_json/data/%s/' %yesterday

    file_list = getFileName(filepath)
    print len(file_list)
    for each in file_list:
        print each
        localfile=each

        remotepath=os.path.basename(localfile)

        uploadfile(ftp, remotepath, localfile)

    ftp.quit()

    time2 = time.time()
    print 'ok,上传FTP成功!'
    print '总共耗时：' + str(time2 - time1) + 's'
