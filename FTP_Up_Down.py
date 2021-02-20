# coding=utf-8
import shutil
import os
import ftplib
import socket
import time


class myFTP:
    ftp = ftplib.FTP()

    def __init__(self, host, port=21):
        self.ftp.connect(host, port)

    def Login(self, user, passwd):
        self.ftp.login(user, passwd)
        print(self.ftp.welcome)

    def Cwd(self, path):
        self.ftp.cwd(path)

    def Nlst(self):
        return self.ftp.nlst()

    def DownLoadFile(self, LocalFile, RemoteFile):  # 下载单个文件
        file_handler = open(LocalFile, 'wb')
        print(file_handler)
        # self.ftp.retrbinary("RETR %s" % (RemoteFile), file_handler.write)#接收服务器上文件并写入本地文件
        self.ftp.retrbinary('RETR ' + RemoteFile, file_handler.write)
        file_handler.close()
        return True

    def DownLoadFileTree(self, LocalDir, RemoteDir):  # 下载整个目录下的文件
        print("remoteDir:", RemoteDir)
        if not os.path.exists(LocalDir):
            os.makedirs(LocalDir)
        self.ftp.cwd(RemoteDir)
        RemoteNames = self.ftp.nlst()
        print("RemoteNames", RemoteNames)
        for file in RemoteNames:
            Local = os.path.join(LocalDir, file)
            print(self.ftp.nlst(file))
            if self.ftp.nlst(file) == []: # 空文件夹情况
                if not os.path.exists(Local):
                    os.makedirs(Local)
            else:
                if self.ftp.nlst(file)[0] != file: # 取nlst后不为本身，说明为目录(有瑕疵)
                    if not os.path.exists(Local):
                        os.makedirs(Local)
                    try:
                        self.DownLoadFileTree(Local, file)
                    except:
                        print('Error downloading directory')
                else:
                    try:
                        self.DownLoadFile(Local, file)
                    except:
                        print('Error downloading file')
        self.ftp.cwd("..")
        return

    def UpLoadFile(self, Local, File):
        if os.path.isfile(Local) == False:
            return False
        file_handler = open(Local, "rb")
        self.ftp.storbinary('STOR %s' % File, file_handler, 4096)  # 上传文件
        file_handler.close()
        return True

    def UpLoadFileTree(self, LocalDir, RemoteDir):
        if os.path.isdir(LocalDir) == False:
            return False
        print("LocalDir:", LocalDir)
        LocalNames = os.listdir(LocalDir)
        print("list:", LocalNames)
        print(RemoteDir)
        self.ftp.cwd(RemoteDir)
        for Local in LocalNames:
            src = os.path.join(LocalDir, Local)
            if os.path.isdir(src):
                if os.path.isdir(LocalDir) == True:
                    RemoteDir = self.ftp.pwd()
                    self.ftp.mkd(RemoteDir + '/' + Local)
                self.UpLoadFileTree(src, Local)
            else:
                self.UpLoadFile(src, Local)
        self.ftp.cwd("..")
        return

    def Mkd(self, path):
        self.ftp.mkd(path)

    def close(self):
        self.ftp.quit()

# 清理所有非空文件夹和文件
def clean_dir_of_all(path):
    list = os.listdir(path)
    if len(list) != 0:
        for i in range(0, len(list)):
            path_to_clean = os.path.join(path, list[i])
            if '.' not in list[i]:
                shutil.rmtree(path_to_clean)  # 清理文件夹，可非空
            else:
                os.remove(path_to_clean)  # 清理文件
    else:
        pass

if __name__ == "__main__":
    ftp = myFTP('10.132.203.206')
    ftp.Login('zonghs', 'zonghs123')
    local_path = './WorkSpace'
    # local_path = r'C:\Users\YANGYI\source\repos\GC_Logging_Helper_Release'
    remote_path = '/oracle_data9/arc_data/SGI1/2016年油套管检测归档/工区备份'

    # 备份文件夹改名
    myname = socket.getfqdn(socket.gethostname())  # 获取本机电脑名
    myaddr = socket.gethostbyname(myname)  # 获取本机ip
    myaddr = myaddr.replace('.', '-')
    timeStr = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    timeStr = timeStr.replace(':', '-').replace(' ', '-')
    print(remote_path + '/' + timeStr + '_' + myaddr + '_' + myname)
    ftp.Mkd(remote_path + '/' + timeStr + '_' + myaddr + '_' + myname)

    ftp.UpLoadFileTree(local_path, remote_path + '/' + timeStr + '_' + myaddr + '_' + myname)