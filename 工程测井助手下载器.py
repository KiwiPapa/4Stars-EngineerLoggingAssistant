# coding=utf-8
import shutil
import os
from FTP_UP_DOWN_CLASS import MyFTP
import easygui as g

# 清理所有非空文件夹和文件
def clean_dir_of_all(path):
    list = os.listdir(path)
    if len(list) != 0:
        for i in range(0, len(list)):
            path_to_clean = os.path.join(path, list[i])
            if '工程测井助手' in list[i]: # 不删除主exe
                pass
            elif 'python' in list[i]: # 删除同一目录下的文件时会调取当前目录的python3.dll，导致删除错误，因此跳过是最好的选择
                pass
            elif list[i] in ['.vs', '.git', '.idea']:
                pass
            else:
                if '.' not in list[i]:
                    shutil.rmtree(path_to_clean)  # 清理文件夹，可非空
                else:
                    os.remove(path_to_clean)  # 清理文件
    else:
        pass

if __name__ == "__main__":
    # 先检查更新
    PATH = ".\\"
    listdir = []

    for fileName in os.listdir(PATH):
        listdir.append(fileName.split('.')[-1])

    ftp = MyFTP('10.132.203.206')
    ftp.Login('zonghs', 'zonghs123')
    local_path = './'
    # local_path = r'C:\Users\YANGYI\source\repos\GC_Logging_Helper_Release'
    remote_path_part_update = '/oracle_data9/arc_data/SGI1/2016年油套管检测归档/工程测井助手最新版本(部分更新)'
    remote_path = '/oracle_data9/arc_data/SGI1/2016年油套管检测归档/工程测井助手最新版本(全部更新)'

    # 打开本地版本号
    try:
        with open(local_path + '/版本号.txt', "r") as f:
            license_str = f.read()
        local_license_date = int(license_str)

        # 打开服务器版本号
        ftp.Cwd(remote_path)
        filenames = ftp.Nlst()
        filename = '版本号.txt'
        LocalFile = local_path + '/temp/版本号.txt'
        RemoteFile = filename

        # 接收服务器上文件并写入本地文件
        if not os.path.exists(local_path + '/temp'):
            os.makedirs(local_path + '/temp')
        ftp.DownLoadFile(LocalFile, RemoteFile)

        with open(local_path + '/temp/版本号.txt', "r") as f:
            license_str = f.read()
        remote_license_date = int(license_str)

        if local_license_date < remote_license_date:
            msg = "需要更新，请选择更新模式"
            choicess_list = ["部分更新", "全局更新", "退出"]
            reply = g.choicebox(msg, choices=choicess_list)

            if reply == "部分更新":
                # clean_dir_of_all(local_path)
                ftp.DownLoadFileTree(local_path, remote_path_part_update)
            elif reply == "全部更新":
                clean_dir_of_all(local_path)
                ftp.DownLoadFileTree(local_path, remote_path)
            else:
                pass
        elif local_license_date >= remote_license_date:
            print("本地软件版本已经是最新，无需更新。")
    except:
        print('获取版本信息异常。')

    input('按回车键退出')