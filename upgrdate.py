import os
import sys
import subprocess 

#编写bat脚本，删除旧程序，运行新程序
def WriteRestartCmd(exe_name):
    b = open("upgrade.bat",'w')
    TempList = "@echo off\n";   #关闭bat脚本的输出
    TempList += "if not exist "+exe_name+" exit \n";    #新文件不存在,退出脚本执行
    TempList += "sleep 3\n" #3秒后删除旧程序（3秒后程序已运行结束，不延时的话，会提示被占用，无法删除）
    TempList += "del "+ os.path.realpath(sys.argv[0]) + "\n"    #删除当前文件
    TempList += "start " + exe_name     #启动新程序
    b.write(TempList)
    b.close()
    subprocess.Popen("upgrade.bat")
    sys.exit()  #进行升级，退出此程序

def main():
#新程序启动时，删除旧程序制造的脚本
    if os.path.isfile("upgrade.bat"):
        os.remove("upgrade.bat")
    WriteRestartCmd("newVersion.exe")

if __name__ == '__main__':
    main()
    sys.exit()