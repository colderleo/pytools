import datetime
import tools_common
from tools_colorprint import print


#-*- coding:utf-8 -*-#

#filename: prt_cmd_color.py

import ctypes,sys
STD_INPUT_HANDLE = -10
STD_OUTPUT_HANDLE = -11
STD_ERROR_HANDLE = -12


def main():
    import platform

    print('test', color='red', dt=True, dt_color='green')
    print('test', color='blue', dt=True, dt_color='green')
    print('test 2', dt=True)


# get handle
std_out_handle = ctypes.windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)

def set_cmd_text_color(color, handle=std_out_handle):
    Bool = ctypes.windll.kernel32.SetConsoleTextAttribute(handle, color)
    return Bool

def mprint(write,color):
    set_cmd_text_color(color)
    sys.stdout.write(write)
    set_cmd_text_color(15)#使下一行变为白色



if __name__ == "__main__":
    main()

