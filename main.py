import datetime
import tools_common
import coloredlogs
from tools_colorprint import print


#-*- coding:utf-8 -*-#

#filename: prt_cmd_color.py

import ctypes,sys
STD_INPUT_HANDLE = -10
STD_OUTPUT_HANDLE = -11
STD_ERROR_HANDLE = -12


def main():
    import platform

    print('test', color='red')
    print('test2', color='blue')
    print('test3', color='blue', dt=True, dt_color='green')

    print('test', color='red', end='')
    print('test2', color='blue', end='')



# get handle
std_out_handle = ctypes.windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)



if __name__ == "__main__":
    main()

