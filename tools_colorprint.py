# -*- encoding: utf-8 -*-
"""
    from: https://github.com/rembish/colorprint
    Module colorprint
    ~~~~~~~~~~~~~~
    Python module to print in color using py3k-style print function. It uses
    funny hack, which allow to create print function instead standard print
    routine and give it some "black" magic.
    Print function is like imported for __future__, but has three additional
    parameters: color (foreground of text output), background (it's background)
    and format (bold, blink and so on).
    You can read more at __future__.print_function documentation.
    Usage example
    -------------
        >>> from __future__ import print_function
        >>> from colorprint import print
        >>> print('Hello', 'world', color='blue', end='', sep=', ')
        >>> print('!', color='red', format=['bold', 'blink'])
        Hello, world!
        ^-- blue    ^-- blinking, bold and red
    :copyright: 2012 Aleksey Rembish
    :license: BSD
"""
from __future__ import print_function
import datetime, platform, sys, ctypes

try:
    import __builtin__
except ImportError:
    import builtins as __builtin__
    basestring = str


__all__ = ['print']

__author__ = 'Aleksey Rembish'
__email__ = 'alex@rembish.ru'
__description__ = 'Python module to print in color using py3k-style print function'
__url__ = 'https://github.com/don-ramon/colorprint'
__copyright__ = '(c) 2012 %s' % __author__
__license__ = 'BSD'
__version__ = '0.1'


linux_colors = {
    'grey': 30,  'red': 31,
    'green': 32, 'yellow': 33,
    'blue': 34,  'magenta': 35,
    'cyan': 36,  'white': 37,
}

linux_backgrounds = {
    'grey': 40,  'red': 41,
    'green': 42, 'yellow': 43,
    'blue': 44,  'magenta': 45,
    'cyan': 46,  'white': 47,
}

linux_formats = {
    'bold': 1, 'dark': 2,
    'underline': 4, 'blink': 5,
    'reverse': 7, 'concealed': 8,
}


# Windows CMD命令行 字体颜色定义 text colors
FOREGROUND_BLACK = 0x00 # black.
FOREGROUND_DARKBLUE = 0x01 # dark blue.
FOREGROUND_DARKGREEN = 0x02 # dark green.
FOREGROUND_DARKSKYBLUE = 0x03 # dark skyblue.
FOREGROUND_DARKRED = 0x04 # dark red.
FOREGROUND_DARKPINK = 0x05 # dark pink.
FOREGROUND_DARKYELLOW = 0x06 # dark yellow.
FOREGROUND_DARKWHITE = 0x07 # dark white.
FOREGROUND_DARKGRAY = 0x08 # dark gray.
FOREGROUND_BLUE = 0x09 # blue.
FOREGROUND_GREEN = 0x0a # green.
FOREGROUND_SKYBLUE = 0x0b # skyblue.
FOREGROUND_RED = 0x0c # red.
FOREGROUND_PINK = 0x0d # pink.
FOREGROUND_YELLOW = 0x0e # yellow.
FOREGROUND_WHITE = 0x0f # white.

windows_colors = {
    'red': FOREGROUND_RED,
    'green': FOREGROUND_GREEN, 
    'yellow': FOREGROUND_YELLOW,
    'blue': FOREGROUND_BLUE,
    'white': FOREGROUND_WHITE,
    'black': FOREGROUND_BLACK,
}

# windows cmd param
STD_INPUT_HANDLE = -10
STD_OUTPUT_HANDLE = -11
STD_ERROR_HANDLE = -12


is_windows = bool(platform.system().lower() == 'windows')

# get handle
std_out_handle = ctypes.windll.kernel32.GetStdHandle(STD_OUTPUT_HANDLE)

def set_cmd_text_color(color, handle=std_out_handle):
    Bool = ctypes.windll.kernel32.SetConsoleTextAttribute(handle, color)
    return Bool

#reset white
def resetColor():
    set_cmd_text_color(FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE)


def print(*args, sep=' ', end='\n', file=sys.stdout, color=None, background=None, formats=[],
        dt:bool=False, dt_color='yellow', **kwargs):
    '''
    print(value, ..., sep=' ', end='\n', file=sys.stdout, color=None, background=None, format=None)
    Prints the values to a stream, or to sys.stdout by default.
    Optional keyword arguments:
    file: a file-like object (stream); defaults to the current sys.stdout.
    sep:  string inserted between values, default a space.
    end:  string appended after the last value, default a newline.
    Additional keyword arguments:
    color: prints values in specified color:
        grey red green yellow blue magenta cyan white
    background: prints values on specified color (same as color)
    format: prints values using specifiend format(s) (can be string or list):
        bold dark underline reverse concealed
    dt:datetime'''

    if isinstance(formats, basestring):
        formats = [formats]

    if dt:
        now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f'{now_str} ', end='', color=dt_color, flush=True)

    if color or background:
        kwargs['end'] = ""

        if is_windows:
            if color:
                set_cmd_text_color(windows_colors[color])

            sys.stdout.write(*args)
            __builtin__.print(file=file, end=end, flush=True)
            resetColor()

        else:
            if color:
                __builtin__.print(f'\033[{linux_colors[color]}m', file=file, end='')
            if background:
                __builtin__.print(f'\033[{linux_backgrounds[background]}m', file=file, end='')
            for fmt in formats:
                __builtin__.print(f'\033[{linux_formats[fmt]}m', file=file, end='')

            __builtin__.print(*args, **kwargs)
            __builtin__.print('\033[0m', file=file, end=end)

    else:
        __builtin__.print(*args, **kwargs)





if __name__ == '__main__':
    print('Hello', 'world', color='white', background='blue', format='underline', end='', sep=', ')
    print('!', color='red', format=['bold', 'blink'])