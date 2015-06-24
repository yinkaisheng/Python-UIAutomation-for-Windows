# -*- coding:utf-8 -*-
'''
Author: yinkaisheng
Mail: yinkaisheng@foxmail.com
QQ: 396230688

This module is for automation on Windows(Windows XP with SP3, Windows Vista, Windows 7 and Windows 8/8.1).
It supports automation for the applications which implmented IUIAutomation, such as MFC, Windows Forms, WPF, Windows 8 Metro App, Qt.
The details: http://www.cnblogs.com/Yinkaisheng/p/3444132.html
Run 'automation.py -h' for help.
'''
import sys
import os
import time
import ctypes
import ctypes.wintypes
IsPy3 = sys.version > '3'
if not IsPy3:
    import codecs
#from platform import win32_ver

# version = win32_ver() # ('8', '6.2.9200', '', 'Multiprocessor Free')

AUTHOR_MAIL = 'yinkaisheng@foxmail.com'
MSG_CAPTION = 'Tip'
METRO_WINDOW_CLASS_NAME = 'Windows.UI.Core.CoreWindow'
SEARCH_INTERVAL = 0.5 # search control interval seconds
MAX_MOVE_SECOND = 1 # simulate mouse move or drag max seconds

class AutomationClient():
    def __init__(self):
        dir = os.path.dirname(__file__)
        if dir:
            oldDir = os.getcwd()
            os.chdir(dir)
        self.dll = ctypes.cdll.AutomationClient
        if dir:
            os.chdir(oldDir)
        if not self.dll.InitInstance():
            raise RuntimeError('Can not get an instance of IUIAutomation')

    def __del__(self):
        self.dll.ReleaseInstance()

ClientObject = AutomationClient()

'''
class WString():
    def __init__(self):
        self.wcharArray = 0
        self.cwcharp = 0
    def __del__(self):
        if self.wcharArray:
            ClientObject.dll.DeleteWCharArray(self.wcharArray)
    @property
    def value(self):
        if self.cwcharp:
            return self.cwcharp.value
'''

class ControlType():
    '''This class defines the values of control type'''
    ButtonControl = 0xc350
    CalendarControl = 0xc351
    CheckBoxControl = 0xc352
    ComboBoxControl = 0xc353
    CustomControl = 0xc369
    DataGridControl = 0xc36c
    DataItemControl = 0xc36d
    DocumentControl = 0xc36e
    EditControl = 0xc354
    GroupControl = 0xc36a
    HeaderControl = 0xc372
    HeaderItemControl = 0xc373
    HyperlinkControl = 0xc355
    ImageControl = 0xc356
    ListControl = 0xc358
    ListItemControl = 0xc357
    MenuBarControl = 0xc35a
    MenuControl = 0xc359
    MenuItemControl = 0xc35b
    PaneControl = 0xc371
    ProgressBarControl = 0xc35c
    RadioButtonControl = 0xc35d
    ScrollBarControl = 0xc35e
    SemanticZoomControl = 0xc377
    SeparatorControl = 0xc376
    SliderControl = 0xc35f
    SpinnerControl = 0xc360
    SplitButtonControl = 0xc36f
    StatusBarControl = 0xc361
    TabControl = 0xc362
    TabItemControl = 0xc363
    TableControl = 0xc374
    TextControl = 0xc364
    ThumbControl = 0xc36b
    TitleBarControl = 0xc375
    ToolBarControl = 0xc365
    ToolTipControl = 0xc366
    TreeControl = 0xc367
    TreeItemControl = 0xc368
    WindowControl = 0xc370

ControlTypeNameDict = {
    ControlType.ButtonControl : 'ButtonControl',
    ControlType.CalendarControl : 'CalendarControl',
    ControlType.CheckBoxControl : 'CheckBoxControl',
    ControlType.ComboBoxControl : 'ComboBoxControl',
    ControlType.CustomControl : 'CustomControl',
    ControlType.DataGridControl : 'DataGridControl',
    ControlType.DataItemControl : 'DataItemControl',
    ControlType.DocumentControl : 'DocumentControl',
    ControlType.EditControl : 'EditControl',
    ControlType.GroupControl : 'GroupControl',
    ControlType.HeaderControl : 'HeaderControl',
    ControlType.HeaderItemControl : 'HeaderItemControl',
    ControlType.HyperlinkControl : 'HyperlinkControl',
    ControlType.ImageControl : 'ImageControl',
    ControlType.ListControl : 'ListControl',
    ControlType.ListItemControl : 'ListItemControl',
    ControlType.MenuBarControl : 'MenuBarControl',
    ControlType.MenuControl : 'MenuControl',
    ControlType.MenuItemControl : 'MenuItemControl',
    ControlType.PaneControl : 'PaneControl',
    ControlType.ProgressBarControl : 'ProgressBarControl',
    ControlType.RadioButtonControl : 'RadioButtonControl',
    ControlType.ScrollBarControl : 'ScrollBarControl',
    ControlType.SemanticZoomControl : 'SemanticZoomControl',
    ControlType.SeparatorControl : 'SeparatorControl',
    ControlType.SliderControl : 'SliderControl',
    ControlType.SpinnerControl : 'SpinnerControl',
    ControlType.SplitButtonControl : 'SplitButtonControl',
    ControlType.StatusBarControl : 'StatusBarControl',
    ControlType.TabControl : 'TabControl',
    ControlType.TabItemControl : 'TabItemControl',
    ControlType.TableControl : 'TableControl',
    ControlType.TextControl : 'TextControl',
    ControlType.ThumbControl : 'ThumbControl',
    ControlType.TitleBarControl : 'TitleBarControl',
    ControlType.ToolBarControl : 'ToolBarControl',
    ControlType.ToolTipControl : 'ToolTipControl',
    ControlType.TreeControl : 'TreeControl',
    ControlType.TreeItemControl : 'TreeItemControl',
    ControlType.WindowControl : 'WindowControl',
    }

class PatternId():
    '''This class defines the values of pattern id'''
    UIA_AnnotationPatternId = 0x2727
    UIA_DockPatternId = 0x271b
    UIA_DragPatternId = 0x272e
    UIA_DropTargetPatternId = 0x272f
    UIA_ExpandCollapsePatternId = 0x2715
    UIA_GridItemPatternId = 0x2717
    UIA_GridPatternId = 0x2716
    UIA_InvokePatternId = 0x2710
    UIA_ItemContainerPatternId = 0x2723
    UIA_LegacyIAccessiblePatternId = 0x2722
    UIA_MultipleViewPatternId = 0x2718
    UIA_ObjectModelPatternId = 0x2726
    UIA_RangeValuePatternId = 0x2713
    UIA_ScrollItemPatternId = 0x2721
    UIA_ScrollPatternId = 0x2714
    UIA_SelectionItemPatternId = 0x271a
    UIA_SelectionPatternId = 0x2711
    UIA_SpreadsheetItemPatternId = 0x272b
    UIA_SpreadsheetPatternId = 0x272a
    UIA_StylesPatternId = 0x2729
    UIA_SynchronizedInputPatternId = 0x2725
    UIA_TableItemPatternId = 0x271d
    UIA_TablePatternId = 0x271c
    UIA_TextChildPatternId = 0x272d
    UIA_TextPattern2Id = 0x2728
    UIA_TextPatternId = 0x271e
    UIA_TogglePatternId = 0x271f
    UIA_TransformPattern2Id = 0x272c
    UIA_TransformPatternId = 0x2720
    UIA_ValuePatternId = 0x2712
    UIA_VirtualizedItemPatternId = 0x2724
    UIA_WindowPatternId = 0x2719

PatternDict = {
    PatternId.UIA_DockPatternId : 'DockPattern',
    PatternId.UIA_ExpandCollapsePatternId : 'ExpandCollapsePattern',
    PatternId.UIA_GridItemPatternId : 'GridItemPattern',
    PatternId.UIA_GridPatternId : 'GridPattern',
    PatternId.UIA_InvokePatternId : 'InvokePattern',
    PatternId.UIA_ItemContainerPatternId : 'ItemContainerPattern',
    PatternId.UIA_LegacyIAccessiblePatternId : 'LegacyIAccessiblePattern',
    PatternId.UIA_MultipleViewPatternId : 'MultipleViewPattern',
    PatternId.UIA_RangeValuePatternId : 'RangeValuePattern',
    PatternId.UIA_ScrollItemPatternId : 'ScrollItemPattern',
    PatternId.UIA_ScrollPatternId : 'ScrollPattern',
    PatternId.UIA_SelectionItemPatternId : 'SelectionItemPattern',
    PatternId.UIA_SelectionPatternId : 'SelectionPattern',
    PatternId.UIA_SynchronizedInputPatternId : 'SynchronizedInputPattern',
    PatternId.UIA_TableItemPatternId : 'TableItemPattern',
    PatternId.UIA_TablePatternId : 'TablePattern',
    PatternId.UIA_TextPatternId : 'TextPattern',
    PatternId.UIA_TogglePatternId : 'TogglePattern',
    PatternId.UIA_TransformPatternId : 'TransformPattern',
    PatternId.UIA_ValuePatternId : 'ValuePattern',
    PatternId.UIA_VirtualizedItemPatternId : 'VirtualizedItemPattern',
    PatternId.UIA_WindowPatternId : 'WindowPattern'}

class Point(ctypes.Structure):
    _fields_ = [("x", ctypes.c_ulong), ("y", ctypes.c_ulong)]

class Coord(ctypes.Structure):
    _fields_ = [('X', ctypes.c_short), ('Y', ctypes.c_short)]

class SmallRect(ctypes.Structure):
    _fields_ = [('Left', ctypes.c_short),
               ('Top', ctypes.c_short),
               ('Right', ctypes.c_short),
               ('Bottom', ctypes.c_short),
              ]

class Rect(ctypes.Structure):
    _fields_ = [('left', ctypes.c_int),
               ('top', ctypes.c_int),
               ('right', ctypes.c_int),
               ('bottom', ctypes.c_int),
              ]

class ConsoleScreenBufferInfo(ctypes.Structure):
    _fields_ = [('dwSize', Coord),
               ('dwCursorPosition', Coord),
               ('wAttributes', ctypes.c_uint),
               ('srWindow', SmallRect),
               ('dwMaximumWindowSize', Coord),
              ]

class tagPROCESSENTRY32(ctypes.Structure):
    _fields_ = [('dwSize',              ctypes.wintypes.DWORD),
                ('cntUsage',            ctypes.wintypes.DWORD),
                ('th32ProcessID',       ctypes.wintypes.DWORD),
                ('th32DefaultHeapID',   ctypes.POINTER(ctypes.wintypes.ULONG)),
                ('th32ModuleID',        ctypes.wintypes.DWORD),
                ('cntThreads',          ctypes.wintypes.DWORD),
                ('th32ParentProcessID', ctypes.wintypes.DWORD),
                ('pcPriClassBase',      ctypes.wintypes.LONG),
                ('dwFlags',             ctypes.wintypes.DWORD),
                ('szExeFile',           ctypes.c_wchar * 260)
               ]

class MSG(ctypes.Structure):
    _fields_=[("hwnd",      ctypes.c_uint),
              ("message",   ctypes.c_uint),
              ("wParam",    ctypes.c_uint),
              ("lParam",    ctypes.c_uint),
              ("time",      ctypes.c_uint),
              ("pt",        ctypes.c_ulonglong)]
              
class MouseEventFlags():
    '''This class defines the MouseEventFlags from Win32'''
    Absolute = 0x8000
    LeftDown = 0x0002
    LeftUp = 0x0004
    MiddleDown = 0x0020
    MiddleUp = 0x0040
    Move = 0x0001
    RightDown = 0x0008
    RightUp = 0x0010
    Wheel = 0x0800

class KeyboardEventFlags():
    '''This class defines the KeyboardEventFlags from Win32'''
    KeyDown = 0x0000
    ExtendedKey = 0x0001
    KeyUp = 0x0002

class HotKey():
    MOD_ALT = 0x0001
    MOD_CONTROL = 0x0002
    MOD_SHIFT = 0x0004
    
class Keys():
    '''This class defines the Key Code from Win32'''
    VK_LBUTTON = 0x01                       #Left mouse button
    VK_RBUTTON = 0x02                       #Right mouse button
    VK_CANCEL = 0x03                        #Control-break processing
    VK_MBUTTON = 0x04                       #Middle mouse button (three-button mouse)
    VK_XBUTTON1 = 0x05                      #X1 mouse button
    VK_XBUTTON2 = 0x06                      #X2 mouse button
    VK_BACK = 0x08                          #BACKSPACE key
    VK_TAB = 0x09                           #TAB key
    VK_CLEAR = 0x0C                         #CLEAR key
    VK_RETURN = 0x0D                        #ENTER key
    VK_SHIFT = 0x10                         #SHIFT key
    VK_CONTROL = 0x11                       #CTRL key
    VK_MENU = 0x12                          #ALT key
    VK_PAUSE = 0x13                         #PAUSE key
    VK_CAPITAL = 0x14                       #CAPS LOCK key
    VK_KANA = 0x15                          #IME Kana mode
    VK_HANGUEL = 0x15                       #IME Hanguel mode (maintained for compatibility; use VK_HANGUL)
    VK_HANGUL = 0x15                        #IME Hangul mode
    VK_JUNJA = 0x17                         #IME Junja mode
    VK_FINAL = 0x18                         #IME final mode
    VK_HANJA = 0x19                         #IME Hanja mode
    VK_KANJI = 0x19                         #IME Kanji mode
    VK_ESCAPE = 0x1B                        #ESC key
    VK_CONVERT = 0x1C                       #IME convert
    VK_NONCONVERT = 0x1D                    #IME nonconvert
    VK_ACCEPT = 0x1E                        #IME accept
    VK_MODECHANGE = 0x1F                    #IME mode change request
    VK_SPACE = 0x20                         #SPACEBAR
    VK_PRIOR = 0x21                         #PAGE UP key
    VK_NEXT = 0x22                          #PAGE DOWN key
    VK_END = 0x23                           #END key
    VK_HOME = 0x24                          #HOME key
    VK_LEFT = 0x25                          #LEFT ARROW key
    VK_UP = 0x26                            #UP ARROW key
    VK_RIGHT = 0x27                         #RIGHT ARROW key
    VK_DOWN = 0x28                          #DOWN ARROW key
    VK_SELECT = 0x29                        #SELECT key
    VK_PRINT = 0x2A                         #PRINT key
    VK_EXECUTE = 0x2B                       #EXECUTE key
    VK_SNAPSHOT = 0x2C                      #PRINT SCREEN key
    VK_INSERT = 0x2D                        #INS key
    VK_DELETE = 0x2E                        #DEL key
    VK_HELP = 0x2F                          #HELP key
    VK_0 = 0x30                             #0 key
    VK_1 = 0x31                             #1 key
    VK_2 = 0x32                             #2 key
    VK_3 = 0x33                             #3 key
    VK_4 = 0x34                             #4 key
    VK_5 = 0x35                             #5 key
    VK_6 = 0x36                             #6 key
    VK_7 = 0x37                             #7 key
    VK_8 = 0x38                             #8 key
    VK_9 = 0x39                             #9 key
    VK_A = 0x41                             #A key
    VK_B = 0x42                             #B key
    VK_C = 0x43                             #C key
    VK_D = 0x44                             #D key
    VK_E = 0x45                             #E key
    VK_F = 0x46                             #F key
    VK_G = 0x47                             #G key
    VK_H = 0x48                             #H key
    VK_I = 0x49                             #I key
    VK_J = 0x4A                             #J key
    VK_K = 0x4B                             #K key
    VK_L = 0x4C                             #L key
    VK_M = 0x4D                             #M key
    VK_N = 0x4E                             #N key
    VK_O = 0x4F                             #O key
    VK_P = 0x50                             #P key
    VK_Q = 0x51                             #Q key
    VK_R = 0x52                             #R key
    VK_S = 0x53                             #S key
    VK_T = 0x54                             #T key
    VK_U = 0x55                             #U key
    VK_V = 0x56                             #V key
    VK_W = 0x57                             #W key
    VK_X = 0x58                             #X key
    VK_Y = 0x59                             #Y key
    VK_Z = 0x5A                             #Z key
    VK_LWIN = 0x5B                          #Left Windows key (Natural keyboard)
    VK_RWIN = 0x5C                          #Right Windows key (Natural keyboard)
    VK_APPS = 0x5D                          #Applications key (Natural keyboard)
    VK_SLEEP = 0x5F                         #Computer Sleep key
    VK_NUMPAD0 = 0x60                       #Numeric keypad 0 key
    VK_NUMPAD1 = 0x61                       #Numeric keypad 1 key
    VK_NUMPAD2 = 0x62                       #Numeric keypad 2 key
    VK_NUMPAD3 = 0x63                       #Numeric keypad 3 key
    VK_NUMPAD4 = 0x64                       #Numeric keypad 4 key
    VK_NUMPAD5 = 0x65                       #Numeric keypad 5 key
    VK_NUMPAD6 = 0x66                       #Numeric keypad 6 key
    VK_NUMPAD7 = 0x67                       #Numeric keypad 7 key
    VK_NUMPAD8 = 0x68                       #Numeric keypad 8 key
    VK_NUMPAD9 = 0x69                       #Numeric keypad 9 key
    VK_MULTIPLY = 0x6A                      #Multiply key
    VK_ADD = 0x6B                           #Add key
    VK_SEPARATOR = 0x6C                     #Separator key
    VK_SUBTRACT = 0x6D                      #Subtract key
    VK_DECIMAL = 0x6E                       #Decimal key
    VK_DIVIDE = 0x6F                        #Divide key
    VK_F1 = 0x70                            #F1 key
    VK_F2 = 0x71                            #F2 key
    VK_F3 = 0x72                            #F3 key
    VK_F4 = 0x73                            #F4 key
    VK_F5 = 0x74                            #F5 key
    VK_F6 = 0x75                            #F6 key
    VK_F7 = 0x76                            #F7 key
    VK_F8 = 0x77                            #F8 key
    VK_F9 = 0x78                            #F9 key
    VK_F10 = 0x79                           #F10 key
    VK_F11 = 0x7A                           #F11 key
    VK_F12 = 0x7B                           #F12 key
    VK_F13 = 0x7C                           #F13 key
    VK_F14 = 0x7D                           #F14 key
    VK_F15 = 0x7E                           #F15 key
    VK_F16 = 0x7F                           #F16 key
    VK_F17 = 0x80                           #F17 key
    VK_F18 = 0x81                           #F18 key
    VK_F19 = 0x82                           #F19 key
    VK_F20 = 0x83                           #F20 key
    VK_F21 = 0x84                           #F21 key
    VK_F22 = 0x85                           #F22 key
    VK_F23 = 0x86                           #F23 key
    VK_F24 = 0x87                           #F24 key
    VK_NUMLOCK = 0x90                       #NUM LOCK key
    VK_SCROLL = 0x91                        #SCROLL LOCK key
    VK_LSHIFT = 0xA0                        #Left SHIFT key
    VK_RSHIFT = 0xA1                        #Right SHIFT key
    VK_LCONTROL = 0xA2                      #Left CONTROL key
    VK_RCONTROL = 0xA3                      #Right CONTROL key
    VK_LMENU = 0xA4                         #Left MENU key
    VK_RMENU = 0xA5                         #Right MENU key
    VK_BROWSER_BACK = 0xA6                  #Browser Back key
    VK_BROWSER_FORWARD = 0xA7               #Browser Forward key
    VK_BROWSER_REFRESH = 0xA8               #Browser Refresh key
    VK_BROWSER_STOP = 0xA9                  #Browser Stop key
    VK_BROWSER_SEARCH = 0xAA                #Browser Search key
    VK_BROWSER_FAVORITES = 0xAB             #Browser Favorites key
    VK_BROWSER_HOME = 0xAC                  #Browser Start and Home key
    VK_VOLUME_MUTE = 0xAD                   #Volume Mute key
    VK_VOLUME_DOWN = 0xAE                   #Volume Down key
    VK_VOLUME_UP = 0xAF                     #Volume Up key
    VK_MEDIA_NEXT_TRACK = 0xB0              #Next Track key
    VK_MEDIA_PREV_TRACK = 0xB1              #Previous Track key
    VK_MEDIA_STOP = 0xB2                    #Stop Media key
    VK_MEDIA_PLAY_PAUSE = 0xB3              #Play/Pause Media key
    VK_LAUNCH_MAIL = 0xB4                   #Start Mail key
    VK_LAUNCH_MEDIA_SELECT = 0xB5           #Select Media key
    VK_LAUNCH_APP1 = 0xB6                   #Start Application 1 key
    VK_LAUNCH_APP2 = 0xB7                   #Start Application 2 key
    VK_OEM_1 = 0xBA                         #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the ';:' key
    VK_OEM_PLUS = 0xBB                      #For any country/region, the '+' key
    VK_OEM_COMMA = 0xBC                     #For any country/region, the ',' key
    VK_OEM_MINUS = 0xBD                     #For any country/region, the '-' key
    VK_OEM_PERIOD = 0xBE                    #For any country/region, the '.' key
    VK_OEM_2 = 0xBF                         #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the '/?' key
    VK_OEM_3 = 0xC0                         #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the '`~' key
    VK_OEM_4 = 0xDB                         #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the '[{' key
    VK_OEM_5 = 0xDC                         #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the '\|' key
    VK_OEM_6 = 0xDD                         #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the ']}' key
    VK_OEM_7 = 0xDE                         #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the 'single-quote/double-quote' key
    VK_OEM_8 = 0xDF                         #Used for miscellaneous characters; it can vary by keyboard.
    VK_OEM_102 = 0xE2                       #Either the angle bracket key or the backslash key on the RT 102-key keyboard
    VK_PROCESSKEY = 0xE5                    #IME PROCESS key
    VK_PACKET = 0xE7                        #Used to pass Unicode characters as if they were keystrokes. The VK_PACKET key is the low word of a 32-bit Virtual Key value used for non-keyboard input methods. For more information, see Remark in KEYBDINPUT, SendInput, WM_KEYDOWN, and WM_KeyUp
    VK_ATTN = 0xF6                          #Attn key
    VK_CRSEL = 0xF7                         #CrSel key
    VK_EXSEL = 0xF8                         #ExSel key
    VK_EREOF = 0xF9                         #Erase EOF key
    VK_PLAY = 0xFA                          #Play key
    VK_ZOOM = 0xFB                          #Zoom key
    VK_NONAME = 0xFC                        #Reserved
    VK_PA1 = 0xFD                           #PA1 key
    VK_OEM_CLEAR = 0xFE                     #Clear key

class ConsoleColor():
    '''This class defines the values of color for printing on console window'''
    Black = 0
    DarkBlue = 1
    DarkGreen = 2
    DarkCyan = 3
    DarkRed = 4
    DarkMagenta = 5
    DarkYellow = 6
    Gray = 7
    DarkGray = 8
    Blue = 9
    Green = 10
    Cyan = 11
    Red = 12
    Magenta = 13
    Yellow = 14
    White = 15

class ExpandCollapseState():
    ExpandCollapseState_Collapsed = 0
    ExpandCollapseState_Expanded = 1
    ExpandCollapseState_PartiallyExpanded = 2
    ExpandCollapseState_LeafNode = 3

class ToggleState():
    ToggleState_Off = 0
    ToggleState_On = 1
    ToggleState_Indeterminate = 2

class WindowVisualState():
    WindowVisualState_Normal = 0
    WindowVisualState_Maximized = 1
    WindowVisualState_Minimized = 2

class ShowWindow():
    Hide = 0
    ShowNormal = 1
    Normal = 1
    ShowMinimized = 2
    ShowMaximized = 3
    Maximize = 3
    ShowNoActivate = 4
    Show = 5
    Minimize = 6
    ShowMinNoActive = 7
    ShowNa = 8
    Restore = 9
    ShowDefault = 10
    ForceMinimize = 11
    Max = 11

class SWP():
    HWND_TOP = 0
    HWND_BOTTOM = 1
    HWND_TOPMOST = -1
    HWND_NOTOPMOST = -2
    SWP_NOSIZE = 0x0001
    SWP_NOMOVE = 0x0002
    SWP_NOZORDER = 0x0004
    SWP_NOREDRAW = 0x0008
    SWP_NOACTIVATE = 0x0010
    SWP_FRAMECHANGED = 0x0020  # The frame changed: send WM_NCCALCSIZE
    SWP_SHOWWINDOW = 0x0040
    SWP_HIDEWINDOW = 0x0080
    SWP_NOCOPYBITS = 0x0100
    SWP_NOOWNERZORDER = 0x0200  # Don't do owner Z ordering
    SWP_NOSENDCHANGING = 0x0400  # Don't send WM_WINDOWPOSCHANGING
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_DEFERERASE = 0x2000
    SWP_ASYNCWINDOWPOS = 0x4000

class MB():
    OK = 0x00000000
    OKCANCEL = 0x00000001
    ABORTRETRYIGNORE = 0x00000002
    YESNOCANCEL = 0x00000003
    YESNO = 0x00000004
    RETRYCANCEL = 0x00000005
    CANCELTRYCONTINUE = 0x00000006
    ICONHAND = 0x00000010
    ICONQUESTION = 0x00000020
    ICONEXCLAMATION = 0x00000030
    ICONASTERISK = 0x00000040
    USERICON = 0x00000080
    ICONWARNING = 0x00000030
    ICONERROR = 0x00000010
    ICONINFORMATION = 0x00000040
    ICONSTOP = 0x00000010
    DEFBUTTON1 = 0x00000000
    DEFBUTTON2 = 0x00000100
    DEFBUTTON3 = 0x00000200
    DEFBUTTON4 = 0x00000300
    APPLMODAL = 0x00000000
    SYSTEMMODAL = 0x00001000
    TASKMODAL = 0x00002000
    HELP = 0x00004000 # Help Button
    NOFOCUS = 0x00008000
    SETFOREGROUND = 0x00010000
    DEFAULT_DESKTOP_ONLY = 0x00020000
    TOPMOST = 0x00040000
    RIGHT = 0x00080000
    RTLREADING = 0x00100000
    SERVICE_NOTIFICATION = 0x00200000
    SERVICE_NOTIFICATION = 0x00040000
    SERVICE_NOTIFICATION_NT3X = 0x00040000

    TYPEMASK = 0x0000000F
    ICONMASK = 0x000000F0
    DEFMASK = 0x00000F00
    MODEMASK = 0x00003000
    MISCMASK = 0x0000C000

    IDOK = 1
    IDCANCEL = 2
    IDABORT = 3
    IDRETRY = 4
    IDIGNORE = 5
    IDYES = 6
    IDNO = 7
    IDCLOSE = 8
    IDHELP = 9
    IDTRYAGAIN = 10
    IDCONTINUE = 11
    IDTIMEOUT = 32000

class Win32API():
    '''Some native methods for python calling'''
    StdOutputHandle = -11
    ConsoleOutputHandle = None
    DefaultColor = None
    KeyDict = None
    CharacterDict = None

    @staticmethod
    def GetClipboardText():
        ctypes.windll.user32.OpenClipboard(0)
        text = ctypes.windll.user32.GetClipboardData(13) # CF_TEXT = 1 CF_UNICODETEXT=13
        ctypes.windll.user32.CloseClipboard()
        ctext = ctypes.c_wchar_p(text)
        return ctext.value
        '''size = ctypes.windll.kernel32.MultiByteToWideChar(936, 0, text, -1, 0, 0)
        pwChar = ClientObject.dll.NewWCharArray(size)
        ctypes.windll.kernel32.MultiByteToWideChar(936, 0, text, -1, pwChar, size)
        wstr = WString()
        wstr.wcharArray = pwChar
        wstr.cwcharp = ctypes.c_wchar_p(pwChar)
        return wstr'''

    @staticmethod
    def SetConsoleColor(color):
        '''Change the text color on console window'''
        if not Win32API.DefaultColor:
            if not Win32API.ConsoleOutputHandle:
                Win32API.ConsoleOutputHandle = ctypes.windll.kernel32.GetStdHandle(Win32API.StdOutputHandle)
            bufferInfo = ConsoleScreenBufferInfo()
            ctypes.windll.kernel32.GetConsoleScreenBufferInfo(Win32API.ConsoleOutputHandle, ctypes.byref(bufferInfo))
            Win32API.DefaultColor = int(bufferInfo.wAttributes & 0xFF)
        if IsPy3:
            sys.stdout.flush() # need flush stdout in python 3
        ctypes.windll.kernel32.SetConsoleTextAttribute(Win32API.ConsoleOutputHandle, color)

    @staticmethod
    def ResetConsoleColor():
        '''Reset the default text color on console window'''
        if IsPy3:
            sys.stdout.flush() # need flush stdout in python 3
        ctypes.windll.kernel32.SetConsoleTextAttribute(Win32API.ConsoleOutputHandle, Win32API.DefaultColor)

    @staticmethod
    def WindowFromPoint(x, y):
        '''Return hwnd'''
        point = Point()
        point.x = x
        point.y = y
        return ctypes.windll.user32.WindowFromPoint(point)

    @staticmethod
    def GetCursorPos():
        '''Return tuple (x, y)'''
        point = Point()
        ctypes.windll.user32.GetCursorPos(ctypes.byref(point))
        return int(point.x), int(point.y)

    @staticmethod
    def SetCursorPos(x, y):
        '''Set cursor to point x, y'''
        ctypes.windll.user32.SetCursorPos(x, y)

    @staticmethod
    def GetDoubleClickTime():
        '''Get the double click time of mouse'''
        return ctypes.windll.user32.GetDoubleClickTime()

    @staticmethod
    def mouse_event(dwFlags, dx, dy, dwData, dwExtraInfo):
        '''Call API mouse_event from user32.dll'''
        ctypes.windll.user32.mouse_event(dwFlags, dx, dy, dwData, dwExtraInfo)

    @staticmethod
    def keybd_event(bVk, bScan, dwFlags, dwExtraInfo):
        '''Call API keybd_event from user32.dll'''
        ctypes.windll.user32.keybd_event(bVk, bScan, dwFlags, dwExtraInfo)

    @staticmethod
    def PostMessage(handle, msg, wparam, lparam):
        '''Call API PostMessageW from user32.dll'''
        return ctypes.windll.user32.PostMessageW(handle, msg, wparam, lparam)

    @staticmethod
    def SendMessage(handle, msg, wparam, lparam):
        '''Call API SendMessageW from user32.dll'''
        return ctypes.windll.user32.SendMessageW(handle, msg, wparam, lparam)

    @staticmethod
    def MouseClick(x, y):
        '''
        Simulate mouse click at point x, y
        x and y must be integer
        '''
        Win32API.SetCursorPos(x, y)
        Win32API.mouse_event(MouseEventFlags.LeftDown | MouseEventFlags.Absolute, x, y, 0, 0)
        time.sleep(0.1)
        Win32API.mouse_event(MouseEventFlags.LeftUp | MouseEventFlags.Absolute, x, y, 0, 0)

    @staticmethod
    def MouseMiddleClick(x, y):
        '''
        Simulate mouse middle click at point x, y
        x and y must be integer
        '''
        Win32API.SetCursorPos(x, y)
        Win32API.mouse_event(MouseEventFlags.MiddleDown | MouseEventFlags.Absolute, x, y, 0, 0)
        time.sleep(0.1)
        Win32API.mouse_event(MouseEventFlags.MiddleUp | MouseEventFlags.Absolute, x, y, 0, 0)

    @staticmethod
    def MouseRightClick(x, y):
        '''
        Simulate mouse right click at point x, y
        x and y must be integer
        '''
        Win32API.SetCursorPos(x, y)
        Win32API.mouse_event(MouseEventFlags.RightDown | MouseEventFlags.Absolute, x, y, 0, 0)
        time.sleep(0.1)
        Win32API.mouse_event(MouseEventFlags.RightUp | MouseEventFlags.Absolute, x, y, 0, 0)

    @staticmethod
    def MouseMoveTo(x, y, moveSpeed = 1):
        '''
        Simulate mouse move to point x, y from current cursor
        x and y must be integer
        moveTime: double, the max simulating time
        '''
        if moveSpeed <= 0:
            moveTime = 0
        else:
            moveTime = MAX_MOVE_SECOND / moveSpeed
        curX, curY = Win32API.GetCursorPos()
        xCount = abs(x - curX)
        yCount = abs(y - curY)
        maxPoint = max(xCount, yCount)
        screenWidth, screenHeight = Win32API.GetScreenSize()
        maxSide = max(screenWidth, screenHeight)
        minSide = min(screenWidth, screenHeight)
        if maxPoint > minSide:
            maxPoint = minSide
        if maxPoint < maxSide:
            maxPoint = 100 + int((maxSide-100) / maxSide * maxPoint)
            moveTime = moveTime * maxPoint * 1.0 / maxSide
        stepCount = maxPoint // 20
        print
        if stepCount > 1:
            xStep = (x - curX) * 1.0 / stepCount
            yStep = (y - curY) * 1.0 / stepCount
            interval = moveTime / stepCount
            for i in range(stepCount):
                cx = curX + int(xStep * i)
                cy = curY + int(yStep * i)
                # upper-left(0,0), lower-right(65536,65536)
                # Win32API.mouse_event(MouseEventFlags.Move | MouseEventFlags.Absolute, cx*65536//screenWidth, cy*65536//screenHeight, 0, 0)
                Win32API.SetCursorPos(cx, cy)
                time.sleep(interval)
        Win32API.SetCursorPos(x, y)

    @staticmethod
    def MouseDragTo(x1, y1, x2, y2, moveSpeed = 1):
        '''
        Simulate mouse drag from point x1, y1 to point x2, y2
        x1, y1, x2, y2, step must be integer
        dragTime: double, the drag will be done in dragTime
        '''
        if moveSpeed <= 0:
            moveTime = 0
        else:
            moveTime = MAX_MOVE_SECOND / moveSpeed
        xCount = abs(x2 - x1)
        yCount = abs(y2 - y1)
        maxPoint = max(xCount, yCount)
        screenWidth, screenHeight = Win32API.GetScreenSize()
        maxSide = max(screenWidth, screenHeight)
        minSide = min(screenWidth, screenHeight)
        if maxPoint < maxSide:
            maxPoint = 100 + int((maxSide-100) / maxSide * maxPoint)
            dragTime = dragTime * maxPoint * 1.0 / maxSide
        stepCount = maxPoint // 20
        Win32API.SetCursorPos(x1, y1)
        Win32API.mouse_event(MouseEventFlags.LeftDown| MouseEventFlags.Absolute, x1*65536//screenWidth, y1*65536//screenHeight, 0, 0)
        if stepCount > 1:
            xStep = (x2 - x1) * 1.0 / stepCount
            yStep = (y2 - y1) * 1.0 / stepCount
            interval = dragTime / stepCount
            for i in range(stepCount):
                x1 += xStep
                y1 += yStep
                Win32API.mouse_event(MouseEventFlags.Move, int(xStep), int(yStep), 0, 0)
                Win32API.SetCursorPos(int(x1), int(y1))
                time.sleep(interval)
        Win32API.mouse_event(MouseEventFlags.Absolute | MouseEventFlags.LeftUp, x2*65536//screenWidth, y2*65536//screenHeight, 0, 0)

    @staticmethod
    def GetScreenSize():
        '''Return tuple (width, height)'''
        SM_CXSCREEN = 0
        SM_CYSCREEN = 1
        w = ctypes.windll.user32.GetSystemMetrics(SM_CXSCREEN)
        h = ctypes.windll.user32.GetSystemMetrics(SM_CYSCREEN)
        return w, h

    @staticmethod
    def GetPixel(x, y, handle = 0):
        '''
        If handle is 0, get pixel from desktop, return r, g, b, int, hex
        Not all devices support GetPixel. An application should call GetDeviceCaps to determine whether a specified device supports this function. Console window doesn't support.
        '''
        hdc = ctypes.windll.user32.GetWindowDC(handle)
        pixel = ctypes.windll.gdi32.GetPixel(hdc, x, y)
        ctypes.windll.user32.ReleaseDC(handle, hdc)
        # if pixel == 0xFFFFFFFF: # CLR_INVALID
            # return None
        r = pixel & 0x0000FF
        g = (pixel & 0x00FF00) >> 8
        b = (pixel & 0xFF0000) >> 16
        return r, g, b, pixel, '0x{0:02X}{1:02X}{2:02X}'.format(b,g,r), '#{0:02X}{1:02X}{2:02X}'.format(r,g,b)

    @staticmethod
    def MessageBox(content, title, flags = MB.OK):
        '''Call API MessageBox from user32.dll'''
        c_content = ctypes.c_wchar_p(content)
        c_title = ctypes.c_wchar_p(title)
        return ctypes.windll.user32.MessageBoxW(0, c_content, c_title, flags)

    @staticmethod
    def SetForegroundWindow(hWnd):
        '''
        Set a window to foreground
        hWnd: integer, handle of a Win32 window
        '''
        return ctypes.windll.user32.SetForegroundWindow(hWnd)

    @staticmethod
    def SetWindowTopmost(hWnd, isTopmost):
        '''
        Set a window to Topmost
        hWnd: integer, handle of a Win32 window
        isTopmost: bool, be topmost or not
        '''
        topValue = SWP.HWND_TOPMOST if isTopmost else SWP.HWND_NOTOPMOST
        return Win32API.SetWindowPos(hWnd, topValue, 0, 0, 0, 0, SWP.SWP_NOSIZE|SWP.SWP_NOMOVE)

    @staticmethod
    def ShowWindow(hWnd, cmdShow):
        '''ShowWindow(hWnd, ShowWindow.Show), see values in class ShowWindow'''
        return ctypes.windll.user32.ShowWindow(hWnd, cmdShow)

    @staticmethod
    def MoveWindow(hWnd, x, y, width, height, repaint = 0):
        '''Call API MoveWindow from user32.dll'''
        return ctypes.windll.user32.MoveWindow(hWnd, x, y, width, height, repaint)

    @staticmethod
    def SetWindowPos(hWnd, hWndInsertAfter, x, y, width, height, flags):
        '''Call API SetWindowPos from user32.dll, flags see class SWP'''
        return ctypes.windll.user32.SetWindowPos(hWnd, hWndInsertAfter, x, y, width, height, flags)

    @staticmethod
    def GetWindowText(hWnd):
        '''Get window text'''
        MAX_PATH = 260
        wcharArray = ClientObject.dll.NewWCharArray(MAX_PATH)
        if wcharArray:
            ctypes.windll.user32.GetWindowTextW(hWnd, wcharArray, MAX_PATH)
            c_text = ctypes.c_wchar_p(wcharArray)
            text = c_text.value[:]
            ClientObject.dll.DeleteWCharArray(wcharArray)
            return text

    @staticmethod
    def SetWindowText(hWnd, text):
        '''Set window text'''
        c_text = ctypes.c_wchar_p(text)
        return ctypes.windll.user32.SetWindowTextW(hWnd, c_text)

    @staticmethod
    def GetForegroundWindow():
        return ctypes.windll.user32.GetForegroundWindow()

    @staticmethod
    def GetProcessCommandLine(processId):
        wstr = ClientObject.dll.GetProcessCommandLine(processId)
        if wstr:
            cmdLine = ctypes.c_wchar_p(wstr)
            cmdLine = cmdLine.value[:]
            ClientObject.dll.DeleteWCharArray(wstr)
            return cmdLine
        else:
            return ''

    @staticmethod
    def GetParentProcessId(processId = -1):
        return ClientObject.dll.GetParentProcessId(processId)

    @staticmethod
    def EnumProcess():
        '''Return a namedtuple's iter (for p in this p.pid p.name)'''
        import collections
        hSnapshot = ctypes.windll.kernel32.CreateToolhelp32Snapshot(15, 0)
        processEntry32 = tagPROCESSENTRY32()
        processClass = collections.namedtuple('processInfo', 'pid name')
        processEntry32.dwSize = ctypes.sizeof(processEntry32)
        continueFind = ctypes.windll.kernel32.Process32FirstW(hSnapshot, ctypes.byref(processEntry32))
        while continueFind:
            pid = processEntry32.th32ProcessID
            name = (processEntry32.szExeFile)
            yield processClass(pid, name)
            continueFind = ctypes.windll.kernel32.Process32NextW(hSnapshot, ctypes.byref(processEntry32))

    @staticmethod
    def SendKey(key):
        '''
        Simulate typing a key
        key: a value in class Keys
        '''
        Win32API.keybd_event(key, 0, KeyboardEventFlags.ExtendedKey, 0)
        Win32API.keybd_event(key, 0, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey, 0)

    #@staticmethod
    #def SendWait(keys):
        #'''this method needs .Net and PythonForNet, will call System.Windows.Forms.SendKeys.SendWait'''
        #import clr
        #import System.Windows.Forms
        #System.Windows.Forms.SendKeys.SendWait(keys)

    @staticmethod
    def PressKey(key):
        '''
        Simulate a key down for key
        key: a value in class Keys
        '''
        Win32API.keybd_event(Keys.VK_LWIN, 0, KeyboardEventFlags.ExtendedKey, 0)

    @staticmethod
    def ReleaseKey(key):
        '''
        Simulate a key up for key
        key: a value in class Keys
        '''
        Win32API.keybd_event(key, 0, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey, 0)

    @staticmethod
    def VKtoSC(key):
        '''key: a value in class Keys'''
        SCDict = {
            Keys.VK_LSHIFT : 0x02A,
            Keys.VK_RSHIFT : 0x136,
            Keys.VK_LCONTROL : 0x01D,
            Keys.VK_RCONTROL : 0x11D,
            Keys.VK_LMENU : 0x038,
            Keys.VK_RMENU : 0x138,
            Keys.VK_LWIN : 0x15B,
            Keys.VK_RWIN : 0x15C,
            Keys.VK_NUMPAD0 : 0x52,
            Keys.VK_NUMPAD1 : 0x4F,
            Keys.VK_NUMPAD2 : 0x50,
            Keys.VK_NUMPAD3 : 0x51,
            Keys.VK_NUMPAD4 : 0x4B,
            Keys.VK_NUMPAD5 : 0x4C,
            Keys.VK_NUMPAD6 : 0x4D,
            Keys.VK_NUMPAD7 : 0x47,
            Keys.VK_NUMPAD8 : 0x48,
            Keys.VK_NUMPAD9 : 0x49,
            Keys.VK_DECIMAL : 0x53,
            Keys.VK_NUMLOCK : 0x145,
            Keys.VK_DIVIDE :  0x135,
            Keys.VK_MULTIPLY : 0x037,
            Keys.VK_SUBTRACT : 0x04A,
            Keys.VK_ADD : 0x04E,
        }
        if key in SCDict:
            return SCDict[key]
        scanCode = ctypes.windll.user32.MapVirtualKeyA(key, 0)
        if not scanCode:
            return 0
        keyList = [Keys.VK_APPS, Keys.VK_CANCEL, Keys.VK_SNAPSHOT, Keys.VK_DIVIDE, Keys.VK_NUMLOCK]
        if key in keyList:
            scanCode |= 0x0100
        return scanCode

    @staticmethod
    def SendKeys(text, interval = 0.01, debug = False):
        '''
        Simulate typing keys on keyboard
        text: unicode, keys to type
        interval: double, seconds between keys
        debug: bool, if True, print the Keys
        example:
        Win32API.SendKeys('{Ctrl}{End}{SPACE}{CTRL}{1 3}  {CTRL}(12)123ab{c 3}A(BC){(}你好{)}{ENTER}')
        Win32API.SendKeys('0123456789{ENTER}')
        Win32API.SendKeys('ABCDEFGHIJKLMNOPQRSTUVWXYZ{ENTER}')
        Win32API.SendKeys('abcdefghijklmnopqrstuvwxyz{ENTER}')
        Win32API.SendKeys('`~!@#$%^&*{(}{)}-_=+{ENTER}')
        Win32API.SendKeys('[]{{}{}}\\|;:\'\",<.>/?{ENTER}')
        '''
        holdKeys = ['WIN', 'LWIN', 'RWIN', 'SHIFT', 'LSHIFT', 'RSHIFT', 'CTRL', 'LCTRL', 'RCTRL', 'ALT', 'LALT', 'RALT']
        if not Win32API.KeyDict:
            Win32API.KeyDict = {
                'LBUTTON' : 0x01,                       #Left mouse button
                'RBUTTON' : 0x02,                       #Right mouse button
                'CANCEL' : 0x03,                        #Control-break processing
                'MBUTTON' : 0x04,                       #Middle mouse button (three-button mouse)
                'XBUTTON1' : 0x05,                      #X1 mouse button
                'XBUTTON2' : 0x06,                      #X2 mouse button
                'BACK' : 0x08,                          #BACKSPACE key
                'TAB' : 0x09,                           #TAB key
                'CLEAR' : 0x0C,                         #CLEAR key
                'RETURN' : 0x0D,                        #ENTER key
                'ENTER' : 0x0D,                         #ENTER key
                'SHIFT' : 0x10,                         #SHIFT key
                'CTRL' : 0x11,                          #CTRL key
                'CONTROL' : 0x11,                       #CTRL key
                'ALT' : 0x12,                           #ALT key
                'PAUSE' : 0x13,                         #PAUSE key
                'CAPITAL' : 0x14,                       #CAPS LOCK key
                'KANA' : 0x15,                          #IME Kana mode
                'HANGUEL' : 0x15,                       #IME Hanguel mode (maintained for compatibility; use VK_HANGUL)
                'HANGUL' : 0x15,                        #IME Hangul mode
                'JUNJA' : 0x17,                         #IME Junja mode
                'FINAL' : 0x18,                         #IME final mode
                'HANJA' : 0x19,                         #IME Hanja mode
                'KANJI' : 0x19,                         #IME Kanji mode
                'ESC' : 0x1B,                           #ESC key
                'ESCAPE' : 0x1B,                        #ESC key
                'CONVERT' : 0x1C,                       #IME convert
                'NONCONVERT' : 0x1D,                    #IME nonconvert
                'ACCEPT' : 0x1E,                        #IME accept
                'MODECHANGE' : 0x1F,                    #IME mode change request
                ' ' : 0x20,                             #SPACEBAR
                'SPACE' : 0x20,                         #SPACEBAR
                'PRIOR' : 0x21,                         #PAGE UP key
                'NEXT' : 0x22,                          #PAGE DOWN key
                'END' : 0x23,                           #END key
                'HOME' : 0x24,                          #HOME key
                'LEFT' : 0x25,                          #LEFT ARROW key
                'UP' : 0x26,                            #UP ARROW key
                'RIGHT' : 0x27,                         #RIGHT ARROW key
                'DOWN' : 0x28,                          #DOWN ARROW key
                'SELECT' : 0x29,                        #SELECT key
                'PRINT' : 0x2A,                         #PRINT key
                'EXECUTE' : 0x2B,                       #EXECUTE key
                'SNAPSHOT' : 0x2C,                      #PRINT SCREEN key
                'INSERT' : 0x2D,                        #INS key
                'DELETE' : 0x2E,                        #DEL key
                'HELP' : 0x2F,                          #HELP key
                'WIN' : 0x5B,                           #Left Windows key (Natural keyboard)
                'LWIN' : 0x5B,                          #Left Windows key (Natural keyboard)
                'RWIN' : 0x5C,                          #Right Windows key (Natural keyboard)
                'APPS' : 0x5D,                          #Applications key (Natural keyboard)
                'SLEEP' : 0x5F,                         #Computer Sleep key
                'NUMPAD0' : 0x60,                       #Numeric keypad 0 key
                'NUMPAD1' : 0x61,                       #Numeric keypad 1 key
                'NUMPAD2' : 0x62,                       #Numeric keypad 2 key
                'NUMPAD3' : 0x63,                       #Numeric keypad 3 key
                'NUMPAD4' : 0x64,                       #Numeric keypad 4 key
                'NUMPAD5' : 0x65,                       #Numeric keypad 5 key
                'NUMPAD6' : 0x66,                       #Numeric keypad 6 key
                'NUMPAD7' : 0x67,                       #Numeric keypad 7 key
                'NUMPAD8' : 0x68,                       #Numeric keypad 8 key
                'NUMPAD9' : 0x69,                       #Numeric keypad 9 key
                'MULTIPLY' : 0x6A,                      #Multiply key
                'ADD' : 0x6B,                           #Add key
                'SEPARATOR' : 0x6C,                     #Separator key
                'SUBTRACT' : 0x6D,                      #Subtract key
                'DECIMAL' : 0x6E,                       #Decimal key
                'DIVIDE' : 0x6F,                        #Divide key
                'F1' : 0x70,                            #F1 key
                'F2' : 0x71,                            #F2 key
                'F3' : 0x72,                            #F3 key
                'F4' : 0x73,                            #F4 key
                'F5' : 0x74,                            #F5 key
                'F6' : 0x75,                            #F6 key
                'F7' : 0x76,                            #F7 key
                'F8' : 0x77,                            #F8 key
                'F9' : 0x78,                            #F9 key
                'F10' : 0x79,                           #F10 key
                'F11' : 0x7A,                           #F11 key
                'F12' : 0x7B,                           #F12 key
                'F13' : 0x7C,                           #F13 key
                'F14' : 0x7D,                           #F14 key
                'F15' : 0x7E,                           #F15 key
                'F16' : 0x7F,                           #F16 key
                'F17' : 0x80,                           #F17 key
                'F18' : 0x81,                           #F18 key
                'F19' : 0x82,                           #F19 key
                'F20' : 0x83,                           #F20 key
                'F21' : 0x84,                           #F21 key
                'F22' : 0x85,                           #F22 key
                'F23' : 0x86,                           #F23 key
                'F24' : 0x87,                           #F24 key
                'NUMLOCK' : 0x90,                       #NUM LOCK key
                'SCROLL' : 0x91,                        #SCROLL LOCK key
                'LSHIFT' : 0xA0,                        #Left SHIFT key
                'RSHIFT' : 0xA1,                        #Right SHIFT key
                'LCONTROL' : 0xA2,                      #Left CONTROL key
                'RCONTROL' : 0xA3,                      #Right CONTROL key
                'LALT' : 0xA4,                         #Left MENU key
                'RALT' : 0xA5,                         #Right MENU key
                'BROWSER_BACK' : 0xA6,                  #Browser Back key
                'BROWSER_FORWARD' : 0xA7,               #Browser Forward key
                'BROWSER_REFRESH' : 0xA8,               #Browser Refresh key
                'BROWSER_STOP' : 0xA9,                  #Browser Stop key
                'BROWSER_SEARCH' : 0xAA,                #Browser Search key
                'BROWSER_FAVORITES' : 0xAB,             #Browser Favorites key
                'BROWSER_HOME' : 0xAC,                  #Browser Start and Home key
                'VOLUME_MUTE' : 0xAD,                   #Volume Mute key
                'VOLUME_DOWN' : 0xAE,                   #Volume Down key
                'VOLUME_UP' : 0xAF,                     #Volume Up key
                'MEDIA_NEXT_TRACK' : 0xB0,              #Next Track key
                'MEDIA_PREV_TRACK' : 0xB1,              #Previous Track key
                'MEDIA_STOP' : 0xB2,                    #Stop Media key
                'MEDIA_PLAY_PAUSE' : 0xB3,              #Play/Pause Media key
                'LAUNCH_MAIL' : 0xB4,                   #Start Mail key
                'LAUNCH_MEDIA_SELECT' : 0xB5,           #Select Media key
                'LAUNCH_APP1' : 0xB6,                   #Start Application 1 key
                'LAUNCH_APP2' : 0xB7,                   #Start Application 2 key
                'OEM_1' : 0xBA,                         #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the ';:' key
                'OEM_PLUS' : 0xBB,                      #For any country/region, the '+' key
                'OEM_COMMA' : 0xBC,                     #For any country/region, the ',' key
                'OEM_MINUS' : 0xBD,                     #For any country/region, the '-' key
                'OEM_PERIOD' : 0xBE,                    #For any country/region, the '.' key
                'OEM_2' : 0xBF,                         #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the '/?' key
                'OEM_3' : 0xC0,                         #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the '`~' key
                'OEM_4' : 0xDB,                         #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the '[{' key
                'OEM_5' : 0xDC,                         #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the '\|' key
                'OEM_6' : 0xDD,                         #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the ']}' key
                'OEM_7' : 0xDE,                         #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the 'single-quote/double-quote' key
                'OEM_8' : 0xDF,                         #Used for miscellaneous characters; it can vary by keyboard.
                'OEM_102' : 0xE2,                       #Either the angle bracket key or the backslash key on the RT 102-key keyboard
                'PROCESSKEY' : 0xE5,                    #IME PROCESS key
                'PACKET' : 0xE7,                        #Used to pass Unicode characters as if they were keystrokes. The VK_PACKET key is the low word of a 32-bit Virtual Key value used for non-keyboard input methods. For more information, see Remark in KEYBDINPUT, SendInput, WM_KEYDOWN, and WM_KeyUp
                'ATTN' : 0xF6,                          #Attn key
                'CRSEL' : 0xF7,                         #CrSel key
                'EXSEL' : 0xF8,                         #ExSel key
                'EREOF' : 0xF9,                         #Erase EOF key
                'PLAY' : 0xFA,                          #Play key
                'ZOOM' : 0xFB,                          #Zoom key
                'NONAME' : 0xFC,                        #Reserved
                'PA1' : 0xFD,                           #PA1 key
                'OEM_CLEAR' : 0xFE,                     #Clear key
            }
        if not Win32API.CharacterDict:
            Win32API.CharacterDict = {
                '0' : 0x30,                             #0 key
                '1' : 0x31,                             #1 key
                '2' : 0x32,                             #2 key
                '3' : 0x33,                             #3 key
                '4' : 0x34,                             #4 key
                '5' : 0x35,                             #5 key
                '6' : 0x36,                             #6 key
                '7' : 0x37,                             #7 key
                '8' : 0x38,                             #8 key
                '9' : 0x39,                             #9 key
                'a' : 0x41,                             #A key
                'A' : 0x41,                             #A key
                'b' : 0x42,                             #B key
                'B' : 0x42,                             #B key
                'c' : 0x43,                             #C key
                'C' : 0x43,                             #C key
                'd' : 0x44,                             #D key
                'D' : 0x44,                             #D key
                'e' : 0x45,                             #E key
                'E' : 0x45,                             #E key
                'f' : 0x46,                             #F key
                'F' : 0x46,                             #F key
                'g' : 0x47,                             #G key
                'G' : 0x47,                             #G key
                'h' : 0x48,                             #H key
                'H' : 0x48,                             #H key
                'i' : 0x49,                             #I key
                'I' : 0x49,                             #I key
                'j' : 0x4A,                             #J key
                'J' : 0x4A,                             #J key
                'k' : 0x4B,                             #K key
                'K' : 0x4B,                             #K key
                'l' : 0x4C,                             #L key
                'L' : 0x4C,                             #L key
                'm' : 0x4D,                             #M key
                'M' : 0x4D,                             #M key
                'n' : 0x4E,                             #N key
                'N' : 0x4E,                             #N key
                'o' : 0x4F,                             #O key
                'O' : 0x4F,                             #O key
                'p' : 0x50,                             #P key
                'P' : 0x50,                             #P key
                'q' : 0x51,                             #Q key
                'Q' : 0x51,                             #Q key
                'r' : 0x52,                             #R key
                'R' : 0x52,                             #R key
                's' : 0x53,                             #S key
                'S' : 0x53,                             #S key
                't' : 0x54,                             #T key
                'T' : 0x54,                             #T key
                'u' : 0x55,                             #U key
                'U' : 0x55,                             #U key
                'v' : 0x56,                             #V key
                'V' : 0x56,                             #V key
                'w' : 0x57,                             #W key
                'W' : 0x57,                             #W key
                'x' : 0x58,                             #X key
                'X' : 0x58,                             #X key
                'y' : 0x59,                             #Y key
                'Y' : 0x59,                             #Y key
                'z' : 0x5A,                             #Z key
                'Z' : 0x5A,                             #Z key
            }
        keys = []
        printKeys = []
        i = 0
        insertIndex = 0
        length = len(text)
        hold = False
        include = False
        lastKey = ''
        while True:
            if text[i] == '{':
                rindex = text.find('}', i)
                if rindex == i+1:#{}}
                    rindex = text.find('}', i+2)
                key = text[i+1:rindex]
                key = key.split()
                upperKey = key[0].upper()
                count = 1
                if len(key) > 1:
                    count = int(key[1])
                for j in range(count):
                    if hold:
                        if upperKey in Win32API.KeyDict:
                            keyValue = Win32API.KeyDict[upperKey]
                            if lastKey == key[0]:
                                insertIndex += 1
                            printKeys.insert(insertIndex, (key[0], 'KeyDown | ExtendedKey'))
                            printKeys.insert(insertIndex+1, (key[0], 'KeyUp | ExtendedKey'))
                            keys.insert(insertIndex, (keyValue, KeyboardEventFlags.ExtendedKey))
                            keys.insert(insertIndex+1, (keyValue, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey))
                        elif key[0] in Win32API.CharacterDict:
                            keyValue = Win32API.CharacterDict[key[0]]
                            if lastKey == key[0]:
                                insertIndex += 1
                            printKeys.insert(insertIndex, (key[0], 'KeyDown | ExtendedKey'))
                            printKeys.insert(insertIndex+1, (key[0], 'KeyUp | ExtendedKey'))
                            keys.insert(insertIndex, (keyValue, KeyboardEventFlags.ExtendedKey))
                            keys.insert(insertIndex+1, (keyValue, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey))
                        else:
                            printKeys.insert(insertIndex, (key[0], 'UnicodeChar'))
                            keys.insert(insertIndex, (key[0], 'UnicodeChar'))
                        if include:
                            insertIndex += 1
                        else:
                            if key[0] in holdKeys:
                                insertIndex += 1
                            else:
                                hold = False
                    else:
                        if upperKey in Win32API.KeyDict:
                            keyValue = Win32API.KeyDict[upperKey]
                            printKeys.append((key[0], 'KeyDown | ExtendedKey'))
                            printKeys.append((key[0], 'KeyUp | ExtendedKey'))
                            keys.append((keyValue, KeyboardEventFlags.ExtendedKey))
                            keys.append((keyValue, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey))
                            if upperKey in holdKeys:
                                hold = True
                                insertIndex = len(keys) - 1
                            else:
                                hold = False
                        else:
                            printKeys.append((key[0], 'UnicodeChar'))
                            keys.append((key[0], 'UnicodeChar'))
                    lastKey = key[0]
                i = rindex + 1
            elif text[i] == '(':
                include = True
                lastKey = text[i]
                i += 1
            elif text[i] == ')':
                include = False
                hold = False
                lastKey = text[i]
                i += 1
            else:
                if hold:
                    if text[i] in Win32API.CharacterDict:
                        keyValue = Win32API.CharacterDict[text[i]]
                        if include and lastKey == text[i]:
                            insertIndex += 1
                        printKeys.insert(insertIndex, (text[i], 'KeyDown | ExtendedKey'))
                        printKeys.insert(insertIndex + 1, (text[i], 'KeyUp | ExtendedKey'))
                        keys.insert(insertIndex, (keyValue, KeyboardEventFlags.ExtendedKey))
                        keys.insert(insertIndex + 1, (keyValue, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey))
                    else:
                        printKeys.insert(insertIndex, (text[i], 'UnicodeChar'))
                        keys.insert(insertIndex, (text[i], 'UnicodeChar'))
                    if include:
                        insertIndex += 1
                    else:
                        hold = False
                else:
                    printKeys.append((text[i], 'UnicodeChar'))
                    keys.append((text[i], 'UnicodeChar'))
                lastKey = text[i]
                i += 1
            if i >= length:
                break
        if debug:
            for key in printKeys:
                sys.stdout.write(key + '\n')
            sys.stdout.write('\n')
        for key in keys:
            if key[1] == 'UnicodeChar':
                wchar = ctypes.c_wchar_p(key[0])
                ClientObject.dll.SendUnicodeChar(wchar)
                #Win32API.PostMessage(GetFocusedControl().Handle32, 0x102, -key[0], 0)#UnicodeChar = 0x102
            else:
                scanCode = Win32API.VKtoSC(key[0])
                Win32API.keybd_event(key[0], scanCode, key[1], 0)
            time.sleep(interval)
        #make sure hold keys are not pressed
        win = ctypes.windll.user32.GetAsyncKeyState(Keys.VK_LWIN)
        ctrl = ctypes.windll.user32.GetAsyncKeyState(Keys.VK_CONTROL)
        alt = ctypes.windll.user32.GetAsyncKeyState(Keys.VK_MENU)
        shift = ctypes.windll.user32.GetAsyncKeyState(Keys.VK_SHIFT)
        if win & 0x8000:
            Logger.WriteLine('ERROR: WIN is pressed, it should not be pressed!', ConsoleColor.Red)
            Win32API.keybd_event(Keys.VK_LWIN, 0, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey, 0)
        if ctrl & 0x8000:
            Logger.WriteLine('ERROR: CTRL is pressed, it should not be pressed!', ConsoleColor.Red)
            Win32API.keybd_event(Keys.VK_CONTROL, 0, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey, 0)
        if alt & 0x8000:
            Logger.WriteLine('ERROR: ALT is pressed, it should not be pressed!', ConsoleColor.Red)
            Win32API.keybd_event(Keys.VK_MENU, 0, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey, 0)
        if shift & 0x8000:
            Logger.WriteLine('ERROR: SHIFT is pressed, it should not be pressed!', ConsoleColor.Red)
            Win32API.keybd_event(Keys.VK_SHIFT, 0, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey, 0)

class Control():
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        '''
        element: integer
        searchFromControl: Control
        searchDepth: integer
        foundIndex: integer, value must be greater or equal to 1
        searchWaitTime: float
        searchPorpertyDict: a dict that defines how to search, only the following keys are valid
                            ControlType: integer in class ControlType
                            ClassName: str or unicode
                            AutomationId: str or unicode
                            Name: str or unicode
                            SubName: str or unicode
        '''
        self._element = element
        self._name = 0
        self._className = 0
        self._automationId = 0
        self.searchFromControl = searchFromControl
        self.searchDepth = searchDepth
        self.searchWaitTime = searchWaitTime
        self.foundIndex = foundIndex
        self.searchPorpertyDict = searchPorpertyDict

    def __del__(self):
        '''
        Warning: when script exits, module ctypes may be None,
        ctypes sometimes is None before all controls are destoryed
        but __del__ needs method in dll(which needs ctypes) to release resources
        '''
        if self._element:
            ClientObject.dll.ReleaseElement(self._element)
            self._element = 0
        if self._name:
            ClientObject.dll.FreeBSTR(self._name)
            self._name = 0
        if self._className:
            ClientObject.dll.FreeBSTR(self._className)
            self._className = 0
        if self._automationId:
            ClientObject.dll.FreeBSTR(self._automationId)
            self._automationId = 0

    def SetSearchFromControl(self, searchFromControl):
        '''searchFromControl: control'''
        self.searchFromControl = searchFromControl

    def SetSearchDepth(self, searchDepth):
        '''searchDepth: integer'''
        self.searchDepth = searchDepth

    def AddSearchProperty(self, **searchPorpertyDict):
        '''searchPorpertyDict: dict'''
        self.searchPorpertyDict.update(searchPorpertyDict)

    def RemoveSearchProperty(self, **searchPorpertyDict):
        for key in searchPorpertyDict:
            del self.searchPorpertyDict[key]

    def CompareFunction(self, control):
        '''This function defines how to search, return True if found'''
        if 'ControlType' in self.searchPorpertyDict:
            if self.searchPorpertyDict['ControlType'] != control.ControlType:
                return False
        if 'ClassName' in self.searchPorpertyDict:
            if self.searchPorpertyDict['ClassName'] != control.ClassName:
                return False
        if 'AutomationId' in self.searchPorpertyDict:
            if self.searchPorpertyDict['AutomationId'] != control.AutomationId:
                return False
        if 'SubName' in self.searchPorpertyDict:
            if self.searchPorpertyDict['SubName'] not in control.Name:
                return False
        if 'Name' in self.searchPorpertyDict:
            if self.searchPorpertyDict['Name'] != control.Name:
                return False
        return True

    def Exists(self, maxSearchSeconds = 5, searchIntervalSeconds = SEARCH_INTERVAL):
        '''Find control every searchIntervalSeconds seconds in maxSearchSeconds seconds, if found, return the control'''
        if len(self.searchPorpertyDict) == 0:
            raise LookupError("control's searchPorpertyDict must not be empty!")
        self._element = 0
        if self.searchFromControl:
            self.searchFromControl.Element # search searchFromControl first before timing
        start = time.clock()
        while True:
            time.sleep(searchIntervalSeconds)
            control = FindControl(self.searchFromControl, self.CompareFunction, self.searchDepth, False, self.foundIndex)
            if control:
                self._element = control.Element
                control._element = 0 # control will be destroyed, but the element needs to be stroed in self._element
                return True
            else:
                if time.clock() - start > maxSearchSeconds:
                    return False

    def Refind(self, maxSearchSeconds = 15, searchIntervalSeconds = SEARCH_INTERVAL, raiseException = True):
        '''Refind the control every searchIntervalSeconds seconds in maxSearchSeconds seconds, raise an LookupError if timed out'''
        if not self.Exists(maxSearchSeconds, searchIntervalSeconds):
            if raiseException:
                info = 'Find Control Time Out: {'
                for key in self.searchPorpertyDict:
                    if key == 'ControlType':
                        info += '{0}: {1}, '.format(key, ControlTypeNameDict[self.searchPorpertyDict[key]])
                    else:
                        info += '{0}: {1}, '.format(key, repr(self.searchPorpertyDict[key])) # example Name : 'Notepad'
                info += '}'
                raise LookupError(info)

    @property
    def Element(self):
        '''Return value of control's IUIAutomationElement'''
        if not self._element:
            #print 'property Element call Refind'
            self.Refind(searchIntervalSeconds = self.searchWaitTime)
        return self._element

    @property
    def Name(self):
        '''Return unicode Name'''
        if self._name:
            ClientObject.dll.FreeBSTR(self._name)
        self._name = ClientObject.dll.GetElementName(self.Element)
        if self._name:
            name = ctypes.c_wchar_p(self._name)
            return name.value
        return ''

    @property
    def ControlType(self):
        '''Return an integer in class ControlType'''
        return ClientObject.dll.GetElementControlType(self.Element)

    @property
    def ControlTypeName(self):
        '''Return str ControlTypeName'''
        return ControlTypeNameDict[self.ControlType]

    @property
    def ClassName(self):
        '''Return unicode ClassName'''
        if self._className:
            ClientObject.dll.FreeBSTR(self._className)
        self._className = ClientObject.dll.GetElementClassName(self.Element)
        if self._className:
            name = ctypes.c_wchar_p(self._className)
            return name.value
        return ''

    @property
    def AutomationId(self):
        '''Return unicode AutomationId'''
        if self._automationId:
            ClientObject.dll.FreeBSTR(self._automationId)
        self._automationId = ClientObject.dll.GetElementAutomationId(self.Element)
        if self._automationId:
            name = ctypes.c_wchar_p(self._automationId)
            return name.value
        return ''

    @property
    def IsEnabled(self):
        '''Return bool'''
        return ClientObject.dll.GetElementIsEnabled(self.Element)

    @property
    def IsOffScreen(self):
        '''Return bool'''
        return ClientObject.dll.GetElementIsOffscreen(self.Element)

    @property
    def BoundingRectangle(self):
        '''Return tuple (left, top, right, bottom)'''
        rect = Rect()
        ClientObject.dll.GetElementBoundingRectangle(self.Element, ctypes.byref(rect))
        return (rect.left, rect.top, rect.right, rect.bottom)

    @property
    def Handle32(self):
        '''Return control's handle'''
        return ClientObject.dll.GetElementHandle(self.Element)

    def SetFocus(self):
        '''Make the control have focus'''
        ClientObject.dll.SetElementFocus(self.Element)

    def MoveCursor(self, xRatio = 0.5, yRatio = 0.5):
        '''Move cursor to control's rect, default to center'''
        left, top, right, bottom = self.BoundingRectangle
        x = int(left + (right - left) * xRatio)
        y = int(top + (bottom - top) * yRatio)
        Win32API.MouseMoveTo(x, y)

    def MoveCursorToMyCenter(self):
        '''Move cursor to control's center'''
        self.MoveCursor()

    def Click(self, ratioX = 0.5, ratioY = 0.5, simulateMove = True):
        '''
        Click(0.5, 0.5): click center
        Click(10, 10): click left+10, top+10
        simulateMove：bool, if True, first move cursor to control smoothly
        '''
        left, top, right, bottom = self.BoundingRectangle
        if type(ratioX) is float:
            w = right - left
            x = left + int(w * ratioX)
        else:
            x = left + ratioX
        if type(ratioY) is float:
            h = bottom - top
            y = top + int(h * ratioY)
        else:
            y = top + ratioY
        if simulateMove:
            Win32API.MouseMoveTo(x, y)
        Win32API.MouseClick(x, y)

    def MiddleClick(self, ratioX = 0.5, ratioY = 0.5, simulateMove = True):
        '''
        Click(0.5, 0.5): click center
        Click(10, 10): click left+10, top+10
        simulateMove：bool, if True, first move cursor to control smoothly
        '''
        left, top, right, bottom = self.BoundingRectangle
        if type(ratioX) is float:
            w = right - left
            x = left + int(w * ratioX)
        else:
            x = left + ratioX
        if type(ratioY) is float:
            h = bottom - top
            y = top + int(h * ratioY)
        else:
            y = top + ratioY
        if simulateMove:
            Win32API.MouseMoveTo(x, y)
        Win32API.MouseMiddleClick(x, y)

    def RightClick(self, ratioX = 0.5, ratioY = 0.5, simulateMove = True):
        '''
        RightClick(0.5, 0.5): right click center
        RightClick(10, 10): right click left+10, top+10
        simulateMove：bool, if True, first move cursor to control smoothly
        '''
        left, top, right, bottom = self.BoundingRectangle
        if type(ratioX) is float:
            w = right - left
            x = left + int(w * ratioX)
        else:
            x = left + ratioX
        if type(ratioY) is float:
            h = bottom - top
            y = top + int(h * ratioY)
        else:
            y = top + ratioY
        if simulateMove:
            Win32API.MouseMoveTo(x, y)
        Win32API.MouseRightClick(x, y)

    def DoubleClick(self, ratioX = 0.5, ratioY = 0.5, simulateMove = True):
        '''
        DoubleClick(0.5, 0.5): double click center
        DoubleClick(10, 10): double click left+10, top+10
        simulateMove：bool, if True, first move cursor to control smoothly
        '''
        left, top, right, bottom = self.BoundingRectangle
        if type(ratioX) is float:
            w = right - left
            x = left + int(w * ratioX)
        else:
            x = left + ratioX
        if type(ratioY) is float:
            h = bottom - top
            y = top + int(h * ratioY)
        else:
            y = top + ratioY
        if simulateMove:
            Win32API.MouseMoveTo(x, y)
        Win32API.MouseClick(x, y)
        time.sleep(Win32API.GetDoubleClickTime() * 1.0 / 2000)
        Win32API.MouseClick(x, y)

    def GetParentControl(self):
        '''Return Control'''
        comEle = ClientObject.dll.GetParentElement(self.Element)
        if comEle:
            return Control.CreateControlFromElement(comEle)

    def GetFirstChildControl(self):
        '''Return Control'''
        comEle = ClientObject.dll.GetFirstChildElement(self.Element)
        if comEle:
            return Control.CreateControlFromElement(comEle)

    def GetLastChildControl(self):
        '''Return Control'''
        comEle = ClientObject.dll.GetLastChildElement(self.Element)
        if comEle:
            return Control.CreateControlFromElement(comEle)

    def GetNextSiblingControl(self):
        '''Return Control'''
        comEle = ClientObject.dll.GetNextSiblingElement(self.Element)
        if comEle:
            return Control.CreateControlFromElement(comEle)

    def GetPreviousSiblingControl(self):
        '''Return Control'''
        comEle = ClientObject.dll.GetPreviousSiblingElement(self.Element)
        if comEle:
            return Control.CreateControlFromElement(comEle)

    def GetChildren(self):
        '''Return a list of control's children'''
        children = []
        child = self.GetFirstChildControl()
        while child:
            children.append(child)
            child = child.GetNextSiblingControl()
        return children

    def ShowWindow(self, cmdShow):
        '''ShowWindow(ShowWindow.Show), see values in class ShowWindow'''
        hWnd = self.Handle32
        if hWnd:
            return Win32API.ShowWindow(hWnd, cmdShow)

    def MoveWindow(self, x, y, width, height, repaint = 0):
        '''ShowWindow(ShowWindow.Show), see values in class ShowWindow'''
        hWnd = self.Handle32
        if hWnd:
            return Win32API.MoveWindow(hWnd, x, y, width, height, repaint)

    def GetWindowText(self):
        hWnd = self.Handle32
        if hWnd:
            return Win32API.GetWindowText(hWnd)

    def SetWindowText(self, text):
        hWnd = self.Handle32
        if hWnd:
            return Win32API.SetWindowText(hWnd, text)

    def GetPixel(self, x, y):
        hWnd = self.Handle32
        if hWnd:
            return Win32API.GetPixel(x, y, hWnd)

    def Convert(self):
        '''
        Convert Control to a specific Control
        for example: if self's ControlType is EditControl, return an EditControl
        '''
        return Control.CreateControlFromControl(self)

    @staticmethod
    def CreateControlFromElement(element):
        '''element: value of IUIAutomationElement'''
        controlType = ClientObject.dll.GetElementControlType(element)
        controlDict = {
            ControlType.ButtonControl : ButtonControl,
            ControlType.CalendarControl : CalendarControl,
            ControlType.CheckBoxControl : CheckBoxControl,
            ControlType.ComboBoxControl : ComboBoxControl,
            ControlType.CustomControl : CustomControl,
            ControlType.DataGridControl : DataGridControl,
            ControlType.DataItemControl : DataItemControl,
            ControlType.DocumentControl : DocumentControl,
            ControlType.EditControl : EditControl,
            ControlType.GroupControl : GroupControl,
            ControlType.HeaderControl : HeaderControl,
            ControlType.HeaderItemControl : HeaderItemControl,
            ControlType.HyperlinkControl : HyperlinkControl,
            ControlType.ImageControl : ImageControl,
            ControlType.ListControl : ListControl,
            ControlType.ListItemControl : ListItemControl,
            ControlType.MenuBarControl : MenuBarControl,
            ControlType.MenuControl : MenuControl,
            ControlType.MenuItemControl : MenuItemControl,
            ControlType.PaneControl : PaneControl,
            ControlType.ProgressBarControl : ProgressBarControl,
            ControlType.RadioButtonControl : RadioButtonControl,
            ControlType.ScrollBarControl : ScrollBarControl,
            ControlType.SemanticZoomControl : SemanticZoomControl,
            ControlType.SeparatorControl : SeparatorControl,
            ControlType.SliderControl : SliderControl,
            ControlType.SpinnerControl : SpinnerControl,
            ControlType.SplitButtonControl : SplitButtonControl,
            ControlType.StatusBarControl : StatusBarControl,
            ControlType.TabControl : TabControl,
            ControlType.TabItemControl : TabItemControl,
            ControlType.TableControl : TableControl,
            ControlType.TextControl : TextControl,
            ControlType.ThumbControl : ThumbControl,
            ControlType.TitleBarControl : TitleBarControl,
            ControlType.ToolBarControl : ToolBarControl,
            ControlType.ToolTipControl : ToolTipControl,
            ControlType.TreeControl : TreeControl,
            ControlType.TreeItemControl : TreeItemControl,
            ControlType.WindowControl : WindowControl,
        }
        if controlType in controlDict:
            return controlDict[controlType](element)

    @staticmethod
    def CreateControlFromControl(control):
        '''
        control: Control, will add ref for control's element
        return a specific Control
        for example: if control's ControlType is EditControl, return an EditControl
        '''
        controlType = control.ControlType
        newControl = Control.CreateControlFromElement(control.Element)
        if newControl:
            ClientObject.dll.ElementAddRef(control.Element)
            return newControl

    def __str__(self):
        if IsPy3:
            return 'ControlType: {0}    ClassName: {1}    AutomationId: {2}    Rect: {3}    Name: {4}'.format(self.ControlTypeName, self.ClassName, self.AutomationId, self.BoundingRectangle, self.Name)
        else:
            strClassName = self.ClassName.encode('gbk')
            strAutomationId = self.AutomationId.encode('gbk')
            try:
                strName = self.Name.encode('gbk')
            except BaseException:
                strName = 'Error occured: Name can\'t be converted to gbk, try unicode'
            return 'ControlType: {0}    ClassName: {1}    AutomationId: {2}    Rect: {3}    Name: {4}'.format(self.ControlTypeName, strClassName, strAutomationId, self.BoundingRectangle, strName)


    # def __repr__(self):
        # return '[{0}]'.format(self)

    def __unicode__(self):
        return u'ControlType: {0}    ClassName: {1}    AutomationId: {2}    Rect: {3}    Name: {4}'.format(self.ControlTypeName, self.ClassName, self.AutomationId, self.BoundingRectangle, self.Name)


#Patterns -----
class InvokePattern():
    def Invoke(self):
        '''invoke'''
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_InvokePatternId)
        if pattern:
            ClientObject.dll.InvokePatternInvoke(pattern)
            ClientObject.dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('InvokePattern is not supported!', ConsoleColor.Yellow)


class TogglePattern():
    def Toggle(self):
        '''Toggle or UnToggle the control'''
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_TogglePatternId)
        if pattern:
            ClientObject.dll.TogglePatternToggle(pattern)
            ClientObject.dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('TogglePattern is not supported!', ConsoleColor.Yellow)


    def CurrentToggleState(self):
        '''Return an integer of class ToggleState'''
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_TogglePatternId)
        if pattern:
            state = ClientObject.dll.TogglePatternCurrentToggleState(pattern)
            ClientObject.dll.ReleasePattern(pattern)
            return state
        else:
            Logger.WriteLine('TogglePattern is not supported!', ConsoleColor.Yellow)


class ExpandCollapsePattern():
    def Expand(self):
        '''Expand the control'''
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_ExpandCollapsePatternId)
        if pattern:
            ClientObject.dll.ExpandCollapsePatternExpand(pattern)
            ClientObject.dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('ExpandCollapsePattern is not supported!', ConsoleColor.Yellow)

    def Collapse(self):
        '''Collapse the control'''
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_ExpandCollapsePatternId)
        if pattern:
            ClientObject.dll.ExpandCollapsePatternCollapse(pattern)
            ClientObject.dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('ExpandCollapsePattern is not supported!', ConsoleColor.Yellow)


class ValuePattern():
    def CurrentValue(self):
        '''Return unicode string'''
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_ValuePatternId)
        if pattern:
            value = ClientObject.dll.ValuePatternCurrentValue(pattern)
            c_value = ctypes.c_wchar_p(value)
            ClientObject.dll.ReleasePattern(pattern)
            return c_value.value
        else:
            Logger.WriteLine('ValuePattern is not supported!', ConsoleColor.Yellow)

    def SetValue(self, value):
        '''Set unicode string to control's value'''
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_ValuePatternId)
        if pattern:
            c_value = ctypes.c_wchar_p(value)
            value = ClientObject.dll.ValuePatternSetValue(pattern, c_value)
            ClientObject.dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('ValuePattern is not supported!', ConsoleColor.Yellow)


class ScrollItemPattern():
    def ScrollIntoView(self):
        '''Scroll the contorl into view, so it can be seen'''
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_ScrollItemPatternId)
        if pattern:
            ClientObject.dll.ScrollItemPatternScrollIntoView(pattern)
            ClientObject.dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('ScrollItemPattern is not supported!', ConsoleColor.Yellow)


class ScrollPattern():
    def CurrentHorizontallyScrollable(self):
        '''Return bool'''
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_ScrollPatternId)
        if pattern:
            scroll = ClientObject.dll.ScrollPatternCurrentHorizontallyScrollable(pattern)
            ClientObject.dll.ReleasePattern(pattern)
            return scroll
        else:
            Logger.WriteLine('ScrollPattern is not supported!', ConsoleColor.Yellow)

    def CurrentHorizontalViewSize(self):
        '''Return integer'''
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_ScrollPatternId)
        if pattern:
            size = ClientObject.dll.ScrollPatternCurrentHorizontalViewSize(pattern)
            ClientObject.dll.ReleasePattern(pattern)
            return size
        else:
            Logger.WriteLine('ScrollPattern is not supported!', ConsoleColor.Yellow)

    def CurrentHorizontalScrollPercent(self):
        '''Return integer'''
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_ScrollPatternId)
        if pattern:
            percent = ClientObject.dll.ScrollPatternCurrentHorizontalScrollPercent(pattern)
            ClientObject.dll.ReleasePattern(pattern)
            return percent
        else:
            Logger.WriteLine('ScrollPattern is not supported!', ConsoleColor.Yellow)


    def CurrentVerticallyScrollable(self):
        '''Return bool'''
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_ScrollPatternId)
        if pattern:
            scroll = ClientObject.dll.ScrollPatternCurrentVerticallyScrollable(pattern)
            ClientObject.dll.ReleasePattern(pattern)
            return scroll
        else:
            Logger.WriteLine('ScrollPattern is not supported!', ConsoleColor.Yellow)

    def CurrentVerticalViewSize(self):
        '''Return integer'''
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_ScrollPatternId)
        if pattern:
            size = ClientObject.dll.ScrollPatternCurrentVerticalViewSize(pattern)
            ClientObject.dll.ReleasePattern(pattern)
            return size
        else:
            Logger.WriteLine('ScrollPattern is not supported!', ConsoleColor.Yellow)

    def CurrentVerticalScrollPercent(self):
        '''Return integer'''
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_ScrollPatternId)
        if pattern:
            percent = ClientObject.dll.ScrollPatternCurrentVerticalScrollPercent(pattern)
            ClientObject.dll.ReleasePattern(pattern)
            return percent
        else:
            Logger.WriteLine('ScrollPattern is not supported!', ConsoleColor.Yellow)


    def SetScrollPercent(self, horizontalPercent, verticalPercent):
        '''Need two integers as parameters'''
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_ScrollPatternId)
        if pattern:
            ClientObject.dll.ScrollPatternSetScrollPercent(pattern, horizontalPercent, verticalPercent)
            ClientObject.dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('ScrollPattern is not supported!', ConsoleColor.Yellow)


class SelectionPattern():
    def GetCurrentSelection(self):
        '''Return an IUIAutomationElementArray'''
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_SelectionPatternId)
        if pattern:
            pElementArray = ClientObject.dll.SelectionPatternGetCurrentSelection(pattern)
            ClientObject.dll.ReleasePattern(pattern)
            return pElementArray
        else:
            Logger.WriteLine('SelectionPattern is not supported!', ConsoleColor.Yellow)


class SelectionItemPattern():
    def Select(self):
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_SelectionItemPatternId)
        if pattern:
            ClientObject.dll.SelectionItemPatternSelect(pattern)
            ClientObject.dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('SelectionItemPattern is not supported!', ConsoleColor.Yellow)

    def AddToSelection(self):
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_SelectionItemPatternId)
        if pattern:
            ClientObject.dll.SelectionItemPatternAddToSelection(pattern)
            ClientObject.dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('SelectionItemPattern is not supported!', ConsoleColor.Yellow)

    def RemoveFromSelection(self):
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_SelectionItemPatternId)
        if pattern:
            ClientObject.dll.SelectionItemPatternRemoveFromSelection(pattern)
            ClientObject.dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('SelectionItemPattern is not supported!', ConsoleColor.Yellow)

    def CurrentIsSelected(self):
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_SelectionItemPatternId)
        if pattern:
            isSelect = ClientObject.dll.SelectionItemPatternCurrentIsSelected(pattern)
            ClientObject.dll.ReleasePattern(pattern)
            return isSelect
        else:
            Logger.WriteLine('SelectionItemPattern is not supported!', ConsoleColor.Yellow)


class RangeValuePattern():
    def CurrentValue(self):
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_RangeValuePatternId)
        if pattern:
            value = ClientObject.dll.RangeValuePatternCurrentValue(pattern)
            ClientObject.dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('RangeValuePattern is not supported!', ConsoleColor.Yellow)

    def SetValue(self, value):
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_RangeValuePatternId)
        if pattern:
            ClientObject.dll.RangeValuePatternSetValue(pattern, value)
            ClientObject.dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('RangeValuePattern is not supported!', ConsoleColor.Yellow)

    def CurrentMaximum(self):
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_RangeValuePatternId)
        if pattern:
            value = ClientObject.dll.RangeValuePatternCurrentMaximum(pattern)
            ClientObject.dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('RangeValuePattern is not supported!', ConsoleColor.Yellow)

    def CurrentMinimum(self):
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_RangeValuePatternId)
        if pattern:
            value = ClientObject.dll.RangeValuePatternCurrentMinimum(pattern)
            ClientObject.dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('RangeValuePattern is not supported!', ConsoleColor.Yellow)


class WindowPattern():
    def CurrentWindowVisualState(self):
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_WindowPatternId)
        if pattern:
            value = ClientObject.dll.WindowPatternCurrentWindowVisualState(pattern)
            ClientObject.dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('WindowPattern is not supported!', ConsoleColor.Yellow)

    def SetWindowVisualState(self, value):
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_WindowPatternId)
        if pattern:
            ClientObject.dll.WindowPatternSetWindowVisualState(pattern, value)
            ClientObject.dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('WindowPattern is not supported!', ConsoleColor.Yellow)

    def CurrentCanMaximize(self):
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_WindowPatternId)
        if pattern:
            value = ClientObject.dll.WindowPatternCurrentCanMaximize(pattern)
            ClientObject.dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('WindowPattern is not supported!', ConsoleColor.Yellow)

    def Maximize(self):
        if self.CurrentCanMaximize():
            self.SetWindowVisualState(WindowVisualState.WindowVisualState_Maximized)

    def CurrentCanMinimize(self):
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_WindowPatternId)
        if pattern:
            value = ClientObject.dll.WindowPatternCurrentCanMinimize(pattern)
            ClientObject.dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('WindowPattern is not supported!', ConsoleColor.Yellow)

    def Minimize(self):
        if self.CurrentCanMinimize():
            self.SetWindowVisualState(WindowVisualState.WindowVisualState_Minimized)

    def Normal(self):
        self.SetWindowVisualState(WindowVisualState.WindowVisualState_Normal)

    def IsMaximize(self):
        return self.CurrentWindowVisualState() == WindowVisualState.WindowVisualState_Maximized

    def IsMinimize(self):
        return self.CurrentWindowVisualState() == WindowVisualState.WindowVisualState_Minimized

    def CurrentIsModal(self):
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_WindowPatternId)
        if pattern:
            value = ClientObject.dll.WindowPatternCurrentIsModal(pattern)
            ClientObject.dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('WindowPattern is not supported!', ConsoleColor.Yellow)

    def SetTopmost(self, isTopmost = True):
        return Win32API.SetWindowTopmost(self.Handle32, isTopmost)

    def CurrentIsTopmost(self):
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_WindowPatternId)
        if pattern:
            value = ClientObject.dll.WindowPatternCurrentIsTopmost(pattern)
            ClientObject.dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('WindowPattern is not supported!', ConsoleColor.Yellow)

    def MoveToCenter(self):
        left, top, right, bottom = self.BoundingRectangle
        width, height = right - left, bottom - top
        screenWidth, screenHeight = Win32API.GetScreenSize()
        x, y = (screenWidth-width)//2, (screenHeight-height)//2
        if x < 0: x = 0
        if y < 0: y = 0
        return Win32API.SetWindowPos(self.Handle32, SWP.HWND_TOP, x, y, 0, 0, SWP.SWP_NOSIZE)

    def Close(self):
        pattern = ClientObject.dll.GetElementPattern(self.Element, PatternId.UIA_WindowPatternId)
        if pattern:
            ClientObject.dll.WindowPatternClose(pattern)
            ClientObject.dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('WindowPattern is not supported!', ConsoleColor.Yellow)

    def MetroClose(self):
        window = WindowControl(searchDepth = 1, ClassName = METRO_WINDOW_CLASS_NAME)
        if window.Exists(0, 0):
            screenWidth, screenHeight = Win32API.GetScreenSize()
            Win32API.MouseMoveTo(screenWidth//2, 0)
            Win32API.MouseDragTo(screenWidth//2, 0, screenWidth//2, screenHeight)
            time.sleep(1)
        else:
            Logger.WriteLine('Window is not Metro!', ConsoleColor.Yellow)

    def SetActive(self):
        return Win32API.SetForegroundWindow(self.Handle32)


class ButtonControl(Control, InvokePattern, TogglePattern, ExpandCollapsePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ButtonControl)


class CalendarControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.CalendarControl)


class CheckBoxControl(Control, TogglePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.CheckBoxControl)


class ComboBoxControl(Control, ValuePattern, ExpandCollapsePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ComboBoxControl)

    def Select(self, name):
        self.Expand()
        listItemControl = ListItemControl(searchFromControl = self, ControlType = ControlType.ListItemControl, Name = name)
        if listItemControl.Exists(1, SEARCH_INTERVAL):
            listItemControl.ScrollIntoView()
            listItemControl.Click()


class CustomControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.CustomControl)


class DataGridControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.DataGridControl)


class DataItemControl(Control, ValuePattern, TogglePattern, ExpandCollapsePattern, SelectionItemPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.DataItemControl)


class DocumentControl(Control, ValuePattern, ScrollPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.DocumentControl)


class EditControl(Control, ValuePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.EditControl)


class GroupControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.GroupControl)


class HeaderControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.HeaderControl)


class HeaderItemControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.HeaderItemControl)


class HyperlinkControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.HyperlinkControl)


class ImageControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ImageControl)


class ListControl(Control, ScrollPattern, SelectionPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ListControl)

    def GetSelectedItems(self):
        '''Return a list of children which are selected'''
        lists = []
        iUIAutomationElementArray = self.GetCurrentSelection()
        if iUIAutomationElementArray:
            length = ClientObject.dll.ElementArrayGetLength(iUIAutomationElementArray)
            for i in range(length):
                lists.append(Control.CreateControlFromElement(ClientObject.dll.ElementArrayGetElement(iUIAutomationElementArray, i)))
            ClientObject.dll.ReleaseElementArray(iUIAutomationElementArray)
        return lists


class ListItemControl(Control, ExpandCollapsePattern, SelectionItemPattern, InvokePattern, TogglePattern, ScrollItemPattern, ValuePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ListItemControl)


class MenuBarControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.MenuBarControl)


class MenuControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.MenuControl)


class MenuItemControl(Control, ExpandCollapsePattern, SelectionItemPattern, InvokePattern, TogglePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.MenuItemControl)


class PaneControl(Control, ScrollPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.PaneControl)


class ProgressBarControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ProgressBarControl)


class RadioButtonControl(Control, SelectionItemPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.RadioButtonControl)


class ScrollBarControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ScrollBarControl)


class SemanticZoomControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.SemanticZoomControl)


class SeparatorControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.SeparatorControl)


class SliderControl(Control, RangeValuePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.SliderControl)


class SpinnerControl(Control, RangeValuePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.SpinnerControl)


class SplitButtonControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.SplitButtonControl)


class StatusBarControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.StatusBarControl)


class TabControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.TabControl)


class TabItemControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.TabItemControl)


class TableControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.TableControl)


class TextControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.TextControl)


class ThumbControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ThumbControl)


class TitleBarControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.TitleBarControl)


class ToolBarControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ToolBarControl)


class ToolTipControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ToolTipControl)


class TreeControl(Control, ScrollPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.TreeControl)


class TreeItemControl(Control, ExpandCollapsePattern, InvokePattern, TogglePattern, ScrollItemPattern, SelectionItemPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.TreeItemControl)


class WindowControl(Control, WindowPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 10000, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.WindowControl)


class Logger():
    LogFile = '@AutomationLog.txt'
    LineSep = '\n'
    @staticmethod
    def Write(log, consoleColor = -1, writeToFile = True):
        '''
        consoleColor: value in class ConsoleColor, such as ConsoleColor.DarkGreen
        if consoleColor == -1, use default color
        '''
        isValidColor = (consoleColor >= ConsoleColor.Black and consoleColor <= ConsoleColor.White)
        if isValidColor:
            Win32API.SetConsoleColor(consoleColor)
        try:
            sys.stdout.write(log)
        except Exception as e:
            Win32API.SetConsoleColor(ConsoleColor.Red)
            sys.stdout.write(str(type(e)) + ' can\'t print the log!')
        if isValidColor:
            Win32API.ResetConsoleColor()
        if not writeToFile:
            return
        if IsPy3:
            logFile = open(Logger.LogFile, 'a+', encoding = 'utf-8')
        else:
            logFile = codecs.open(Logger.LogFile, 'a+', 'utf-8')
        try:
            logFile.write(log)
            # logFile.flush() # need flush in python 3, otherwise log won't be saved
        except Exception as ex:
            logFile.close()
            sys.stdout.write('can not write log with exception: {0} {1}'.format(type(ex), ex))

    @staticmethod
    def WriteLine(log, consoleColor = -1, writeToFile = True):
        '''
        consoleColor: value in class ConsoleColor, such as ConsoleColor.DarkGreen
        if consoleColor == -1, use default color
        '''
        Logger.Write(log + Logger.LineSep, consoleColor, writeToFile)

    @staticmethod
    def Log(log, consoleColor = -1, writeToFile = True):
        '''
        consoleColor: value in class ConsoleColor, such as ConsoleColor.DarkGreen
        if consoleColor == -1, use default color
        '''
        t = time.localtime()
        log = '{0}-{1:02}-{2:02} {3:02}:{4:02}:{5:02} - {6}{7}'.format(t.tm_year, t.tm_mon, t.tm_mday,
            t.tm_hour, t.tm_min, t.tm_sec, log, Logger.LineSep)
        Logger.Write(log, consoleColor, writeToFile)

    @staticmethod
    def DeleteLog():
        if os.path.exists(Logger.LogFile):
            os.remove(Logger.LogFile)

def GetRootControl():
    return Control.CreateControlFromElement(ClientObject.dll.GetRootElement())

def GetFocusedControl():
    return Control.CreateControlFromElement(ClientObject.dll.GetFocusedElement())

def GetForegroundControl():
    '''return ControlFromHandle(Win32API.GetForegroundWindow())'''
    focusEle = ClientObject.dll.GetFocusedElement()
    parentEle = focusEle
    elementList = []
    while parentEle:
        elementList.insert(0, parentEle)
        parentEle = ClientObject.dll.GetParentElement(parentEle)
    if len(elementList) == 1:
        parentEle = elementList[0]
    else:
        parentEle = elementList[1]
    return Control.CreateControlFromElement(parentEle)

def ControlFromPoint(x, y):
    element = ClientObject.dll.ElementFromPoint(x, y)
    return Control.CreateControlFromElement(element)

def ControlFromCursor():
    x, y = Win32API.GetCursorPos()
    return ControlFromPoint(x, y)
    # return Control.CreateControlFromElement(ClientObject.dll.ElementFromHandle(Win32API.WindowFromPoint(x, y)))

def ControlFromHandle(handle):
    return Control.CreateControlFromElement(ClientObject.dll.ElementFromHandle(handle))

def LogControl(control, depth = 0, showAllName = True, showMore = False):
    '''
    control: Control
    depth: integer
    showAllName: bool
    showMore: bool
    '''
    name = control.Name
    if not showAllName and name and len(name) > 30:
        name = name[:30] + '...'
    indent = ' ' * depth * 4
    Logger.Write('{0}ControlType: '.format(indent))
    Logger.Write(control.ControlTypeName, ConsoleColor.DarkGreen)
    Logger.Write('    ClassName: ')
    Logger.Write(control.ClassName, ConsoleColor.DarkGreen)
    Logger.Write('    AutomationId: ')
    Logger.Write(control.AutomationId, ConsoleColor.DarkGreen)
    Logger.Write('    Rect: ')
    left, top, right, bottom = control.BoundingRectangle
    Logger.Write(str(control.BoundingRectangle), ConsoleColor.DarkGreen)
    Logger.Write('    Name: ')
    Logger.Write(name, ConsoleColor.DarkGreen)
    Logger.Write('    Handle: ')
    handle = control.Handle32
    Logger.Write('0x{0:x}({1})'.format(handle, handle), ConsoleColor.DarkGreen)
    if showMore:
        Logger.Write('    SupportedPattern:')
        for key in PatternDict:
            pattern = ClientObject.dll.GetElementPattern(control.Element, key)
            if pattern:
                ClientObject.dll.ReleasePattern(pattern)
                Logger.Write(' ' + PatternDict[key], ConsoleColor.DarkGreen)
    Logger.Write(Logger.LineSep)

def EnumContorlUnderCursor(showAllName = True, showMore = False):
    control = ControlFromCursor()
    lists = []
    while control:
        lists.insert(0, control)
        control = control.GetParentControl()
    for (i, control) in enumerate(lists):
        LogControl(control, i, showAllName, showMore)

def EnumControl(control, maxDepth = 10000, showAllName = True, showMore = False):
    '''
    control: Control
    maxDepth: integer
    showAllName: bool
    showMore: bool
    '''
    depth = 0
    LogControl(control, depth, showAllName, showMore)
    if maxDepth <= 0:
        return
    child = control.GetFirstChildControl()
    controlList = [child]
    while depth >= 0:
        lastControl = controlList[-1]
        if lastControl:
            LogControl(lastControl, depth + 1, showAllName, showMore)
            child = lastControl.GetNextSiblingControl()
            controlList[depth] = child
            if depth < maxDepth - 1:
                child = lastControl.GetFirstChildControl()
                if child:
                    depth += 1
                    controlList.append(child)
        else:
            del controlList[depth]
            depth -= 1

def FindControl(control, compareFunc, maxDepth=10000, findFromSelf = False, foundIndex = 1):
    '''
    control: Control
    compareFunc: compare function, should True or False
    maxDepth: integer
    findFromSelf: bool
    foundIndex: integer, value must be greater or equal to 1
    '''
    foundCount = 0
    if not control:
        control = GetRootControl()
    depth = 0
    if findFromSelf and compareFunc(control):
        foundCount += 1
        if foundCount == foundIndex:
            return control.Convert()
    if maxDepth <= 0:
        return
    child = control.GetFirstChildControl()
    controlList = [child]
    while depth >= 0:
        lastControl = controlList[-1]
        if lastControl:
            if compareFunc(lastControl):
                foundCount += 1
                if foundCount == foundIndex:
                    return lastControl.Convert()
            child = lastControl.GetNextSiblingControl()
            controlList[depth] = child
            if depth < maxDepth - 1:
                child = lastControl.GetFirstChildControl()
                if child:
                    depth += 1
                    controlList.append(child)
        else:
            del controlList[depth]
            depth -= 1

def ShowMetroStartMenu():
    '''Show Metro Strat Menu, only works on Windows 8'''
    paneMenu = PaneControl(searchDepth = 1, ClassName = 'ImmersiveLauncher', AutomationId = 'Start menu window')
    if not paneMenu.Exists(0, 0):
        Win32API.SendKeys('{Win}')
        time.sleep(1)

def ShowDesktop():
    paneTray = PaneControl(searchDepth = 1, ClassName = 'Shell_TrayWnd')
    if paneTray.Exists():
        WM_COMMAND = 0x111
        MIN_ALL = 419
        MIN_ALL_UNDO = 416
        Win32API.PostMessage(paneTray.Handle32, WM_COMMAND, MIN_ALL, 0)
        time.sleep(1)

def ClickCharmBar(buttonId):
    charmBar = WindowControl(searchDepth = 1, ClassName = 'NativeHWNDHost', AutomationId = 'Charm Bar')
    if not charmBar.Exists(0):
        Win32API.SendKeys('{Win}C')
    button = ButtonControl(searchFromControl = charmBar, AutomationId = buttonId)
    if button.Exists(1):
        button.Click()

def RunMetroApp(appName):
    ClickCharmBar('Search')
    searchPane = PaneControl(searchDepth = 1, ClassName = 'SearchPane')
    if searchPane.Exists():
        edit = EditControl(searchFromControl = searchPane, ClassName = 'TouchEditInner')
        edit.SetValue(appName)
        startMenu = PaneControl(searchDepth = 1, ClassName = 'ImmersiveLauncher', AutomationId = 'Start menu window')
        appItem = ListItemControl(searchFromControl = startMenu, Name = appName)
        appItem.Click()
        time.sleep(1)

def RunHotKey(function, startHotKey, stopHotKey = None):
    if len(startHotKey) == 2 and (ctypes.windll.user32.RegisterHotKey(0, 1, startHotKey[0], startHotKey[1])):
        sys.stdout.write('RegisterHotKey {0}\n'.format(startHotKey))
    if len(stopHotKey) == 2 and (ctypes.windll.user32.RegisterHotKey(0, 2, stopHotKey[0], stopHotKey[1])):
        sys.stdout.write('RegisterHotKey {0}\n'.format(stopHotKey))
    from threading import Thread
    funcThread = None
    msg=MSG()
    while ctypes.windll.user32.GetMessageW(ctypes.byref(msg), 0, 0, 0) != 0:
        if msg.message == 0x0312: # WM_HOTKEY=0x0312
            if 1 == msg.wParam and msg.lParam&0x0000FFFF == startHotKey[0] and msg.lParam>>16&0x0000FFFF == startHotKey[1]:
                if not funcThread:
                    funcThread=Thread(None, function)
                if not funcThread.isAlive():
                    funcThread.start()#todo
                

def usage():
    sys.stdout.write('''usage
-h      show command help
-t      delay time, begin to enumerate after Value seconds, this must be an integer
        you can delay a few seconds and make a window active so automation can enumerate the active window
-d      enumerate tree depth, this must be an integer, if it is null, enumerate the whole tree
-r      enumerate from root:desktop window, if it is null, enumerate from foreground window
-f      enumerate from focused control, if it is null, enumerate from foreground window
-c      enumerate the control under cursor
-s      show tree of the the control under cursor
-n      show control full name
-m      show more properties

examples:
automation.py -t3
automation.py -t3 -r -d1 -m -n
automation.py -c -t3

''')

def main():
    import getopt
    sys.stdout.write(str(sys.argv))
    options, args = getopt.getopt(sys.argv[1:], 'hrfcsmnd:t:',
                                  ['help', 'root', 'focus', 'cursor', 'cursorTree', 'showMore', 'showAllName', 'depth=', 'time='])
    root = False
    focus = False
    cursor = False
    cursorTree = False
    showAllName = False
    showMore = False
    depth = 10000
    seconds = 3
    for (o, v) in options:
        if o in ('-h', '-help'):
            usage()
            exit(0)
        elif o in ('-r', '-root'):
            root = True
        elif o in ('-f', '-focus'):
            focus = True
        elif o in ('-c', '-cursor'):
            cursor = True
        elif o in ('-s', '-cursorTree'):
            cursorTree = True
        elif o in ('-n', '-showAllName'):
            showAllName = True
        elif o in ('-m', '-showMore'):
            showMore = True
        elif o in ('-d', '-depth'):
            depth = int(v)
        elif o in ('-t', '-time'):
            seconds = int(v)
    time.sleep(seconds)
    Logger.Log('Starts')
    control = None
    if root:
        control = GetRootControl()
    if focus:
        control = GetFocusedControl()
    if cursor:
        control = ControlFromCursor()
    if cursorTree:
        EnumContorlUnderCursor(showAllName, showMore)
    else:
        if not control:
            control = GetForegroundControl()
        EnumControl(control, depth, showAllName, showMore)
    Logger.Log('Ends' + Logger.LineSep)

if __name__ == '__main__':
    main()
    # control = GetForegroundControl()
    # del control
    # if use control object in global area, must del the control when not use it, otherwise it may throw an exception
    # the exception is ctypes was None when control's __del__ was called!
    # It seems that the module ctypes was deleted before control's __del__ was called
    # ClientObject and control are all global object, ClientObject may be deleted before control,
    # but control's __del__ uses ClientObject's member
    # You'd better not use control object in global area

