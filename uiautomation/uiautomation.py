#!python3
# -*- coding: utf-8 -*-
"""
Author: yinkaisheng@live.com
Source: https://github.com/yinkaisheng/Python-UIAutomation-for-Windows

This module is for UIAutomation on Windows(Windows XP with SP3, Windows Vista and Windows 7/8/8.1/10).
It supports UIAutomation for the applications which implmented IUIAutomation, such as MFC, Windows Form, WPF, Modern UI(Metro UI), Qt and Firefox.
Run 'automation.py -h' for help.

uiautomation is shared under the Apache Licene 2.0.
This means that the code can be freely copied and distributed, and costs nothing to use.
"""
import os
import sys
import time
import datetime
import re
import ctypes
import ctypes.wintypes

IsPy3 = sys.version_info[0] >= 3
if not IsPy3:
    import codecs

AUTHOR_MAIL = 'yinkaisheng@live.com'
METRO_WINDOW_CLASS_NAME = 'Windows.UI.Core.CoreWindow'  # for Windows 8 and 8.1
SEARCH_INTERVAL = 0.5  # search control interval seconds
MAX_MOVE_SECOND = 1  # simulate mouse move or drag max seconds
TIME_OUT_SECOND = 15
OPERATION_WAIT_TIME = 0.5
MAX_PATH = 260


class _AutomationClient:
    _instance = None

    @classmethod
    def instance(cls):
        """Singleton instance (this prevents creation on import)."""
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def __init__(self):
        bin_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bin")
        os.environ["PATH"] = bin_path + os.pathsep + os.environ["PATH"]

        if sys.maxsize > 0xFFFFFFFF:
            if sys.version_info[:2] >= (3, 5):
                self.dll = ctypes.cdll.UIAutomationClient_VC140_X64
            elif sys.version_info[:2] >= (3, 3):
                self.dll = ctypes.cdll.UIAutomationClient_VC100_X64
            else:
                self.dll = ctypes.cdll.UIAutomationClient_VC90_X64
        else:
            if sys.version_info[:2] >= (3, 5):
                self.dll = ctypes.cdll.UIAutomationClient_VC140_X86
            elif sys.version_info[:2] >= (3, 3):
                self.dll = ctypes.cdll.UIAutomationClient_VC100_X86
            else:
                self.dll = ctypes.cdll.UIAutomationClient_VC90_X86

        self.InitFunctionType()

        if not self.dll.InitInstance():
            raise RuntimeError('Can not get an instance of IUIAutomation.\nYou may need to install Windows Update KB971513.\nhttps://github.com/yinkaisheng/WindowsUpdateKB971513ForIUIAutomation')

    def InitFunctionType(self):
        self.dll.FreeBSTR.argtypes = (ctypes.c_size_t, )
        self.dll.GetAutomationObject.restype = ctypes.c_size_t
        self.dll.GetRawTreeWalker.restype = ctypes.c_size_t
        self.dll.ElementAddRef.argtypes = (ctypes.c_size_t, )
        self.dll.CompareElements.argtypes = (ctypes.c_size_t, ctypes.c_size_t)
        self.dll.ReleaseElement.argtypes = (ctypes.c_size_t, )
        self.dll.ReleaseElementArray.argtypes = (ctypes.c_size_t, )
        self.dll.GetRootElement.restype = ctypes.c_size_t
        self.dll.GetFocusedElement.restype = ctypes.c_size_t
        self.dll.ElementFromPoint.restype = ctypes.c_size_t
        self.dll.ElementFromHandle.argtypes = (ctypes.c_size_t, )
        self.dll.ElementFromHandle.restype = ctypes.c_size_t
        self.dll.GetElementName.argtypes = (ctypes.c_size_t, )
        self.dll.GetElementName.restype = ctypes.c_size_t
        self.dll.GetElementControlTypeName.argtypes = (ctypes.c_size_t, )
        self.dll.GetElementControlTypeName.restype = ctypes.c_size_t
        self.dll.GetElementClassName.argtypes = (ctypes.c_size_t, )
        self.dll.GetElementClassName.restype = ctypes.c_size_t
        self.dll.GetElementAutomationId.argtypes = (ctypes.c_size_t, )
        self.dll.GetElementAutomationId.restype = ctypes.c_size_t
        self.dll.GetElementProcessId.argtypes = (ctypes.c_size_t, )
        self.dll.GetElementControlType.argtypes = (ctypes.c_size_t, )
        self.dll.GetElementLocalizedControlType.argtypes = (ctypes.c_size_t, )
        self.dll.GetElementLocalizedControlType.restype = ctypes.c_size_t
        self.dll.GetElementIsEnabled.argtypes = (ctypes.c_size_t, )
        self.dll.GetElementHasKeyboardFocus.argtypes = (ctypes.c_size_t, )
        self.dll.GetElementIsKeyboardFocusable.argtypes = (ctypes.c_size_t, )
        self.dll.GetElementIsOffscreen.argtypes = (ctypes.c_size_t, )
        self.dll.GetElementHandle.argtypes = (ctypes.c_size_t, )
        self.dll.GetElementHandle.restype = ctypes.c_size_t
        self.dll.SetElementFocus.argtypes = (ctypes.c_size_t, )
        self.dll.GetParentElement.argtypes = (ctypes.c_size_t, )
        self.dll.GetParentElement.restype = ctypes.c_size_t
        self.dll.GetNextSiblingElement.argtypes = (ctypes.c_size_t, )
        self.dll.GetNextSiblingElement.restype = ctypes.c_size_t
        self.dll.GetPreviousSiblingElement.argtypes = (ctypes.c_size_t, )
        self.dll.GetPreviousSiblingElement.restype = ctypes.c_size_t
        self.dll.GetFirstChildElement.argtypes = (ctypes.c_size_t, )
        self.dll.GetFirstChildElement.restype = ctypes.c_size_t
        self.dll.GetLastChildElement.argtypes = (ctypes.c_size_t, )
        self.dll.GetLastChildElement.restype = ctypes.c_size_t
        self.dll.ElementArrayGetLength.argtypes = (ctypes.c_size_t, )
        self.dll.ElementArrayGetElement.argtypes = (ctypes.c_size_t, ctypes.c_int)
        self.dll.ElementArrayGetElement.restype = ctypes.c_size_t
        self.dll.GetElementPattern.argtypes = (ctypes.c_size_t, ctypes.c_int)
        self.dll.GetElementPattern.restype = ctypes.c_size_t
        self.dll.ReleasePattern.argtypes = (ctypes.c_size_t, )
        self.dll.InvokePatternInvoke.argtypes = (ctypes.c_size_t, )
        self.dll.TogglePatternToggle.argtypes = (ctypes.c_size_t, )
        self.dll.TogglePatternCurrentToggleState.argtypes = (ctypes.c_size_t, )
        self.dll.ExpandCollapsePatternExpand.argtypes = (ctypes.c_size_t, )
        self.dll.ExpandCollapsePatternCollapse.argtypes = (ctypes.c_size_t, )
        self.dll.ExpandCollapsePatternCurrentExpandCollapseState.argtypes = (ctypes.c_size_t, )
        self.dll.ValuePatternCurrentValue.argtypes = (ctypes.c_size_t, )
        self.dll.ValuePatternCurrentValue.restype = ctypes.c_size_t
        self.dll.ValuePatternSetValue.argtypes = (ctypes.c_size_t, ctypes.c_wchar_p)
        self.dll.ValuePatternCurrentIsReadOnly.argtypes = (ctypes.c_size_t, )
        self.dll.ScrollItemPatternScrollIntoView.argtypes = (ctypes.c_size_t, )
        self.dll.ScrollPatternCurrentHorizontallyScrollable.argtypes = (ctypes.c_size_t, )
        self.dll.ScrollPatternCurrentHorizontalViewSize.argtypes = (ctypes.c_size_t, )
        self.dll.ScrollPatternCurrentHorizontalScrollPercent.argtypes = (ctypes.c_size_t, )
        self.dll.ScrollPatternCurrentVerticallyScrollable.argtypes = (ctypes.c_size_t, )
        self.dll.ScrollPatternCurrentVerticalViewSize.argtypes = (ctypes.c_size_t, )
        self.dll.ScrollPatternCurrentVerticalScrollPercent.argtypes = (ctypes.c_size_t, )
        self.dll.ScrollPatternSetScrollPercent.argtypes = (ctypes.c_size_t, ctypes.c_int, ctypes.c_int)
        self.dll.SelectionPatternGetCurrentSelection.argtypes = (ctypes.c_size_t, )
        self.dll.SelectionPatternGetCurrentSelection.restype = ctypes.c_size_t
        self.dll.SelectionItemPatternSelect.argtypes = (ctypes.c_size_t, )
        self.dll.SelectionItemPatternAddToSelection.argtypes = (ctypes.c_size_t, )
        self.dll.SelectionItemPatternRemoveFromSelection.argtypes = (ctypes.c_size_t, )
        self.dll.SelectionItemPatternCurrentIsSelected.argtypes = (ctypes.c_size_t, )
        self.dll.RangeValuePatternCurrentValue.argtypes = (ctypes.c_size_t, )
        self.dll.RangeValuePatternSetValue.argtypes = (ctypes.c_size_t, ctypes.c_int)
        self.dll.RangeValuePatternCurrentMaximum.argtypes = (ctypes.c_size_t, )
        self.dll.RangeValuePatternCurrentMinimum.argtypes = (ctypes.c_size_t, )
        self.dll.WindowPatternCurrentWindowVisualState.argtypes = (ctypes.c_size_t, )
        self.dll.WindowPatternSetWindowVisualState.argtypes = (ctypes.c_size_t, ctypes.c_int)
        self.dll.WindowPatternCurrentCanMaximize.argtypes = (ctypes.c_size_t, )
        self.dll.WindowPatternCurrentCanMinimize.argtypes = (ctypes.c_size_t, )
        self.dll.WindowPatternCurrentIsModal.argtypes = (ctypes.c_size_t, )
        self.dll.WindowPatternCurrentIsTopmost.argtypes = (ctypes.c_size_t, )
        self.dll.WindowPatternClose.argtypes = (ctypes.c_size_t, )
        self.dll.LegacyIAccessiblePatternSelect.argtypes = (ctypes.c_size_t, ctypes.c_long)
        self.dll.LegacyIAccessiblePatternDoDefaultAction.argtypes = (ctypes.c_size_t, )
        self.dll.LegacyIAccessiblePatternSetValue.argtypes = (ctypes.c_size_t, ctypes.c_wchar_p)
        self.dll.LegacyIAccessiblePatternCurrentChildId.argtypes = (ctypes.c_size_t, )
        self.dll.LegacyIAccessiblePatternCurrentName.argtypes = (ctypes.c_size_t, )
        self.dll.LegacyIAccessiblePatternCurrentName.restype = ctypes.c_size_t
        self.dll.LegacyIAccessiblePatternCurrentValue.argtypes = (ctypes.c_size_t, )
        self.dll.LegacyIAccessiblePatternCurrentValue.restype = ctypes.c_size_t
        self.dll.LegacyIAccessiblePatternCurrentDescription.argtypes = (ctypes.c_size_t, )
        self.dll.LegacyIAccessiblePatternCurrentDescription.restype = ctypes.c_size_t
        self.dll.LegacyIAccessiblePatternCurrentRole.argtypes = (ctypes.c_size_t, )
        self.dll.LegacyIAccessiblePatternCurrentState.argtypes = (ctypes.c_size_t, )
        self.dll.LegacyIAccessiblePatternCurrentHelp.argtypes = (ctypes.c_size_t, )
        self.dll.LegacyIAccessiblePatternCurrentHelp.restype = ctypes.c_size_t
        self.dll.LegacyIAccessiblePatternCurrentKeyboardShortcut.argtypes = (ctypes.c_size_t, )
        self.dll.LegacyIAccessiblePatternCurrentKeyboardShortcut.restype = ctypes.c_size_t
        self.dll.LegacyIAccessiblePatternGetCurrentSelection.argtypes = (ctypes.c_size_t, )
        self.dll.LegacyIAccessiblePatternGetCurrentSelection.restype = ctypes.c_size_t
        self.dll.LegacyIAccessiblePatternCurrentDefaultAction.argtypes = (ctypes.c_size_t, )
        self.dll.LegacyIAccessiblePatternCurrentDefaultAction.restype = ctypes.c_size_t
        self.dll.GridPatternGetItem.argtypes = (ctypes.c_size_t, ctypes.c_int, ctypes.c_int)
        self.dll.GridPatternGetItem.restype = ctypes.c_size_t
        self.dll.GridPatternCurrentRowCount.argtypes = (ctypes.c_size_t, )
        self.dll.GridPatternCurrentColumnCount.argtypes = (ctypes.c_size_t, )
        self.dll.TablePatternCurrentRowHeaders.argtypes = (ctypes.c_size_t, )
        self.dll.TablePatternCurrentRowHeaders.restype = ctypes.c_size_t
        self.dll.TablePatternCurrentColumnHeaders.argtypes = (ctypes.c_size_t, )
        self.dll.TablePatternCurrentColumnHeaders.restype = ctypes.c_size_t
        self.dll.TablePatternCurrentRowOrColumnMajor.argtypes = (ctypes.c_size_t, )
        self.dll.TableItemPatternCurrentRowHeaderItems.argtypes = (ctypes.c_size_t, )
        self.dll.TableItemPatternCurrentRowHeaderItems.restype = ctypes.c_size_t
        self.dll.TableItemPatternCurrentColumnHeaderItems.argtypes = (ctypes.c_size_t, )
        self.dll.TableItemPatternCurrentColumnHeaderItems.restype = ctypes.c_size_t
        self.dll.GridItemPatternCurrentContainingGrid.argtypes = (ctypes.c_size_t, )
        self.dll.GridItemPatternCurrentContainingGrid.restype = ctypes.c_size_t
        self.dll.GridItemPatternCurrentRow.argtypes = (ctypes.c_size_t, )
        self.dll.GridItemPatternCurrentColumn.argtypes = (ctypes.c_size_t, )
        self.dll.GridItemPatternCurrentRowSpan.argtypes = (ctypes.c_size_t, )
        self.dll.GridItemPatternCurrentColumnSpan.argtypes = (ctypes.c_size_t, )
        self.dll.TransformPatternMove.argtypes = (ctypes.c_size_t, ctypes.c_int, ctypes.c_int)
        self.dll.TransformPatternResize.argtypes = (ctypes.c_size_t, ctypes.c_int, ctypes.c_int)
        self.dll.TransformPatternRotate.argtypes = (ctypes.c_size_t, ctypes.c_int)
        self.dll.BitmapCreate.restype = ctypes.c_size_t
        self.dll.BitmapFromWindow.argtypes = (ctypes.c_size_t, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int)
        self.dll.BitmapFromWindow.restype = ctypes.c_size_t
        self.dll.BitmapFromFile.restype = ctypes.c_size_t
        self.dll.BitmapToFile.argtypes = (ctypes.c_size_t, ctypes.c_wchar_p, ctypes.c_wchar_p)
        self.dll.BitmapRelease.argtypes = (ctypes.c_size_t, )
        self.dll.BitmapGetWidthAndHeight.argtypes = (ctypes.c_size_t, )
        self.dll.BitmapGetPixel.argtypes = (ctypes.c_size_t, ctypes.c_uint, ctypes.c_uint)
        self.dll.BitmapSetPixel.argtypes = (ctypes.c_size_t, ctypes.c_uint, ctypes.c_uint, ctypes.c_uint)

        #Windows dll
        ctypes.windll.kernel32.GlobalLock.argtypes = (ctypes.c_void_p, )
        ctypes.windll.kernel32.GlobalLock.restype = ctypes.c_void_p
        ctypes.windll.kernel32.GlobalUnlock.argtypes = (ctypes.c_void_p, )
        ctypes.windll.kernel32.GlobalAlloc.restype = ctypes.c_void_p
        ctypes.windll.kernel32.GetStdHandle.restype = ctypes.c_void_p
        ctypes.windll.kernel32.OpenProcess.restype = ctypes.c_void_p
        ctypes.windll.kernel32.CreateToolhelp32Snapshot.restype = ctypes.c_void_p
        ctypes.windll.kernel32.CloseHandle.argtypes = (ctypes.c_void_p, )
        ctypes.windll.kernel32.TerminateProcess.argtypes = (ctypes.c_void_p, ctypes.c_uint)
        ctypes.windll.user32.GetClipboardData.restype = ctypes.c_void_p
        ctypes.windll.user32.GetWindowDC.restype = ctypes.c_void_p
        ctypes.windll.user32.OpenDesktopW.restype = ctypes.c_void_p
        ctypes.windll.user32.WindowFromPoint.restype = ctypes.c_void_p
        ctypes.windll.user32.SwitchDesktop.argtypes = (ctypes.c_void_p, )
        ctypes.windll.user32.CloseDesktop.argtypes = (ctypes.c_void_p, )

    def __del__(self):
        self.dll.ReleaseInstance()


_rootControl = None


class ControlType:
    """This class defines the values of control type"""
    AppBarControl = 0xc378
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
    ControlType.AppBarControl : 'AppBarControl',
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
    """This class defines the values of pattern id"""
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


class AccessibleRole:
    """This class defines the values of Accessible Role"""
    SystemTitleBar = 0x1
    SystemMenuBar = 0x2
    SystemScrollBar = 0x3
    SystemGrip = 0x4
    SystemSound = 0x5
    SystemCursor = 0x6
    SystemCaret = 0x7
    SystemAlert = 0x8
    SystemWindow = 0x9
    SystemClient = 0xa
    SystemMenuPopup = 0xb
    SystemMenuItem = 0xc
    SystemToolTip = 0xd
    SystemApplication = 0xe
    SystemDocument = 0xf
    SystemPane = 0x10
    SystemChart = 0x11
    SystemDialog = 0x12
    SystemBorder = 0x13
    SystemGrouping = 0x14
    SystemSeparator = 0x15
    SystemToolbar = 0x16
    SystemStatusBar = 0x17
    SystemTable = 0x18
    SystemColumnHeader = 0x19
    SystemRowHeader = 0x1a
    SystemColumn = 0x1b
    SystemRow = 0x1c
    SystemCell = 0x1d
    SystemLink = 0x1e
    SystemHelpBalloon = 0x1f
    SystemCharacter = 0x20
    SystemList = 0x21
    SystemListItem = 0x22
    SystemOutline = 0x23
    SystemOutlineItem = 0x24
    SystemPageTab = 0x25
    SystemPropertyPage = 0x26
    SystemIndicator = 0x27
    SystemGraphic = 0x28
    SystemStaticText = 0x29
    SystemText = 0x2a
    SystemPushButton = 0x2b
    SystemCheckButton = 0x2c
    SystemRadioButton = 0x2d
    SystemComboBox = 0x2e
    SystemDropList = 0x2f
    SystemProgressBar = 0x30
    SystemDial = 0x31
    SystemHotkeyField = 0x32
    SystemSlider = 0x33
    SystemSpinButton = 0x34
    SystemDiagram = 0x35
    SystemAnimation = 0x36
    SystemEquation = 0x37
    SystemButtonDropDown = 0x38
    SystemButtonMenu = 0x39
    SystemButtonDropDownGrid = 0x3a
    SystemWhiteSpace = 0x3b
    SystemPageTabList = 0x3c
    SystemClock = 0x3d
    SystemSplitButton = 0x3e
    SystemIpAddress = 0x3f
    SystemOutlineButton = 0x40


class AccessibleState():
    SystemNormal = 0
    SystemUnavailable = 0x1
    SystemSelected = 0x2
    SystemFocused = 0x4
    SystemPressed = 0x8
    SystemChecked = 0x10
    SystemMixed = 0x20
    SystemIndeterminate = 0x20
    SystemReadOnly = 0x40
    SystemHotTracked = 0x80
    SystemDefault = 0x100
    SystemExpanded = 0x200
    SystemCollapsed = 0x400
    SystemBusy = 0x800
    SystemFloating = 0x1000
    SystemMarqueed = 0x2000
    SystemAnimated = 0x4000
    SystemInvisible = 0x8000
    SystemOffscreen = 0x10000
    SystemSizeable = 0x20000
    SystemMoveable = 0x40000
    SystemSelfVoicing = 0x80000
    SystemFocusable = 0x100000
    SystemSelectable = 0x200000
    SystemLinked = 0x400000
    SystemTraversed = 0x800000
    SystemMultiSelectable = 0x1000000
    SystemExtSelectable = 0x2000000
    SystemAlertLow = 0x4000000
    SystemAlertMedium = 0x8000000
    SystemAlertHigh = 0x10000000
    SystemProtected = 0x20000000
    SystemValid = 0x7fffffff
    SystemHasPopup = 0x40000000


class AccessibleSelectFlag:
    FlagNone = 0
    FlagTakeFocus = 0x1
    FlagTakeSelection = 0x2
    FlagExtendSelection = 0x4
    FlagAddSelection = 0x8
    FlagRemoveSelection = 0x10
    FlagValid = 0x1f


class RowOrColumnMajor:
    RowMajor = 0
    ColumnMajor = 1
    Indeterminate = 2


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
    _fields_ = [
        ('dwSize', Coord),
        ('dwCursorPosition', Coord),
        ('wAttributes', ctypes.c_uint),
        ('srWindow', SmallRect),
        ('dwMaximumWindowSize', Coord),
    ]


class tagPROCESSENTRY32(ctypes.Structure):
    _fields_ = [
        ('dwSize',              ctypes.wintypes.DWORD),
        ('cntUsage',            ctypes.wintypes.DWORD),
        ('th32ProcessID',       ctypes.wintypes.DWORD),
        ('th32DefaultHeapID',   ctypes.POINTER(ctypes.wintypes.ULONG)),
        ('th32ModuleID',        ctypes.wintypes.DWORD),
        ('cntThreads',          ctypes.wintypes.DWORD),
        ('th32ParentProcessID', ctypes.wintypes.DWORD),
        ('pcPriClassBase',      ctypes.wintypes.LONG),
        ('dwFlags',             ctypes.wintypes.DWORD),
        ('szExeFile',           ctypes.c_wchar * MAX_PATH)
    ]


class MSG(ctypes.Structure):
    _fields_ = [
        ("hwnd",      ctypes.c_uint),
        ("message",   ctypes.c_uint),
        ("wParam",    ctypes.c_uint),
        ("lParam",    ctypes.c_uint),
        ("time",      ctypes.c_uint),
        ("pt",        ctypes.c_ulonglong)
    ]


class MouseEventFlags():
    """This class defines the MouseEventFlags from Win32"""
    Absolute = 0x8000
    LeftDown = 0x0002
    LeftUp = 0x0004
    MiddleDown = 0x0020
    MiddleUp = 0x0040
    Move = 0x0001
    RightDown = 0x0008
    RightUp = 0x0010
    Wheel = 0x0800
    HWheel = 0x1000


class KeyboardEventFlags:
    """This class defines the KeyboardEventFlags from Win32"""
    KeyDown = 0x0000
    ExtendedKey = 0x0001
    KeyUp = 0x0002


class ModifierKey:
    MOD_ALT = 0x0001
    MOD_CONTROL = 0x0002
    MOD_SHIFT = 0x0004
    MOD_WIN = 0x0008
    MOD_NOREPEAT = 0x4000


class Keys:
    """This class defines the Key Code from Win32"""
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
    VK_ENTER = 0x0D
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
    VK_PAGEUP = 0x21
    VK_NEXT = 0x22                          #PAGE DOWN key
    VK_PAGEDOWN = 0x22
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
    """This class defines the values of color for printing on console window"""
    Default = -1
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


class ExpandCollapseState:
    Collapsed = 0
    Expanded = 1
    PartiallyExpanded = 2
    LeafNode = 3


class ToggleState:
    Off = 0
    On = 1
    Indeterminate = 2


class WindowVisualState:
    Normal = 0
    Maximized = 1
    Minimized = 2


class ShowWindow:
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


class SWP:
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


class MB:
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


class Win32API:
    """Some native methods for python calling"""
    StdOutputHandle = -11
    ConsoleOutputHandle = None
    DefaultColor = None
    SpecialKeyDict = {
        'LBUTTON' : Keys.VK_LBUTTON,                        #Left mouse button
        'RBUTTON' : Keys.VK_RBUTTON,                        #Right mouse button
        'CANCEL' : Keys.VK_CANCEL,                          #Control-break processing
        'MBUTTON' : Keys.VK_MBUTTON,                        #Middle mouse button (three-button mouse)
        'XBUTTON1' : Keys.VK_XBUTTON1,                      #X1 mouse button
        'XBUTTON2' : Keys.VK_XBUTTON2,                      #X2 mouse button
        'BACK' : Keys.VK_BACK,                              #BACKSPACE key
        'TAB' : Keys.VK_TAB,                                #TAB key
        'CLEAR' : Keys.VK_CLEAR,                            #CLEAR key
        'RETURN' : Keys.VK_RETURN,                          #ENTER key
        'ENTER' : Keys.VK_RETURN,                           #ENTER key
        'SHIFT' : Keys.VK_SHIFT,                            #SHIFT key
        'CTRL' : Keys.VK_CONTROL,                           #CTRL key
        'CONTROL' : Keys.VK_CONTROL,                        #CTRL key
        'ALT' : Keys.VK_MENU,                               #ALT key
        'PAUSE' : Keys.VK_PAUSE,                            #PAUSE key
        'CAPITAL' : Keys.VK_CAPITAL,                        #CAPS LOCK key
        'KANA' : Keys.VK_KANA,                              #IME Kana mode
        'HANGUEL' : Keys.VK_HANGUEL,                        #IME Hanguel mode (maintained for compatibility; use VK_HANGUL)
        'HANGUL' : Keys.VK_HANGUL,                          #IME Hangul mode
        'JUNJA' : Keys.VK_JUNJA,                            #IME Junja mode
        'FINAL' : Keys.VK_FINAL,                            #IME final mode
        'HANJA' : Keys.VK_HANJA,                            #IME Hanja mode
        'KANJI' : Keys.VK_KANJI,                            #IME Kanji mode
        'ESC' : Keys.VK_ESCAPE,                             #ESC key
        'ESCAPE' : Keys.VK_ESCAPE,                          #ESC key
        'CONVERT' : Keys.VK_CONVERT,                        #IME convert
        'NONCONVERT' : Keys.VK_NONCONVERT,                  #IME nonconvert
        'ACCEPT' : Keys.VK_ACCEPT,                          #IME accept
        'MODECHANGE' : Keys.VK_MODECHANGE,                  #IME mode change request
        'SPACE' : Keys.VK_SPACE,                            #SPACEBAR
        'PRIOR' : Keys.VK_PRIOR,                            #PAGE UP key
        'PAGEUP' : Keys.VK_PRIOR,                           #PAGE UP key
        'NEXT' : Keys.VK_NEXT,                              #PAGE DOWN key
        'PAGEDOWN': Keys.VK_NEXT,                           #PAGE DOWN key
        'END' : Keys.VK_END,                                #END key
        'HOME' : Keys.VK_HOME,                              #HOME key
        'LEFT' : Keys.VK_LEFT,                              #LEFT ARROW key
        'UP' : Keys.VK_UP,                                  #UP ARROW key
        'RIGHT' : Keys.VK_RIGHT,                            #RIGHT ARROW key
        'DOWN' : Keys.VK_DOWN,                              #DOWN ARROW key
        'SELECT' : Keys.VK_SELECT,                          #SELECT key
        'PRINT' : Keys.VK_PRINT,                            #PRINT key
        'EXECUTE' : Keys.VK_EXECUTE,                        #EXECUTE key
        'SNAPSHOT' : Keys.VK_SNAPSHOT,                      #PRINT SCREEN key
        'PRINTSCREEN': Keys.VK_SNAPSHOT,                    #PRINT SCREEN key
        'INSERT' : Keys.VK_INSERT,                          #INS key
        'INS' : Keys.VK_INSERT,                             #INS key
        'DELETE' : Keys.VK_DELETE,                          #DEL key
        'DEL' : Keys.VK_DELETE,                             #DEL key
        'HELP' : Keys.VK_HELP,                              #HELP key
        'WIN' : Keys.VK_LWIN,                               #Left Windows key (Natural keyboard)
        'LWIN' : Keys.VK_LWIN,                              #Left Windows key (Natural keyboard)
        'RWIN' : Keys.VK_RWIN,                              #Right Windows key (Natural keyboard)
        'APPS' : Keys.VK_APPS,                              #Applications key (Natural keyboard)
        'SLEEP' : Keys.VK_SLEEP,                            #Computer Sleep key
        'NUMPAD0' : Keys.VK_NUMPAD0,                        #Numeric keypad 0 key
        'NUMPAD1' : Keys.VK_NUMPAD1,                        #Numeric keypad 1 key
        'NUMPAD2' : Keys.VK_NUMPAD2,                        #Numeric keypad 2 key
        'NUMPAD3' : Keys.VK_NUMPAD3,                        #Numeric keypad 3 key
        'NUMPAD4' : Keys.VK_NUMPAD4,                        #Numeric keypad 4 key
        'NUMPAD5' : Keys.VK_NUMPAD5,                        #Numeric keypad 5 key
        'NUMPAD6' : Keys.VK_NUMPAD6,                        #Numeric keypad 6 key
        'NUMPAD7' : Keys.VK_NUMPAD7,                        #Numeric keypad 7 key
        'NUMPAD8' : Keys.VK_NUMPAD8,                        #Numeric keypad 8 key
        'NUMPAD9' : Keys.VK_NUMPAD9,                        #Numeric keypad 9 key
        'MULTIPLY' : Keys.VK_MULTIPLY,                      #Multiply key
        'ADD' : Keys.VK_ADD,                                #Add key
        'SEPARATOR' : Keys.VK_SEPARATOR,                    #Separator key
        'SUBTRACT' : Keys.VK_SUBTRACT,                      #Subtract key
        'DECIMAL' : Keys.VK_DECIMAL,                        #Decimal key
        'DIVIDE' : Keys.VK_DIVIDE,                          #Divide key
        'F1' : Keys.VK_F1,                                  #F1 key
        'F2' : Keys.VK_F2,                                  #F2 key
        'F3' : Keys.VK_F3,                                  #F3 key
        'F4' : Keys.VK_F4,                                  #F4 key
        'F5' : Keys.VK_F5,                                  #F5 key
        'F6' : Keys.VK_F6,                                  #F6 key
        'F7' : Keys.VK_F7,                                  #F7 key
        'F8' : Keys.VK_F8,                                  #F8 key
        'F9' : Keys.VK_F9,                                  #F9 key
        'F10' : Keys.VK_F10,                                #F10 key
        'F11' : Keys.VK_F11,                                #F11 key
        'F12' : Keys.VK_F12,                                #F12 key
        'F13' : Keys.VK_F13,                                #F13 key
        'F14' : Keys.VK_F14,                                #F14 key
        'F15' : Keys.VK_F15,                                #F15 key
        'F16' : Keys.VK_F16,                                #F16 key
        'F17' : Keys.VK_F17,                                #F17 key
        'F18' : Keys.VK_F18,                                #F18 key
        'F19' : Keys.VK_F19,                                #F19 key
        'F20' : Keys.VK_F20,                                #F20 key
        'F21' : Keys.VK_F21,                                #F21 key
        'F22' : Keys.VK_F22,                                #F22 key
        'F23' : Keys.VK_F23,                                #F23 key
        'F24' : Keys.VK_F24,                                #F24 key
        'NUMLOCK' : Keys.VK_NUMLOCK,                        #NUM LOCK key
        'SCROLL' : Keys.VK_SCROLL,                          #SCROLL LOCK key
        'LSHIFT' : Keys.VK_LSHIFT,                          #Left SHIFT key
        'RSHIFT' : Keys.VK_RSHIFT,                          #Right SHIFT key
        'LCONTROL' : Keys.VK_LCONTROL,                      #Left CONTROL key
        'LCTRL' : Keys.VK_LCONTROL,                         #Left CONTROL key
        'RCONTROL' : Keys.VK_RCONTROL,                      #Right CONTROL key
        'RCTRL' : Keys.VK_RCONTROL,                         #Right CONTROL key
        'LALT' : Keys.VK_LMENU,                             #Left MENU key
        'RALT' : Keys.VK_RMENU,                             #Right MENU key
        'BROWSER_BACK' : Keys.VK_BROWSER_BACK,              #Browser Back key
        'BROWSER_FORWARD' : Keys.VK_BROWSER_FORWARD,        #Browser Forward key
        'BROWSER_REFRESH' : Keys.VK_BROWSER_REFRESH,        #Browser Refresh key
        'BROWSER_STOP' : Keys.VK_BROWSER_STOP,              #Browser Stop key
        'BROWSER_SEARCH' : Keys.VK_BROWSER_SEARCH,          #Browser Search key
        'BROWSER_FAVORITES' : Keys.VK_BROWSER_FAVORITES,    #Browser Favorites key
        'BROWSER_HOME' : Keys.VK_BROWSER_HOME,              #Browser Start and Home key
        'VOLUME_MUTE' : Keys.VK_VOLUME_MUTE,                #Volume Mute key
        'VOLUME_DOWN' : Keys.VK_VOLUME_DOWN,                #Volume Down key
        'VOLUME_UP' : Keys.VK_VOLUME_UP,                    #Volume Up key
        'MEDIA_NEXT_TRACK' : Keys.VK_MEDIA_NEXT_TRACK,      #Next Track key
        'MEDIA_PREV_TRACK' : Keys.VK_MEDIA_PREV_TRACK,      #Previous Track key
        'MEDIA_STOP' : Keys.VK_MEDIA_STOP,                  #Stop Media key
        'MEDIA_PLAY_PAUSE' : Keys.VK_MEDIA_PLAY_PAUSE,      #Play/Pause Media key
        'LAUNCH_MAIL' : Keys.VK_LAUNCH_MAIL,                #Start Mail key
        'LAUNCH_MEDIA_SELECT' : Keys.VK_LAUNCH_MEDIA_SELECT,#Select Media key
        'LAUNCH_APP1' : Keys.VK_LAUNCH_APP1,                #Start Application 1 key
        'LAUNCH_APP2' : Keys.VK_LAUNCH_APP2,                #Start Application 2 key
        'OEM_1' : Keys.VK_OEM_1,                            #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the ';:' key
        'OEM_PLUS' : Keys.VK_OEM_PLUS,                      #For any country/region, the '+' key
        'OEM_COMMA' : Keys.VK_OEM_COMMA,                    #For any country/region, the ',' key
        'OEM_MINUS' : Keys.VK_OEM_MINUS,                    #For any country/region, the '-' key
        'OEM_PERIOD' : Keys.VK_OEM_PERIOD,                  #For any country/region, the '.' key
        'OEM_2' : Keys.VK_OEM_2,                            #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the '/?' key
        'OEM_3' : Keys.VK_OEM_3,                            #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the '`~' key
        'OEM_4' : Keys.VK_OEM_4,                            #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the '[{' key
        'OEM_5' : Keys.VK_OEM_5,                            #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the '\|' key
        'OEM_6' : Keys.VK_OEM_6,                            #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the ']}' key
        'OEM_7' : Keys.VK_OEM_7,                            #Used for miscellaneous characters; it can vary by keyboard.For the US standard keyboard, the 'single-quote/double-quote' key
        'OEM_8' : Keys.VK_OEM_8,                            #Used for miscellaneous characters; it can vary by keyboard.
        'OEM_102' : Keys.VK_OEM_102,                        #Either the angle bracket key or the backslash key on the RT 102-key keyboard
        'PROCESSKEY' : Keys.VK_PROCESSKEY,                  #IME PROCESS key
        'PACKET' : Keys.VK_PACKET,                          #Used to pass Unicode characters as if they were keystrokes. The VK_PACKET key is the low word of a 32-bit Virtual Key value used for non-keyboard input methods. For more information, see Remark in KEYBDINPUT, SendInput, WM_KEYDOWN, and WM_KeyUp
        'ATTN' : Keys.VK_ATTN,                              #Attn key
        'CRSEL' : Keys.VK_CRSEL,                            #CrSel key
        'EXSEL' : Keys.VK_EXSEL,                            #ExSel key
        'EREOF' : Keys.VK_EREOF,                            #Erase EOF key
        'PLAY' : Keys.VK_PLAY,                              #Play key
        'ZOOM' : Keys.VK_ZOOM,                              #Zoom key
        'NONAME' : Keys.VK_NONAME,                          #Reserved
        'PA1' : Keys.VK_PA1,                                #PA1 key
        'OEM_CLEAR' : Keys.VK_OEM_CLEAR,                    #Clear key
        }
    CharacterDict = {
        '0' : Keys.VK_0,                             #0 key
        '1' : Keys.VK_1,                             #1 key
        '2' : Keys.VK_2,                             #2 key
        '3' : Keys.VK_3,                             #3 key
        '4' : Keys.VK_4,                             #4 key
        '5' : Keys.VK_5,                             #5 key
        '6' : Keys.VK_6,                             #6 key
        '7' : Keys.VK_7,                             #7 key
        '8' : Keys.VK_8,                             #8 key
        '9' : Keys.VK_9,                             #9 key
        'a' : Keys.VK_A,                             #A key
        'A' : Keys.VK_A,                             #A key
        'b' : Keys.VK_B,                             #B key
        'B' : Keys.VK_B,                             #B key
        'c' : Keys.VK_C,                             #C key
        'C' : Keys.VK_C,                             #C key
        'd' : Keys.VK_D,                             #D key
        'D' : Keys.VK_D,                             #D key
        'e' : Keys.VK_E,                             #E key
        'E' : Keys.VK_E,                             #E key
        'f' : Keys.VK_F,                             #F key
        'F' : Keys.VK_F,                             #F key
        'g' : Keys.VK_G,                             #G key
        'G' : Keys.VK_G,                             #G key
        'h' : Keys.VK_H,                             #H key
        'H' : Keys.VK_H,                             #H key
        'i' : Keys.VK_I,                             #I key
        'I' : Keys.VK_I,                             #I key
        'j' : Keys.VK_J,                             #J key
        'J' : Keys.VK_J,                             #J key
        'k' : Keys.VK_K,                             #K key
        'K' : Keys.VK_K,                             #K key
        'l' : Keys.VK_L,                             #L key
        'L' : Keys.VK_L,                             #L key
        'm' : Keys.VK_M,                             #M key
        'M' : Keys.VK_M,                             #M key
        'n' : Keys.VK_N,                             #N key
        'N' : Keys.VK_N,                             #N key
        'o' : Keys.VK_O,                             #O key
        'O' : Keys.VK_O,                             #O key
        'p' : Keys.VK_P,                             #P key
        'P' : Keys.VK_P,                             #P key
        'q' : Keys.VK_Q,                             #Q key
        'Q' : Keys.VK_Q,                             #Q key
        'r' : Keys.VK_R,                             #R key
        'R' : Keys.VK_R,                             #R key
        's' : Keys.VK_S,                             #S key
        'S' : Keys.VK_S,                             #S key
        't' : Keys.VK_T,                             #T key
        'T' : Keys.VK_T,                             #T key
        'u' : Keys.VK_U,                             #U key
        'U' : Keys.VK_U,                             #U key
        'v' : Keys.VK_V,                             #V key
        'V' : Keys.VK_V,                             #V key
        'w' : Keys.VK_W,                             #W key
        'W' : Keys.VK_W,                             #W key
        'x' : Keys.VK_X,                             #X key
        'X' : Keys.VK_X,                             #X key
        'y' : Keys.VK_Y,                             #Y key
        'Y' : Keys.VK_Y,                             #Y key
        'z' : Keys.VK_Z,                             #Z key
        'Z' : Keys.VK_Z,                             #Z key
        ' ' : Keys.VK_SPACE,                         #Space key
        '`' : Keys.VK_OEM_3,                         #` key
        #'~' : Keys.VK_OEM_3,                         #~ key
        '-' : Keys.VK_OEM_MINUS,                     #- key
        #'_' : Keys.VK_OEM_MINUS,                     #_ key
        '=' : Keys.VK_OEM_PLUS,                      #= key
        #'+' : Keys.VK_OEM_PLUS,                      #+ key
        '[' : Keys.VK_OEM_4,                         #[ key
        #'{' : Keys.VK_OEM_4,                         #{ key
        ']' : Keys.VK_OEM_6,                         #] key
        #'}' : Keys.VK_OEM_6,                         #} key
        '\\' : Keys.VK_OEM_5,                        #\ key
        #'|' : Keys.VK_OEM_5,                         #| key
        ';' : Keys.VK_OEM_1,                         #; key
        #':' : Keys.VK_OEM_1,                         #: key
        '\'' : Keys.VK_OEM_7,                        #' key
        #'"' : Keys.VK_OEM_7,                         #" key
        ',' : Keys.VK_OEM_COMMA,                     #, key
        #'<' : Keys.VK_OEM_COMMA,                     #< key
        '.' : Keys.VK_OEM_PERIOD,                    #. key
        #'>' : Keys.VK_OEM_PERIOD,                    #> key
        '/' : Keys.VK_OEM_2,                         #/ key
        #'?' : Keys.VK_OEM_2,                         #? key
        }

    @staticmethod
    def GetClipboardText():
        if ctypes.windll.user32.OpenClipboard(0):
            if ctypes.windll.user32.IsClipboardFormatAvailable(13): # CF_TEXT=1, CF_UNICODETEXT=13
                hClipboardData = ctypes.windll.user32.GetClipboardData(13)
                hText = ctypes.windll.kernel32.GlobalLock(hClipboardData);
                text = ctypes.c_wchar_p(hText).value[:]
                ctypes.windll.kernel32.GlobalUnlock(hClipboardData);
                ctypes.windll.user32.CloseClipboard()
                return text
        return ''

    @staticmethod
    def SetClipboardText(text):
        if ctypes.windll.user32.OpenClipboard(0):
            ctypes.windll.user32.EmptyClipboard()
            textLen = (len(text) + 1) * 2
            hClipboardData = ctypes.windll.kernel32.GlobalAlloc(0, textLen)  # GMEM_FIXED=0
            hDestText = ctypes.windll.kernel32.GlobalLock(hClipboardData)
            _AutomationClient.instance().dll.WcsCpy(ctypes.c_wchar_p(hDestText), textLen, ctypes.c_wchar_p(text))
            ctypes.windll.kernel32.GlobalUnlock(hClipboardData)
            ctypes.windll.user32.SetClipboardData(13, ctypes.c_void_p(hClipboardData)) # CF_TEXT=1, CF_UNICODETEXT=13
            ctypes.windll.user32.CloseClipboard()

    @staticmethod
    def SetConsoleColor(color):
        """Change the text color on console window"""
        if not Win32API.DefaultColor:
            if not Win32API.ConsoleOutputHandle:
                Win32API.ConsoleOutputHandle = ctypes.windll.kernel32.GetStdHandle(Win32API.StdOutputHandle)
            bufferInfo = ConsoleScreenBufferInfo()
            ctypes.windll.kernel32.GetConsoleScreenBufferInfo(Win32API.ConsoleOutputHandle, ctypes.byref(bufferInfo))
            Win32API.DefaultColor = int(bufferInfo.wAttributes & 0xFF)
        if IsPy3:
            if sys.stdout:
                sys.stdout.flush() # need flush stdout in python 3
        ctypes.windll.kernel32.SetConsoleTextAttribute(Win32API.ConsoleOutputHandle, color)

    @staticmethod
    def ResetConsoleColor():
        """Reset the default text color on console window"""
        if IsPy3:
            if sys.stdout:
                sys.stdout.flush() # need flush stdout in python 3
        ctypes.windll.kernel32.SetConsoleTextAttribute(Win32API.ConsoleOutputHandle, Win32API.DefaultColor)

    @staticmethod
    def WindowFromPoint(x, y):
        """Return hwnd"""
        point = Point()
        point.x = x
        point.y = y
        return ctypes.windll.user32.WindowFromPoint(point)

    @staticmethod
    def GetCursorPos():
        """Return tuple (x, y)"""
        point = Point()
        ctypes.windll.user32.GetCursorPos(ctypes.byref(point))
        return int(point.x), int(point.y)

    @staticmethod
    def SetCursorPos(x, y):
        """Set cursor to point x, y"""
        ctypes.windll.user32.SetCursorPos(x, y)

    @staticmethod
    def GetDoubleClickTime():
        """Get the double click time of mouse"""
        return ctypes.windll.user32.GetDoubleClickTime()

    @staticmethod
    def mouse_event(dwFlags, dx, dy, dwData, dwExtraInfo):
        """Call API mouse_event from user32.dll"""
        ctypes.windll.user32.mouse_event(dwFlags, dx, dy, dwData, dwExtraInfo)

    @staticmethod
    def keybd_event(bVk, bScan, dwFlags, dwExtraInfo):
        """Call API keybd_event from user32.dll"""
        ctypes.windll.user32.keybd_event(bVk, bScan, dwFlags, dwExtraInfo)

    @staticmethod
    def PostMessage(handle, msg, wparam, lparam):
        """Call API PostMessageW from user32.dll"""
        return ctypes.windll.user32.PostMessageW(handle, msg, wparam, lparam)

    @staticmethod
    def SendMessage(handle, msg, wparam, lparam):
        """Call API SendMessageW from user32.dll"""
        return ctypes.windll.user32.SendMessageW(handle, msg, wparam, lparam)

    @staticmethod
    def MouseClick(x, y, waitTime = OPERATION_WAIT_TIME):
        """
        Simulate mouse click at point x, y
        x and y must be integer
        """
        Win32API.SetCursorPos(x, y)
        Win32API.mouse_event(MouseEventFlags.LeftDown | MouseEventFlags.Absolute, x, y, 0, 0)
        time.sleep(0.05)
        Win32API.mouse_event(MouseEventFlags.LeftUp | MouseEventFlags.Absolute, x, y, 0, 0)
        time.sleep(waitTime)

    @staticmethod
    def MouseMiddleClick(x, y, waitTime = OPERATION_WAIT_TIME):
        """
        Simulate mouse middle click at point x, y
        x and y must be integer
        """
        Win32API.SetCursorPos(x, y)
        Win32API.mouse_event(MouseEventFlags.MiddleDown | MouseEventFlags.Absolute, x, y, 0, 0)
        time.sleep(0.05)
        Win32API.mouse_event(MouseEventFlags.MiddleUp | MouseEventFlags.Absolute, x, y, 0, 0)
        time.sleep(waitTime)

    @staticmethod
    def MouseRightClick(x, y, waitTime = OPERATION_WAIT_TIME):
        """
        Simulate mouse right click at point x, y
        x and y must be integer
        """
        Win32API.SetCursorPos(x, y)
        Win32API.mouse_event(MouseEventFlags.RightDown | MouseEventFlags.Absolute, x, y, 0, 0)
        time.sleep(0.05)
        Win32API.mouse_event(MouseEventFlags.RightUp | MouseEventFlags.Absolute, x, y, 0, 0)
        time.sleep(waitTime)

    @staticmethod
    def MouseMoveTo(x, y, moveSpeed = 1, waitTime = OPERATION_WAIT_TIME):
        """
        Simulate mouse move to point x, y from current cursor
        x and y must be integer
        moveSpeed: double, 1 normal speed, < 1 move slower, > 1 move faster
        """
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
        time.sleep(waitTime)

    #@staticmethod
    #def MouseDragTo(x1, y1, x2, y2, moveSpeed = 1, waitTime = OPERATION_WAIT_TIME):
    #use MouseDragDrop instead

    @staticmethod
    def MouseDragDrop(x1, y1, x2, y2, moveSpeed = 1, waitTime = OPERATION_WAIT_TIME):
        """
        Simulate mouse drag from point x1, y1 drop to point x2, y2
        x1, y1, x2, y2, must be integer
        moveSpeed: double, 1 normal speed, < 1 move slower, > 1 move faster
        """
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
            moveTime = moveTime * maxPoint * 1.0 / maxSide
        stepCount = maxPoint // 20
        Win32API.SetCursorPos(x1, y1)
        Win32API.mouse_event(MouseEventFlags.LeftDown| MouseEventFlags.Absolute, x1*65536//screenWidth, y1*65536//screenHeight, 0, 0)
        if stepCount > 1:
            xStep = (x2 - x1) * 1.0 / stepCount
            yStep = (y2 - y1) * 1.0 / stepCount
            interval = moveTime / stepCount
            for i in range(stepCount):
                time.sleep(interval)
                x1 += xStep
                y1 += yStep
                Win32API.SetCursorPos(int(x1), int(y1))
        else:
            time.sleep(0.05)
        Win32API.mouse_event(MouseEventFlags.Absolute | MouseEventFlags.LeftUp, x2*65536//screenWidth, y2*65536//screenHeight, 0, 0)
        Win32API.SetCursorPos(x2, y2)
        time.sleep(waitTime)

    @staticmethod
    def MouseWheelDown(wheelTimes = 1, waitTime = OPERATION_WAIT_TIME):
        """
        Simulate mouse wheel down
        wheelTimes, int value
        """
        for i in range(wheelTimes):
            Win32API.mouse_event(MouseEventFlags.Wheel, 0, 0, -120, 0)
            time.sleep(0.1)
        time.sleep(waitTime)

    @staticmethod
    def MouseWheelUp(wheelTimes = 1, waitTime = OPERATION_WAIT_TIME):
        """
        Simulate mouse wheel up
        wheelTimes, int value
        """
        for i in range(wheelTimes):
            Win32API.mouse_event(MouseEventFlags.Wheel, 0, 0, 120, 0)
            time.sleep(0.1)
        time.sleep(waitTime)

    @staticmethod
    def GetScreenSize():
        """Return tuple (width, height)"""
        SM_CXSCREEN = 0
        SM_CYSCREEN = 1
        w = ctypes.windll.user32.GetSystemMetrics(SM_CXSCREEN)
        h = ctypes.windll.user32.GetSystemMetrics(SM_CYSCREEN)
        return w, h

    @staticmethod
    def GetPixelColor(x, y, handle = 0):
        """
        If handle is 0, get pixel from desktop, return bgr
        r = bgr & 0x0000FF
        g = (bgr & 0x00FF00) >> 8
        b = (bgr & 0xFF0000) >> 16

        Not all devices support GetPixel. An application should call GetDeviceCaps to determine whether a specified device supports this function. Console window doesn't support.
        """
        hdc = ctypes.windll.user32.GetWindowDC(handle)
        bgr = ctypes.windll.gdi32.GetPixel(hdc, x, y)
        ctypes.windll.user32.ReleaseDC(handle, hdc)
        return bgr

    @staticmethod
    def MessageBox(content, title, flags = MB.OK):
        """Call API MessageBox from user32.dll"""
        return ctypes.windll.user32.MessageBoxW(0, ctypes.c_wchar_p(content), ctypes.c_wchar_p(title), flags)

    @staticmethod
    def SetForegroundWindow(hWnd):
        """
        Set a window to foreground
        hWnd: integer, handle of a Win32 window
        """
        return ctypes.windll.user32.SetForegroundWindow(hWnd)

    @staticmethod
    def SetWindowTopmost(hWnd, isTopmost):
        """
        Set a window to Topmost
        hWnd: integer, handle of a Win32 window
        isTopmost: bool, be topmost or not
        """
        topValue = SWP.HWND_TOPMOST if isTopmost else SWP.HWND_NOTOPMOST
        return Win32API.SetWindowPos(hWnd, topValue, 0, 0, 0, 0, SWP.SWP_NOSIZE|SWP.SWP_NOMOVE)

    @staticmethod
    def SwitchToThisWindow(hWnd):
        """Call API SwitchToThisWindow from user32.dll"""
        return ctypes.windll.user32.SwitchToThisWindow(hWnd, 1)

    @staticmethod
    def ShowWindow(hWnd, cmdShow):
        """ShowWindow(hWnd, ShowWindow.Show), see values in class ShowWindow"""
        return ctypes.windll.user32.ShowWindow(hWnd, cmdShow)

    @staticmethod
    def MoveWindow(hWnd, x, y, width, height, repaint = 1):
        """Call API MoveWindow from user32.dll"""
        return ctypes.windll.user32.MoveWindow(hWnd, x, y, width, height, repaint)

    @staticmethod
    def SetWindowPos(hWnd, hWndInsertAfter, x, y, width, height, flags):
        """Call API SetWindowPos from user32.dll, flags see class SWP"""
        return ctypes.windll.user32.SetWindowPos(hWnd, hWndInsertAfter, x, y, width, height, flags)

    @staticmethod
    def GetWindowText(hWnd):
        """Get window text"""
        wArray = ctypes.c_wchar * MAX_PATH
        values = wArray()
        ctypes.windll.user32.GetWindowTextW(hWnd, values, MAX_PATH)
        return values.value

    @staticmethod
    def GetEditText(hWnd):
        """Get text of Win32 Edit"""
        textLen = Win32API.SendMessage(hWnd, 0x000E, 0, 0) + 1  #WM_GETTEXTLENGTH
        wArray = ctypes.c_wchar * textLen
        values = wArray()
        Win32API.SendMessage(hWnd, 0x000D, textLen, values)  #WM_GETTEXT
        return values.value

    @staticmethod
    def SetWindowText(hWnd, text):
        """Set window text"""
        return ctypes.windll.user32.SetWindowTextW(hWnd, ctypes.c_wchar_p(text))

    @staticmethod
    def GetConsoleOriginalTitle():
        """GetConsoleOriginalTitle"""
        wArray = ctypes.c_wchar * MAX_PATH
        values = wArray()
        ctypes.windll.kernel32.GetConsoleOriginalTitleW(values, MAX_PATH)
        return values.value

    @staticmethod
    def GetConsoleTitle():
        """GetConsoleTitle"""
        wArray = ctypes.c_wchar * MAX_PATH
        values = wArray()
        ctypes.windll.kernel32.GetConsoleTitleW(values, MAX_PATH)
        return values.value

    @staticmethod
    def SetConsoleTitle(text):
        """SetConsoleTitle"""
        return ctypes.windll.kernel32.SetConsoleTitleW(ctypes.c_wchar_p(text))

    @staticmethod
    def GetForegroundWindow():
        return ctypes.windll.user32.GetForegroundWindow()

    @staticmethod
    def IsDesktopLocked():
        """desktop is locked if press Win+L, Ctrl+Alt+Del or in remote desktop mode"""
        isLocked = False
        desk = ctypes.windll.user32.OpenDesktopW(ctypes.c_wchar_p('Default'), 0, 0, 0x0100)  #DESKTOP_SWITCHDESKTOP = 0x0100
        if desk:
            isLocked = not ctypes.windll.user32.SwitchDesktop(desk)
            ctypes.windll.user32.CloseDesktop(desk)
        return isLocked

    @staticmethod
    def PlayWaveFile(filePath = r'C:\Windows\Media\notify.wav', isAsync = True):
        """play wave file"""
        SND_ASYNC = 0x0001
        SND_NODEFAULT = 0x0002
        flag = SND_NODEFAULT
        if isAsync:
            flag |= SND_ASYNC
        ctypes.windll.winmm.sndPlaySoundW(ctypes.c_wchar_p(filePath), flag)

    @staticmethod
    def GetProcessCommandLine(processId):
        """may not work on Windows 7 or higher bacause of security policy"""
        wArray = ctypes.c_wchar * MAX_PATH
        values = wArray()
        _AutomationClient.instance().dll.GetProcessCommandLine(processId, values, MAX_PATH)
        return values.value

    @staticmethod
    def GetParentProcessId(processId = -1):
        return _AutomationClient.instance().dll.GetParentProcessId(processId)

    @staticmethod
    def IsProcess64Bit(processId):
        """return True if process is 64 bit
            return False if process is 32 bit
            return None if unknown, maybe caused by having no acess right to the process"""
        try:
            func = ctypes.windll.ntdll.ZwWow64ReadVirtualMemory64  #only 64 bit OS has this function
        except Exception as ex:
            return False
        try:
            IsWow64Process = ctypes.windll.kernel32.IsWow64Process
        except Exception as ex:
            return False
        hProcess = ctypes.windll.kernel32.OpenProcess(0x0400, 0, processId);  #PROCESS_QUERY_INFORMATION=0x0400
        if hProcess:
            is64Bit = ctypes.c_int32()
            if IsWow64Process(hProcess, ctypes.byref(is64Bit)):
                ctypes.windll.kernel32.CloseHandle(hProcess)
                return False if is64Bit.value else True
            else:
                ctypes.windll.kernel32.CloseHandle(hProcess)

    @staticmethod
    def TerminateProcess(processId):
        """Terminate process by process id"""
        hProcess = ctypes.windll.kernel32.OpenProcess(0x0001, 0, processId);  #PROCESS_TERMINATE=0x0001
        if hProcess:
            ret = ctypes.windll.kernel32.TerminateProcess(hProcess, -1)
            ctypes.windll.kernel32.CloseHandle(hProcess)
            return ret

    @staticmethod
    def TerminateProcessByName(processName):
        """Terminate process by process name, example: TerminateProcessByName('notepad.exe')"""
        for pid, name in Win32API.EnumProcess():
            if name.lower() == processName.lower():
                Win32API.TerminateProcess(pid)

    @staticmethod
    def EnumProcess():
        """Return a list of namedtuple (pid, name), see TerminateProcessByName"""
        import collections
        hSnapshot = ctypes.windll.kernel32.CreateToolhelp32Snapshot(15, 0)  #TH32CS_SNAPALL is 15
        processEntry32 = tagPROCESSENTRY32()
        processClass = collections.namedtuple('processInfo', 'pid name')
        processEntry32.dwSize = ctypes.sizeof(processEntry32)
        processList = []
        continueFind = ctypes.windll.kernel32.Process32FirstW(ctypes.c_void_p(hSnapshot), ctypes.byref(processEntry32))
        while continueFind:
            pid = processEntry32.th32ProcessID
            name = (processEntry32.szExeFile)
            processList.append(processClass(pid, name))
            continueFind = ctypes.windll.kernel32.Process32NextW(ctypes.c_void_p(hSnapshot), ctypes.byref(processEntry32))
        ctypes.windll.kernel32.CloseHandle(hSnapshot)
        return processList

    @staticmethod
    def SendKey(key, waitTime = OPERATION_WAIT_TIME):
        """
        Simulate typing a key
        key: a value in class Keys
        """
        Win32API.keybd_event(key, 0, KeyboardEventFlags.KeyDown | KeyboardEventFlags.ExtendedKey, 0)
        Win32API.keybd_event(key, 0, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey, 0)
        time.sleep(waitTime)

    #@staticmethod
    #def SendWait(keys):
        #"""this method needs .Net and PythonForNet, will call System.Windows.Forms.SendKeys.SendWait"""
        #import clr
        #import System.Windows.Forms
        #System.Windows.Forms.SendKeys.SendWait(keys)

    @staticmethod
    def KeyDown(key, waitTime = OPERATION_WAIT_TIME):
        """
        Simulate a key down for key
        key: a value in class Keys
        """
        Win32API.keybd_event(key, 0, KeyboardEventFlags.KeyDown | KeyboardEventFlags.ExtendedKey, 0)
        time.sleep(waitTime)

    @staticmethod
    def KeyUp(key, waitTime = OPERATION_WAIT_TIME):
        """
        Simulate a key up for key
        key: a value in class Keys
        """
        Win32API.keybd_event(key, 0, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey, 0)
        time.sleep(waitTime)

    @staticmethod
    def PressKey(key, waitTime = OPERATION_WAIT_TIME):
        """Same as KeyDown"""
        Win32API.KeyDown(key, waitTime)

    @staticmethod
    def ReleaseKey(key, waitTime = OPERATION_WAIT_TIME):
        """Same as KeyUp"""
        Win32API.KeyUp(key, waitTime)

    @staticmethod
    def IsKeyPressed(key):
        """
        Check key is pressed or not
        key: a value in class Keys
        """
        state = ctypes.windll.user32.GetAsyncKeyState(key)
        return True if state & 0x8000 else False

    @staticmethod
    def VKtoSC(key):
        """key: a value in class Keys"""
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
    def SendKeys(text, interval = 0.01, waitTime = OPERATION_WAIT_TIME, debug = False):
        """
        Simulate typing keys on keyboard
        text: str, keys to type
        interval: double, seconds between keys
        debug: bool, if True, print the Keys
        example:
        {Ctrl}, {Delete} ... are special keys' name in Win32API.SpecialKeyDict
        SendKeys('{Ctrl}a{Delete}{Ctrl}v{Ctrl}s{Ctrl}{Shift}s{Win}e{PageDown}') #press Ctrl+a, Delete, Ctrl+v, Ctrl+s, Ctrl+Shift+s, Win+e, PageDown
        SendKeys('{Ctrl}(AB)({Shift}(123))') #press Ctrl+A+B, type (, press Shift+1+2+3, type ), if () follows a hold key, hold key won't release util )
        SendKeys('{Ctrl}{a 3}') #press Ctrl+a at the same time, release Ctrl+a, then type a 2 times
        SendKeys('{a 3}{B 5}') #type a 3 times, type B 5 times
        SendKeys('{{}Hello{}}abc {a}{b}{c} test{} 3}{!}{a} (){(}{)}') #type: {Hello}abc abc test}}}!a ()()
        SendKeys('0123456789{Enter}')
        SendKeys('ABCDEFGHIJKLMNOPQRSTUVWXYZ{Enter}')
        SendKeys('abcdefghijklmnopqrstuvwxyz{Enter}')
        SendKeys('`~!@#$%^&*()-_=+{Enter}')
        SendKeys('[]{{}{}}\\|;:\'\",<.>/?{Enter}')
        """
        holdKeys = ('WIN', 'LWIN', 'RWIN', 'SHIFT', 'LSHIFT', 'RSHIFT', 'CTRL', 'CONTROL', 'LCTRL', 'RCTRL', 'LCONTROL', 'LCONTROL', 'ALT', 'LALT', 'RALT')
        keys = []
        printKeys = []
        i = 0
        insertIndex = 0
        length = len(text)
        hold = False
        include = False
        #lastKey = ''
        lastKeyValue = None
        while True:
            if text[i] == '{':
                rindex = text.find('}', i)
                if rindex == i+1:#{}}
                    rindex = text.find('}', i+2)
                if rindex == -1:
                    raise ValueError('"{" or "{}" is not valid, use "{{}" for "{", use "{}}" for "}"')
                key = text[i+1:rindex]
                key = [it for it in key.split(' ') if it]
                if not key:
                    raise ValueError('"{}" is not valid, use "{{Space}}" or " " for " "'.format(text[i:rindex+1]))
                if (len(key) == 2 and not key[1].isdigit()) or len(key) > 2:
                    raise ValueError('"{}" is not valid'.format(text[i:rindex+1]))
                upperKey = key[0].upper()
                count = 1
                if len(key) > 1:
                    count = int(key[1])
                for j in range(count):
                    if hold:
                        if upperKey in Win32API.SpecialKeyDict:
                            keyValue = Win32API.SpecialKeyDict[upperKey]
                            if type(lastKeyValue) == type(keyValue) and lastKeyValue == keyValue:
                                insertIndex += 1
                            printKeys.insert(insertIndex, (key[0], 'KeyDown | ExtendedKey'))
                            printKeys.insert(insertIndex+1, (key[0], 'KeyUp | ExtendedKey'))
                            keys.insert(insertIndex, (keyValue, KeyboardEventFlags.KeyDown | KeyboardEventFlags.ExtendedKey))
                            keys.insert(insertIndex+1, (keyValue, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey))
                            lastKeyValue = keyValue
                        elif key[0] in Win32API.CharacterDict:
                            keyValue = Win32API.CharacterDict[key[0]]
                            if type(lastKeyValue) == type(keyValue) and lastKeyValue == keyValue:
                                insertIndex += 1
                            printKeys.insert(insertIndex, (key[0], 'KeyDown | ExtendedKey'))
                            printKeys.insert(insertIndex+1, (key[0], 'KeyUp | ExtendedKey'))
                            keys.insert(insertIndex, (keyValue, KeyboardEventFlags.KeyDown | KeyboardEventFlags.ExtendedKey))
                            keys.insert(insertIndex+1, (keyValue, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey))
                            lastKeyValue = keyValue
                        else:
                            printKeys.insert(insertIndex, (key[0], 'UnicodeChar'))
                            keys.insert(insertIndex, (key[0], 'UnicodeChar'))
                            lastKeyValue = key[0]
                        if include:
                            insertIndex += 1
                        else:
                            if upperKey in holdKeys:
                                insertIndex += 1
                            else:
                                hold = False
                    else:
                        if upperKey in Win32API.SpecialKeyDict:
                            keyValue = Win32API.SpecialKeyDict[upperKey]
                            printKeys.append((key[0], 'KeyDown | ExtendedKey'))
                            printKeys.append((key[0], 'KeyUp | ExtendedKey'))
                            keys.append((keyValue, KeyboardEventFlags.KeyDown | KeyboardEventFlags.ExtendedKey))
                            keys.append((keyValue, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey))
                            lastKeyValue = keyValue
                            if upperKey in holdKeys:
                                hold = True
                                insertIndex = len(keys) - 1
                            else:
                                hold = False
                        else:
                            printKeys.append((key[0], 'UnicodeChar'))
                            keys.append((key[0], 'UnicodeChar'))
                            lastKeyValue = key[0]
                    #lastKey = key[0]
                i = rindex + 1
            elif text[i] == '(':
                if hold:
                    include = True
                else:
                    printKeys.append((text[i], 'UnicodeChar'))
                    keys.append((text[i], 'UnicodeChar'))
                    lastKeyValue = text[i]
                #lastKey = text[i]
                i += 1
            elif text[i] == ')':
                if hold:
                    include = False
                    hold = False
                else:
                    printKeys.append((text[i], 'UnicodeChar'))
                    keys.append((text[i], 'UnicodeChar'))
                    lastKeyValue = text[i]
                #lastKey = text[i]
                i += 1
            else:
                if hold:
                    if text[i] in Win32API.CharacterDict:
                        keyValue = Win32API.CharacterDict[text[i]]
                        if include and type(lastKeyValue) == type(keyValue) and lastKeyValue == keyValue:
                            insertIndex += 1
                        printKeys.insert(insertIndex, (text[i], 'KeyDown | ExtendedKey'))
                        printKeys.insert(insertIndex + 1, (text[i], 'KeyUp | ExtendedKey'))
                        keys.insert(insertIndex, (keyValue, KeyboardEventFlags.KeyDown | KeyboardEventFlags.ExtendedKey))
                        keys.insert(insertIndex + 1, (keyValue, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey))
                        lastKeyValue = keyValue
                    else:
                        printKeys.append((text[i], 'UnicodeChar'))
                        keys.append((text[i], 'UnicodeChar'))
                        lastKeyValue = text[i]
                    if include:
                        insertIndex += 1
                    else:
                        hold = False
                else:
                    printKeys.append((text[i], 'UnicodeChar'))
                    keys.append((text[i], 'UnicodeChar'))
                    lastKeyValue = text[i]
                #lastKey = text[i]
                i += 1
            if i >= length:
                break
        if debug:
            for i, key in enumerate(printKeys):
                if key[1] == 'UnicodeChar':
                    Logger.ColorfulWrite('<Color=DarkGreen>{}</Color>, sleep({})\n'.format(key, interval), writeToFile = False)
                else:
                    if i + 1 == len(printKeys):
                        Logger.ColorfulWrite('<Color=DarkGreen>{}</Color>, sleep({})\n'.format(key, interval), writeToFile = False)
                    else:
                        if 'KeyUp'in key[1] and (printKeys[i+1][1] == 'UnicodeChar' or 'KeyDown' in printKeys[i+1][1]):
                            Logger.ColorfulWrite('<Color=DarkGreen>{}</Color>, sleep({})\n'.format(key, interval), writeToFile = False)
                        else:
                            Logger.ColorfulWrite('<Color=DarkGreen>{}</Color>\n'.format(key), writeToFile = False)
            Logger.Write('\n', writeToFile = False)
        for i, key in enumerate(keys):
            if key[1] == 'UnicodeChar':
                _AutomationClient.instance().dll.SendUnicodeChar(ctypes.c_wchar_p(key[0]))
                time.sleep(interval)
                #Win32API.PostMessage(GetFocusedControl().Handle, 0x102, -key[0], 0)#UnicodeChar = 0x102
            else:
                scanCode = Win32API.VKtoSC(key[0])
                Win32API.keybd_event(key[0], scanCode, key[1], 0)
                if i + 1 == len(keys):
                    time.sleep(interval)
                else:
                    if key[1] & KeyboardEventFlags.KeyUp:
                        if keys[i+1][1] == 'UnicodeChar' or keys[i+1][1] & KeyboardEventFlags.KeyUp == 0:
                            time.sleep(interval)
                        else:
                            time.sleep(0.01)  #must sleep for a while, otherwise combined keys may not be caught
                    else:  #KeyboardEventFlags.KeyDown
                        time.sleep(0.01)
        #make sure hold keys are not pressed
        #win = ctypes.windll.user32.GetAsyncKeyState(Keys.VK_LWIN)
        #ctrl = ctypes.windll.user32.GetAsyncKeyState(Keys.VK_CONTROL)
        #alt = ctypes.windll.user32.GetAsyncKeyState(Keys.VK_MENU)
        #shift = ctypes.windll.user32.GetAsyncKeyState(Keys.VK_SHIFT)
        #if win & 0x8000:
            #Logger.WriteLine('ERROR: WIN is pressed, it should not be pressed!', ConsoleColor.Red)
            #Win32API.keybd_event(Keys.VK_LWIN, 0, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey, 0)
        #if ctrl & 0x8000:
            #Logger.WriteLine('ERROR: CTRL is pressed, it should not be pressed!', ConsoleColor.Red)
            #Win32API.keybd_event(Keys.VK_CONTROL, 0, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey, 0)
        #if alt & 0x8000:
            #Logger.WriteLine('ERROR: ALT is pressed, it should not be pressed!', ConsoleColor.Red)
            #Win32API.keybd_event(Keys.VK_MENU, 0, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey, 0)
        #if shift & 0x8000:
            #Logger.WriteLine('ERROR: SHIFT is pressed, it should not be pressed!', ConsoleColor.Red)
            #Win32API.keybd_event(Keys.VK_SHIFT, 0, KeyboardEventFlags.KeyUp | KeyboardEventFlags.ExtendedKey, 0)
        time.sleep(waitTime)


class Bitmap:
    '''A simple Bitmap class wraps Windows GDI+ Gdiplus::Bitmap, but may not have high efficiency'''
    def __init__(self, width = 0, height = 0):
        self._width = width
        self._height = height
        self._bitmap = 0
        if width > 0 and height > 0:
            self._bitmap = _AutomationClient.instance().dll.BitmapCreate(width, height)

    def __del__(self):
        self.Release()

    def _getsize(self):
        size = _AutomationClient.instance().dll.BitmapGetWidthAndHeight(self._bitmap)
        self._width = size & 0xFFFF
        self._height = size >> 16

    def Release(self):
        if self._bitmap:
            _AutomationClient.instance().dll.BitmapRelease(self._bitmap)
            self._bitmap = 0
            self._width = 0
            self._height = 0

    @property
    def Width(self):
        return self._width

    @property
    def Height(self):
        return self._height

    def FromHandle(self, hwnd, left = 0, top = 0, right = 0, bottom = 0):
        """
        capture Win32 Window to image by its handle
        left, top, right, bottom: control's internal postion(from 0,0)
        """
        self.Release()
        root = GetRootControl()
        rect = Rect()
        ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
        left, top, right, bottom = left + rect.left, top + rect.top, right + rect.left, bottom + rect.top
        self._bitmap = _AutomationClient.instance().dll.BitmapFromWindow(root.Handle, left, top, right, bottom)
        self._getsize()
        return self._bitmap > 0

    def FromControl(self, control, x = 0, y = 0, width = 0, height = 0):
        """
        capture control to Bitmap
        x, y: the point in control's internal position(from 0,0)
        width, height: image's width and height, use 0 for control's entire area
        """
        left, top, right, bottom = control.BoundingRectangle
        while (right - left) == 0 or (bottom - top) == 0:
            #some controls maybe visible but their BoundingRectangle are all 0, capture its parent util valid
            control = control.GetParentControl()
            if not control:
                return False
            left, top, right, bottom = control.BoundingRectangle
        if width <= 0:
            width = right - left + width
        if height <= 0:
            height = bottom - top + height
        hWnd = control.Handle
        if hWnd:
            left = x
            top = y
            right = left + width
            bottom = top + height
        else:
            while True:
                control = control.GetParentControl()
                hWnd = control.Handle
                if hWnd:
                    pleft, ptop, pright, pbottom = control.BoundingRectangle
                    left = left - pleft + x
                    top = top - ptop + y
                    right = left + width
                    bottom = top + height
                    break
        return self.FromHandle(hWnd, left, top, right, bottom)

    def FromFile(self, filePath):
        """load image from file"""
        self.Release()
        self._bitmap = _AutomationClient.instance().dll.BitmapFromFile(ctypes.c_wchar_p(filePath))
        self._getsize()
        return self._bitmap > 0

    def ToFile(self, savePath):
        """savePath should end with .bmp, .jpg, .jpeg, .png, .gif, .tif, .tiff"""
        name, ext = os.path.splitext(savePath);
        extMap = {'.bmp': 'image/bmp'
                  , '.jpg': 'image/jpeg'
                  , '.jpeg': 'image/jpeg'
                  , '.gif': 'image/gif'
                  , '.tif': 'image/tiff'
                  , '.tiff': 'image/tiff'
                  , '.png': 'image/png'
                  }
        gdiplusImageFormat = extMap.get(ext.lower(), 'image/png')
        return _AutomationClient.instance().dll.BitmapToFile(self._bitmap, ctypes.c_wchar_p(savePath), ctypes.c_wchar_p(gdiplusImageFormat))

    def GetPixelColor(self, x, y):
        """
        return argb
        b = argb & 0x0000FF
        g = (argb & 0x00FF00) >> 8
        r = (argb & 0xFF0000) >> 16
        a = (argb & 0xFF0000) >> 24
        """
        return _AutomationClient.instance().dll.BitmapGetPixel(self._bitmap, x, y)

    def SetPixelColor(self, x, y, argb):
        return _AutomationClient.instance().dll.BitmapSetPixel(self._bitmap, x, y, argb)

    def GetPixelColorsHorizontally(self, x, y, count):
        """get list of argb form x,y horizontally"""
        colorArray = ctypes.c_uint32 * count
        values = colorArray(*(0 for n in range(count)))
        _AutomationClient.instance().dll.BitmapGetPixelsHorizontally(ctypes.c_size_t(self._bitmap), x, y, values, count)
        return values

    def GetPixelColorsVertically(self, x, y, count):
        """get list of argb form x,y vertically"""
        colorArray = ctypes.c_uint32 * count
        values = colorArray(*(0 for n in range(count)))
        _AutomationClient.instance().dll.BitmapGetPixelsVertically(ctypes.c_size_t(self._bitmap), x, y, values, count)
        return values

    def GetPixelColorsOfRow(self, y):
        """return list of argb of y row"""
        return self.GetPixelColorsHorizontally(0, y, self.Width)

    def GetPixelColorsOfColumn(self, x):
        """return list of argb of x column"""
        return self.GetPixelColorsVertically(x, 0, self.Height)

    def GetPixelColorsOfRect(self, x, y, width, height):
        """return list of argb of rect"""
        allColors = self.GetAllPixelColors()
        colors = []
        for row in range(height):
            colors.extend(allColors[(y+row)*self.Width+x:(y+row)*self.Width+x+width])
        return colors

    def GetPixelColorsOfRects(self, rects):
        """
        return list of argb of rects
        rects example: [[x1, y1, width, height], [x2, y2, width, height]], return [rect1colors, rect2colors]
        """
        allColors = self.GetAllPixelColors()
        colorsOfRects = []
        for rect in rects:
            x, y, width, height = rect
            colors = []
            for row in range(height):
                colors.extend(allColors[(y+row)*self.Width+x:(y+row)*self.Width+x+width])
            colorsOfRects.append(colors)
        return colorsOfRects

    def GetAllPixelColors(self):
        """return all argb of all pixels horizontally from 0,0"""
        return self.GetPixelColorsHorizontally(0, 0, self.Width * self.Height)

    def SetPixelColorsHorizontally(self, x, y, colors):
        """set colors form x,y horizontally"""
        count = len(colors)
        colorArray = ctypes.c_uint32 * count
        values = colorArray(*colors)
        return _AutomationClient.instance().dll.BitmapSetPixelsHorizontally(ctypes.c_size_t(self._bitmap), x, y, values, count)

    def SetPixelColorsVertically(self, x, y, colors):
        """set colors form x,y vertically"""
        count = len(colors)
        colorArray = ctypes.c_uint32 * count
        values = colorArray(*colors)
        return _AutomationClient.instance().dll.BitmapSetPixelsVertically(ctypes.c_size_t(self._bitmap), x, y, values, count)


class LegacyIAccessiblePattern():
    def IsLegacyIAccessiblePatternAvailable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_LegacyIAccessiblePatternId) != 0

    def AccessibleSelect(self, flag):
        """call IUIAutomationLegacyIAccessiblePattern Select, flag: a value in AccessibleSelectFlag"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_LegacyIAccessiblePatternId)
        if pattern:
            _AutomationClient.instance().dll.LegacyIAccessiblePatternSelect(pattern, flag)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('LegacyIAccessiblePattern is not supported!', ConsoleColor.Yellow)

    def AccessibleDoDefaultAction(self):
        """call IUIAutomationLegacyIAccessiblePattern DoDefaultAction"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_LegacyIAccessiblePatternId)
        if pattern:
            _AutomationClient.instance().dll.LegacyIAccessiblePatternDoDefaultAction(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('LegacyIAccessiblePattern is not supported!', ConsoleColor.Yellow)

    def AccessibleSetValue(self, value):
        """call IUIAutomationLegacyIAccessiblePattern SetValue, value: str"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_LegacyIAccessiblePatternId)
        if pattern:
            _AutomationClient.instance().dll.LegacyIAccessiblePatternSetValue(pattern, ctypes.c_wchar_p(value))
            _AutomationClient.instance().dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('LegacyIAccessiblePattern is not supported!', ConsoleColor.Yellow)

    def AccessibleCurrentChildId(self):
        """call IUIAutomationLegacyIAccessiblePattern get_CurrentChildId, return int. If the element is not a child element, CHILDID_SELF (0) is returned"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_LegacyIAccessiblePatternId)
        if pattern:
            value = _AutomationClient.instance().dll.LegacyIAccessiblePatternCurrentChildId(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('LegacyIAccessiblePattern is not supported!', ConsoleColor.Yellow)

    def AccessibleCurrentName(self):
        """call IUIAutomationLegacyIAccessiblePattern get_CurrentName, return str"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_LegacyIAccessiblePatternId)
        if pattern:
            bstrValue = _AutomationClient.instance().dll.LegacyIAccessiblePatternCurrentName(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            if bstrValue:
                value = ctypes.c_wchar_p(bstrValue).value[:]
                _AutomationClient.instance().dll.FreeBSTR(bstrValue)
                return value
            return ''
        else:
            Logger.WriteLine('LegacyIAccessiblePattern is not supported!', ConsoleColor.Yellow)

    def AccessibleCurrentValue(self):
        """call IUIAutomationLegacyIAccessiblePattern get_CurrentValue, return str"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_LegacyIAccessiblePatternId)
        if pattern:
            bstrValue = _AutomationClient.instance().dll.LegacyIAccessiblePatternCurrentValue(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            if bstrValue:
                value = ctypes.c_wchar_p(bstrValue).value[:]
                _AutomationClient.instance().dll.FreeBSTR(bstrValue)
                return value
            return ''
        else:
            Logger.WriteLine('LegacyIAccessiblePattern is not supported!', ConsoleColor.Yellow)

    def AccessibleCurrentDescription(self):
        """call IUIAutomationLegacyIAccessiblePattern get_CurrentDescription, return str"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_LegacyIAccessiblePatternId)
        if pattern:
            bstrValue = _AutomationClient.instance().dll.LegacyIAccessiblePatternCurrentDescription(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            if bstrValue:
                value = ctypes.c_wchar_p(bstrValue).value[:]
                _AutomationClient.instance().dll.FreeBSTR(bstrValue)
                return value
            return ''
        else:
            Logger.WriteLine('LegacyIAccessiblePattern is not supported!', ConsoleColor.Yellow)

    def AccessibleCurrentRole(self):
        """call IUIAutomationLegacyIAccessiblePattern get_CurrentRole, return int, a value in AccessibleRole"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_LegacyIAccessiblePatternId)
        if pattern:
            value = _AutomationClient.instance().dll.LegacyIAccessiblePatternCurrentRole(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('LegacyIAccessiblePattern is not supported!', ConsoleColor.Yellow)

    def AccessibleCurrentState(self):
        """call IUIAutomationLegacyIAccessiblePattern get_CurrentState, return int, a combine value in AccessibleState"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_LegacyIAccessiblePatternId)
        if pattern:
            value = _AutomationClient.instance().dll.LegacyIAccessiblePatternCurrentState(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('LegacyIAccessiblePattern is not supported!', ConsoleColor.Yellow)

    def AccessibleCurrentHelp(self):
        """call IUIAutomationLegacyIAccessiblePattern get_CurrentHelp, return str"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_LegacyIAccessiblePatternId)
        if pattern:
            bstrValue = _AutomationClient.instance().dll.LegacyIAccessiblePatternCurrentHelp(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            if bstrValue:
                value = ctypes.c_wchar_p(bstrValue).value[:]
                _AutomationClient.instance().dll.FreeBSTR(bstrValue)
                return value
            return ''
        else:
            Logger.WriteLine('LegacyIAccessiblePattern is not supported!', ConsoleColor.Yellow)

    def AccessibleCurrentKeyboardShortcut(self):
        """call IUIAutomationLegacyIAccessiblePattern get_CurrentKeyboardShortcut, return str"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_LegacyIAccessiblePatternId)
        if pattern:
            bstrValue = _AutomationClient.instance().dll.LegacyIAccessiblePatternCurrentKeyboardShortcut(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            if bstrValue:
                value = ctypes.c_wchar_p(bstrValue).value[:]
                _AutomationClient.instance().dll.FreeBSTR(bstrValue)
                return value
            return ''
        else:
            Logger.WriteLine('LegacyIAccessiblePattern is not supported!', ConsoleColor.Yellow)

    def AccessibleGetCurrentSelection(self):
        """call IUIAutomationLegacyIAccessiblePattern GetCurrentSelection, return list of Control"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_LegacyIAccessiblePatternId)
        if pattern:
            lists = []
            iUIAutomationElementArray = _AutomationClient.instance().dll.LegacyIAccessiblePatternGetCurrentSelection(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            if iUIAutomationElementArray:
                length = _AutomationClient.instance().dll.ElementArrayGetLength(iUIAutomationElementArray)
                for i in range(length):
                    lists.append(Control.CreateControlFromElement(_AutomationClient.instance().dll.ElementArrayGetElement(iUIAutomationElementArray, i)))
                _AutomationClient.instance().dll.ReleaseElementArray(iUIAutomationElementArray)
            return lists
        else:
            Logger.WriteLine('LegacyIAccessiblePattern is not supported!', ConsoleColor.Yellow)

    def AccessibleCurrentDefaultAction(self):
        """call IUIAutomationLegacyIAccessiblePattern get_CurrentDefaultAction, return str"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_LegacyIAccessiblePatternId)
        if pattern:
            bstrValue = _AutomationClient.instance().dll.LegacyIAccessiblePatternCurrentDefaultAction(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            if bstrValue:
                value = ctypes.c_wchar_p(bstrValue).value[:]
                _AutomationClient.instance().dll.FreeBSTR(bstrValue)
                return value
            return ''
        else:
            Logger.WriteLine('LegacyIAccessiblePattern is not supported!', ConsoleColor.Yellow)

class QTPLikeSyntaxSupport():
    """
    Add syntax support like QTP, with that, user can locate object like following example:
    WindowControl(Name="SomeWindowTitle").ButtonControl(AutomationId="OneOfButton").Click()
    This class inherited by Control class
    """

    def Control(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return Control(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self, **searchPorpertyDict)

    def ButtonControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return ButtonControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self, **searchPorpertyDict)

    def CalendarControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return CalendarControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def CheckBoxControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return CheckBoxControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def ComboBoxControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return ComboBoxControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def CustomControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return CustomControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def DataGridControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return DataGridControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def DataItemControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return DataItemControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def DocumentControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return DocumentControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def EditControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return EditControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def GroupControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return GroupControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def HeaderControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return HeaderControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def HeaderItemControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return HeaderItemControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def HyperlinkControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return HyperlinkControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def ImageControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return ImageControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def ListControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return ListControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def ListItemControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return ListItemControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def MenuControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return MenuControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def MenuBarControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return MenuBarControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def MenuItemControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return MenuItemControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def PaneControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return PaneControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def ProgressBarControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return ProgressBarControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def RadioButtonControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return RadioButtonControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def ScrollBarControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return ScrollBarControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def SemanticZoomControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return SemanticZoomControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def SeparatorControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return SeparatorControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def SliderControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return SliderControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def SpinnerControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return SpinnerControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def SplitButtonControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return SplitButtonControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def StatusBarControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return StatusBarControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def TabControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return TabControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def TabItemControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return TabItemControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def TableControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return TableControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def TextControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return TextControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def ThumbControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return ThumbControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def TitleBarControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return TitleBarControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def ToolBarControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return ToolBarControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def ToolTipControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return ToolTipControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def TreeControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return TreeControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def TreeItemControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return TreeItemControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

    def WindowControl(self, element = 0, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        return WindowControl(element=element, searchDepth=searchDepth, searchWaitTime=searchWaitTime, foundIndex=foundIndex, searchFromControl = self,**searchPorpertyDict)

class Control(LegacyIAccessiblePattern, QTPLikeSyntaxSupport):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        """
        element: integer
        searchFromControl: Control, if is None, search from root control
        searchDepth: integer, max search depth from searchFromControl
        foundIndex: integer, value must be greater or equal to 1
        searchWaitTime: float, wait searchWaitTime before every search
        searchPorpertyDict: a dict that defines how to search, only the following keys are valid
                            ControlType: integer in class ControlType
                            ClassName: str or unicode
                            AutomationId: str or unicode
                            Name: str or unicode
                            SubName: str or unicode
                            RegexName: str or unicode, supports regex
                            Depth: integer, exact depth from searchFromControl, if set, searchDepth will be set to Depth too
        """
        self._element = element
        self._elementDirectAssign = True if element else False
        self.searchFromControl = searchFromControl
        self.searchDepth = searchPorpertyDict.get('Depth', searchDepth)
        self.searchWaitTime = searchWaitTime
        self.foundIndex = foundIndex
        self.searchPorpertyDict = searchPorpertyDict
        regName = searchPorpertyDict.get('RegexName', '')
        self.regexName = re.compile(regName) if regName else None

    def __del__(self):
        """
        Warning: when script exits, module ctypes may be None,
        ctypes sometimes is None before all controls are destoryed
        but __del__ needs method in dll(which needs ctypes) to release resources
        """
        if self._element:
            _AutomationClient.instance().dll.ReleaseElement(self._element)
            self._element = 0

    def SetSearchFromControl(self, searchFromControl):
        """searchFromControl: control"""
        self.searchFromControl = searchFromControl

    def SetSearchDepth(self, searchDepth):
        """searchDepth: integer"""
        self.searchDepth = searchDepth

    def AddSearchProperty(self, **searchPorpertyDict):
        """searchPorpertyDict: dict"""
        self.searchPorpertyDict.update(searchPorpertyDict)
        if 'Depth' in searchPorpertyDict:
            self.searchDepth = searchPorpertyDict['Depth']
        if 'RegexName' in searchPorpertyDict:
            regName = searchPorpertyDict['RegexName']
            self.regexName = re.compile(regName) if regName else None

    def RemoveSearchProperty(self, **searchPorpertyDict):
        for key in searchPorpertyDict:
            del self.searchPorpertyDict[key]
            if key == 'RegexName':
                self.regexName = None

    def _CompareFunction(self, control, depth):
        """This function defines how to search, return True if found"""
        for key, value in self.searchPorpertyDict.items():
            if 'ControlType' == key:
                if value != control.ControlType:
                    return False
            if 'ClassName' == key:
                if value != control.ClassName:
                    return False
            if 'AutomationId' == key:
                if value != control.AutomationId:
                    return False
            if 'Name' == key:
                if value != control.Name:
                    return False
            if 'SubName' == key:
                if value not in control.Name:
                    return False
            if 'RegexName' == key:
                if not self.regexName.match(control.Name):
                    return False
            if 'Depth' == key:
                if value != depth:
                    return False
        return True

    def Exists(self, maxSearchSeconds = 5, searchIntervalSeconds = SEARCH_INTERVAL):
        """Find control every searchIntervalSeconds seconds in maxSearchSeconds seconds, if found, return True else False"""
        if self._element and self._elementDirectAssign:
            #if element is directly assigned, not by searching, just check whether self._element is valid
            #but I can't find an API in UIAutomation that can directly check
            rootElement = GetRootControl().Element
            if self._element == rootElement:
                return True
            else:
                parentElement = _AutomationClient.instance().dll.GetParentElement(self._element)
                if parentElement:
                    _AutomationClient.instance().dll.ReleaseElement(parentElement)
                    return True
                else:
                    return False
        #find the element
        if len(self.searchPorpertyDict) == 0:
            raise LookupError("control's searchPorpertyDict must not be empty!")
        if self._element:
            _AutomationClient.instance().dll.ReleaseElement(self._element)
        self._element = 0
        start = time.clock()
        # Use same timeout(s) parameters for resolve all parents
        prev =  self.searchFromControl
        if prev and not prev._element and not prev.Exists(maxSearchSeconds, searchIntervalSeconds):
            return False
        while True:
            control = FindControl(self.searchFromControl, self._CompareFunction, self.searchDepth, False, self.foundIndex)
            if control:
                self._element = control.Element
                control._element = 0 # control will be destroyed, but the element needs to be stroed in self._element
                #elapsedTime = time.clock() - start
                #Logger.Log('Found time: {:.3f}s, {}'.format(elapsedTime, self))
                return True
            else:
                remain = start + maxSearchSeconds - time.clock()
                if remain > 0:
                    time.sleep(min(remain, searchIntervalSeconds))
                else:
                    return False

    def Disappears(self, maxSearchSeconds = 5, searchIntervalSeconds = SEARCH_INTERVAL):
        """check control disappears in maxSearchSeconds, return True if control disappears"""
        start = time.clock()
        while True:
            if not self.Exists(0, 0):
                return True
            remain = start + maxSearchSeconds - time.clock()
            if remain > 0:
                time.sleep(min(remain, searchIntervalSeconds))
            else:
                return False

    def Refind(self, maxSearchSeconds = TIME_OUT_SECOND, searchIntervalSeconds = SEARCH_INTERVAL, raiseException = True):
        """Refind the control every searchIntervalSeconds seconds in maxSearchSeconds seconds, raise an LookupError if timed out"""
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
        """Return value of control's IUIAutomationElement"""
        if not self._element:
            self.Refind(maxSearchSeconds = TIME_OUT_SECOND, searchIntervalSeconds = self.searchWaitTime)
        return self._element

    @property
    def Name(self):
        """Return unicode Name"""
        bstrValue = _AutomationClient.instance().dll.GetElementName(self.Element)
        if bstrValue:
            name = ctypes.c_wchar_p(bstrValue).value[:]
            _AutomationClient.instance().dll.FreeBSTR(bstrValue)
            return name
        return ''

    @property
    def ControlType(self):
        """Return an integer in class ControlType"""
        return _AutomationClient.instance().dll.GetElementControlType(self.Element)

    @property
    def ControlTypeName(self):
        """Return str ControlTypeName"""
        return ControlTypeNameDict[self.ControlType]

    @property
    def LocalizedControlType(self):
        """Return unicode LocalizedControlType name"""
        bstrValue = _AutomationClient.instance().dll.GetElementLocalizedControlType(self.Element)
        if bstrValue:
            name = ctypes.c_wchar_p(bstrValue).value[:]
            _AutomationClient.instance().dll.FreeBSTR(bstrValue)
            return name
        return ''

    @property
    def ClassName(self):
        """Return unicode ClassName"""
        bstrValue = _AutomationClient.instance().dll.GetElementClassName(self.Element)
        if bstrValue:
            name = ctypes.c_wchar_p(bstrValue).value[:]
            _AutomationClient.instance().dll.FreeBSTR(bstrValue)
            return name
        return ''

    @property
    def AutomationId(self):
        """Return unicode AutomationId"""
        bstrValue = _AutomationClient.instance().dll.GetElementAutomationId(self.Element)
        if bstrValue:
            name = ctypes.c_wchar_p(bstrValue).value[:]
            _AutomationClient.instance().dll.FreeBSTR(bstrValue)
            return name
        return ''

    @property
    def ProcessId(self):
        """Return process id"""
        return _AutomationClient.instance().dll.GetElementProcessId(self.Element)

    @property
    def IsEnabled(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementIsEnabled(self.Element)

    @property
    def HasKeyboardFocus(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementHasKeyboardFocus(self.Element)

    @property
    def IsKeyboardFocusable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementIsKeyboardFocusable(self.Element)

    @property
    def IsOffScreen(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementIsOffscreen(self.Element)

    @property
    def BoundingRectangle(self):
        """Return tuple (left, top, right, bottom)"""
        rect = Rect()
        _AutomationClient.instance().dll.GetElementBoundingRectangle(ctypes.c_size_t(self.Element), ctypes.byref(rect))
        return (rect.left, rect.top, rect.right, rect.bottom)

    @property
    def Handle(self):
        """Return control's handle"""
        return _AutomationClient.instance().dll.GetElementHandle(self.Element)

    def SetFocus(self):
        """Make the control have focus"""
        _AutomationClient.instance().dll.SetElementFocus(self.Element)

    def MoveCursor(self, ratioX = 0.5, ratioY = 0.5, simulateMove = True):
        """Move cursor to control's rect, default to center"""
        left, top, right, bottom = self.BoundingRectangle
        if left == 0 and top == 0 and right == 0 and bottom == 0:
            Logger.WriteLine('{}\'s BoundingRectangle is empty(0,0,0,0). searchPorperty: {}'.format(self.ControlTypeName, self.searchPorpertyDict), ConsoleColor.Yellow)
        if type(ratioX) is float:
            x = left + int((right - left) * ratioX)
        else:
            x = (left if ratioX >= 0 else right) + ratioX
        if type(ratioY) is float:
            y = top + int((bottom - top) * ratioY)
        else:
            y = (top if ratioY >= 0 else bottom) + ratioY
        if simulateMove and MAX_MOVE_SECOND > 0:
            Win32API.MouseMoveTo(x, y, waitTime = 0)
        else:
            Win32API.SetCursorPos(x, y)
        return x, y

    def MoveCursorToMyCenter(self, simulateMove = True):
        """Move cursor to control's center"""
        return self.MoveCursor(simulateMove = simulateMove)

    def Click(self, ratioX = 0.5, ratioY = 0.5, simulateMove = True, waitTime = OPERATION_WAIT_TIME):
        """
        Click(0.5, 0.5): click center
        Click(10, 10): click left+10, top+10
        Click(-10, -10): click right-10, bottom-10
        simulateMove: bool, if True, first move cursor to control smoothly
        """
        x, y = self.MoveCursor(ratioX, ratioY, simulateMove)
        Win32API.MouseClick(x, y, waitTime)

    def MiddleClick(self, ratioX = 0.5, ratioY = 0.5, simulateMove = True, waitTime = OPERATION_WAIT_TIME):
        """
        Click(0.5, 0.5): click center
        Click(10, 10): click left+10, top+10
        Click(-10, -10): click right-10, bottom-10
        simulateMove: bool, if True, first move cursor to control smoothly
        """
        x, y = self.MoveCursor(ratioX, ratioY, simulateMove)
        Win32API.MouseMiddleClick(x, y, waitTime)

    def RightClick(self, ratioX = 0.5, ratioY = 0.5, simulateMove = True, waitTime = OPERATION_WAIT_TIME):
        """
        RightClick(0.5, 0.5): right click center
        RightClick(10, 10): right click left+10, top+10
        RightClick(-10, -10): click right-10, bottom-10
        simulateMove: bool, if True, first move cursor to control smoothly
        """
        x, y = self.MoveCursor(ratioX, ratioY, simulateMove)
        Win32API.MouseRightClick(x, y, waitTime)

    def DoubleClick(self, ratioX = 0.5, ratioY = 0.5, simulateMove = True, waitTime = OPERATION_WAIT_TIME):
        """
        DoubleClick(0.5, 0.5): double click center
        DoubleClick(10, 10): double click left+10, top+10
        DoubleClick(-10, -10): click right-10, bottom-10
        simulateMove: bool, if True, first move cursor to control smoothly
        """
        x, y = self.MoveCursor(ratioX, ratioY, simulateMove)
        Win32API.MouseClick(x, y, 0)
        time.sleep(Win32API.GetDoubleClickTime() * 1.0 / 2000)
        Win32API.MouseClick(x, y, waitTime)

    def WheelDown(self, wheelTimes = 1, waitTime = OPERATION_WAIT_TIME):
        """
        Mouse wheel down
        """
        x, y = Win32API.GetCursorPos()
        self.SetFocus()
        self.MoveCursorToMyCenter(False)
        Win32API.MouseWheelDown(wheelTimes, waitTime)
        Win32API.SetCursorPos(x, y)

    def WheelUp(self, wheelTimes = 1, waitTime = OPERATION_WAIT_TIME):
        """
        Mouse wheel up
        """
        x, y = Win32API.GetCursorPos()
        self.SetFocus()
        self.MoveCursorToMyCenter(False)
        Win32API.MouseWheelUp(wheelTimes, waitTime)
        Win32API.SetCursorPos(x, y)

    def GetParentControl(self):
        """Return Control"""
        comEle = _AutomationClient.instance().dll.GetParentElement(self.Element)
        if comEle:
            return Control.CreateControlFromElement(comEle)

    def GetTopWindow(self):
        """Return control's top window"""
        control = self
        parents = []
        while control:
            if control.ControlType == ControlType.WindowControl:
                return control
            parents.insert(0, control)
            control = control.GetParentControl()
        if len(parents) > 1:
            return parents[1]
        elif len(parents) == 1:
            return parents[0]

    def GetFirstChildControl(self):
        """Return Control"""
        comEle = _AutomationClient.instance().dll.GetFirstChildElement(self.Element)
        if comEle:
            return Control.CreateControlFromElement(comEle)

    def GetLastChildControl(self):
        """Return Control"""
        comEle = _AutomationClient.instance().dll.GetLastChildElement(self.Element)
        if comEle:
            return Control.CreateControlFromElement(comEle)

    def GetNextSiblingControl(self):
        """Return Control"""
        comEle = _AutomationClient.instance().dll.GetNextSiblingElement(self.Element)
        if comEle:
            return Control.CreateControlFromElement(comEle)

    def GetPreviousSiblingControl(self):
        """Return Control"""
        comEle = _AutomationClient.instance().dll.GetPreviousSiblingElement(self.Element)
        if comEle:
            return Control.CreateControlFromElement(comEle)

    def GetChildren(self):
        """Return a list of control's children"""
        children = []
        child = self.GetFirstChildControl()
        while child:
            children.append(child)
            child = child.GetNextSiblingControl()
        return children

    def ShowWindow(self, cmdShow):
        """
        ShowWindow(ShowWindow.Show), only works if Handle is valid
        cmdShow: see values in class ShowWindow
        """
        hWnd = self.Handle
        if hWnd:
            return Win32API.ShowWindow(hWnd, cmdShow)

    def Show(self):
        """call ShowWindow(ShowWindow.Show), only works if Handle is valid"""
        return self.ShowWindow(ShowWindow.Show)

    def Hide(self):
        """call ShowWindow(ShowWindow.Hide), only works if Handle is valid"""
        return self.ShowWindow(ShowWindow.Hide)

    def MoveWindow(self, x, y, width, height, repaint = 1):
        """only works if Handle is valid"""
        hWnd = self.Handle
        if hWnd:
            return Win32API.MoveWindow(hWnd, x, y, width, height, repaint)

    def GetWindowText(self):
        """only works if Handle is valid"""
        hWnd = self.Handle
        if hWnd:
            return Win32API.GetWindowText(hWnd)

    def SetWindowText(self, text):
        """only works if Handle is valid"""
        hWnd = self.Handle
        if hWnd:
            return Win32API.SetWindowText(hWnd, text)

    def SendKey(self, key, waitTime = OPERATION_WAIT_TIME):
        """
        Simulate typing a key
        key: a value in class Keys
        """
        self.SetFocus()
        Win32API.SendKey(key, waitTime)

    def SendKeys(self, keys, interval = 0.01, waitTime = OPERATION_WAIT_TIME):
        """
        Simulate typing keys
        keys: str, keys to type, see docstring of Win32API.SendKeys
        interval: double, seconds between keys
        """
        self.SetFocus()
        Win32API.SendKeys(keys, interval, waitTime)

    def GetPixelColor(self, x, y):
        """return r, g, b, int, 0xhexstring, #hexstring, only works if Handle is valid"""
        hWnd = self.Handle
        if hWnd:
            return Win32API.GetPixelColor(x, y, hWnd)

    def ToBitmap(self, x = 0, y = 0, width = 0, height = 0):
        """
        capture control to Bitmap object
        x, y: the point in control's internal position(from 0,0)
        width, height: image's width and height, use 0 for entire area
        """
        bitmap = Bitmap()
        bitmap.FromControl(self, x, y, width, height)
        return bitmap

    def CaptureToImage(self, savePath, x = 0, y = 0, width = 0, height = 0):
        """
        capture control to image file
        savePath, savePath shoud end with .bmp, .jpg, .jpeg, .png, .gif, .tif, .tiff
        x, y: the point in control's internal position(from 0,0)
        width, height: image's width and height, use 0 for entire area
        """
        bitmap = Bitmap()
        if bitmap.FromControl(self, x, y, width, height):
            return bitmap.ToFile(savePath)

    def Convert(self):
        """
        Convert Control to a specific Control
        for example: if self's ControlType is EditControl, return an EditControl
        """
        return Control.CreateControlFromControl(self)

    @staticmethod
    def CreateControlFromElement(element):
        """element: value of IUIAutomationElement"""
        if element:
            controlType = _AutomationClient.instance().dll.GetElementControlType(element)
            if controlType in ControlDict:
                return ControlDict[controlType](element)
            else:
                Logger.WriteLine("GetElementControlType returns {}, invalid ControlType!".format(controlType), ConsoleColor.Red)  #rarely happens

    @staticmethod
    def CreateControlFromControl(control):
        """
        control: Control, will add ref for control's element
        return a specific Control
        for example: if control's ControlType is EditControl, return an EditControl
        """
        newControl = Control.CreateControlFromElement(control.Element)
        _AutomationClient.instance().dll.ElementAddRef(control.Element)
        return newControl

    def __str__(self):
        if IsPy3:
            return 'ControlType: {0}    ClassName: {1}    AutomationId: {2}    Rect: {3}    Name: {4}    Handle: 0x{5:X}({5})'.format(self.ControlTypeName, self.ClassName, self.AutomationId, self.BoundingRectangle, self.Name, self.Handle)
        else:
            strClassName = self.ClassName.encode('gbk')
            strAutomationId = self.AutomationId.encode('gbk')
            try:
                strName = self.Name.encode('gbk')
            except Exception:
                strName = 'Error occured: Name can\'t be converted to gbk, try unicode'
            return 'ControlType: {0}    ClassName: {1}    AutomationId: {2}    Rect: {3}    Name: {4}    Handle: 0x{5:X}({5})'.format(self.ControlTypeName, strClassName, strAutomationId, self.BoundingRectangle, strName, self.Handle)


    # def __repr__(self):
        # return '[{0}]'.format(self)

    def __unicode__(self):
        return u'ControlType: {0}    ClassName: {1}    AutomationId: {2}    Rect: {3}    Name: {4}    Handle: 0x{0:X}({1})'.format(self.ControlTypeName, self.ClassName, self.AutomationId, self.BoundingRectangle, self.Name, self.Handle)


#Patterns -----

class DockPattern:
    def IsDockPatternAvailable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_DockPatternId) != 0
    #todo


class ExpandCollapsePattern:
    def IsExpandCollapsePatternAvailable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_ExpandCollapsePatternId) != 0

    def Expand(self, waitTime = OPERATION_WAIT_TIME):
        """Expand the control"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_ExpandCollapsePatternId)
        if pattern:
            _AutomationClient.instance().dll.ExpandCollapsePatternExpand(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            time.sleep(waitTime)
        else:
            Logger.WriteLine('ExpandCollapsePattern is not supported!', ConsoleColor.Yellow)

    def Collapse(self, waitTime = OPERATION_WAIT_TIME):
        """Collapse the control"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_ExpandCollapsePatternId)
        if pattern:
            _AutomationClient.instance().dll.ExpandCollapsePatternCollapse(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            time.sleep(waitTime)
        else:
            Logger.WriteLine('ExpandCollapsePattern is not supported!', ConsoleColor.Yellow)

    def CurrentExpandCollapseState(self):
        """Return an integer of class ExpandCollapseState """
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_ExpandCollapsePatternId)
        if pattern:
            state = _AutomationClient.instance().dll.ExpandCollapsePatternCurrentExpandCollapseState(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return state
        else:
            Logger.WriteLine('ExpandCollapsePattern is not supported!', ConsoleColor.Yellow)


class GridItemPattern:
    def IsGridItemPatternAvailable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_GridItemPatternId) != 0

    def CurrentContainingGrid(self):
        """return Control"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_GridItemPatternId)
        if pattern:
            element = _AutomationClient.instance().dll.GridItemPatternCurrentContainingGrid(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            if element:
                return Control.CreateControlFromElement(element)
        else:
            Logger.WriteLine('GridItemPattern is not supported!', ConsoleColor.Yellow)

    def CurrentRow(self):
        """return int"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_GridItemPatternId)
        if pattern:
            value = _AutomationClient.instance().dll.GridItemPatternCurrentRow(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('GridItemPattern is not supported!', ConsoleColor.Yellow)

    def CurrentColumn(self):
        """return int"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_GridItemPatternId)
        if pattern:
            value = _AutomationClient.instance().dll.GridItemPatternCurrentColumn(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('GridItemPattern is not supported!', ConsoleColor.Yellow)

    def CurrentRowSpan(self):
        """return int"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_GridItemPatternId)
        if pattern:
            value = _AutomationClient.instance().dll.GridItemPatternCurrentRowSpan(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('GridItemPattern is not supported!', ConsoleColor.Yellow)

    def CurrentColumnSpan(self):
        """return int"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_GridItemPatternId)
        if pattern:
            value = _AutomationClient.instance().dll.GridItemPatternCurrentColumnSpan(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('GridItemPattern is not supported!', ConsoleColor.Yellow)


class GridPattern:
    def IsGridPatternAvailable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_GridPatternId) != 0

    def GetItem(self, row, column):
        """return Control"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_GridPatternId)
        if pattern:
            element = _AutomationClient.instance().dll.GridPatternGetItem(pattern, row, column)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            if element:
                return Control.CreateControlFromElement(element)
        else:
            Logger.WriteLine('GridPattern is not supported!', ConsoleColor.Yellow)

    def CurrentRowCount(self):
        """return int"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_GridPatternId)
        if pattern:
            value = _AutomationClient.instance().dll.GridPatternCurrentRowCount(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('GridPattern is not supported!', ConsoleColor.Yellow)

    def CurrentColumnCount(self):
        """return int"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_GridPatternId)
        if pattern:
            value = _AutomationClient.instance().dll.GridPatternCurrentColumnCount(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('GridPattern is not supported!', ConsoleColor.Yellow)


class InvokePattern:
    def IsInvokePatternAvailable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_InvokePatternId) != 0

    def Invoke(self, waitTime = OPERATION_WAIT_TIME):
        """invoke"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_InvokePatternId)
        if pattern:
            _AutomationClient.instance().dll.InvokePatternInvoke(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            time.sleep(waitTime)
        else:
            Logger.WriteLine('InvokePattern is not supported!', ConsoleColor.Yellow)


class MultipleViewPattern:
    def IsMultipleViewPatternAvailable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_MultipleViewPatternId) != 0
    #todo


class ScrollItemPattern:
    def IsScrollItemPatternAvailable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_ScrollItemPatternId) != 0

    def ScrollIntoView(self):
        """Scroll the control into view, so it can be seen"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_ScrollItemPatternId)
        if pattern:
            _AutomationClient.instance().dll.ScrollItemPatternScrollIntoView(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('ScrollItemPattern is not supported!', ConsoleColor.Yellow)


class ScrollPattern:
    def IsScrollPatternAvailable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_ScrollPatternId) != 0

    def CurrentHorizontallyScrollable(self):
        """Return bool"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_ScrollPatternId)
        if pattern:
            scroll = _AutomationClient.instance().dll.ScrollPatternCurrentHorizontallyScrollable(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return scroll
        else:
            Logger.WriteLine('ScrollPattern is not supported!', ConsoleColor.Yellow)

    def CurrentHorizontalViewSize(self):
        """Return integer"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_ScrollPatternId)
        if pattern:
            size = _AutomationClient.instance().dll.ScrollPatternCurrentHorizontalViewSize(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return size
        else:
            Logger.WriteLine('ScrollPattern is not supported!', ConsoleColor.Yellow)

    def CurrentHorizontalScrollPercent(self):
        """Return integer"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_ScrollPatternId)
        if pattern:
            percent = _AutomationClient.instance().dll.ScrollPatternCurrentHorizontalScrollPercent(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return percent
        else:
            Logger.WriteLine('ScrollPattern is not supported!', ConsoleColor.Yellow)

    def CurrentVerticallyScrollable(self):
        """Return bool"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_ScrollPatternId)
        if pattern:
            scroll = _AutomationClient.instance().dll.ScrollPatternCurrentVerticallyScrollable(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return scroll
        else:
            Logger.WriteLine('ScrollPattern is not supported!', ConsoleColor.Yellow)

    def CurrentVerticalViewSize(self):
        """Return integer"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_ScrollPatternId)
        if pattern:
            size = _AutomationClient.instance().dll.ScrollPatternCurrentVerticalViewSize(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return size
        else:
            Logger.WriteLine('ScrollPattern is not supported!', ConsoleColor.Yellow)

    def CurrentVerticalScrollPercent(self):
        """Return integer"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_ScrollPatternId)
        if pattern:
            percent = _AutomationClient.instance().dll.ScrollPatternCurrentVerticalScrollPercent(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return percent
        else:
            Logger.WriteLine('ScrollPattern is not supported!', ConsoleColor.Yellow)

    def SetScrollPercent(self, horizontalPercent, verticalPercent):
        """Need two integers as parameters"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_ScrollPatternId)
        if pattern:
            _AutomationClient.instance().dll.ScrollPatternSetScrollPercent(pattern, horizontalPercent, verticalPercent)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('ScrollPattern is not supported!', ConsoleColor.Yellow)


class SelectionItemPattern:
    def IsSelectionItemPatternAvailable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_SelectionItemPatternId) != 0

    def Select(self):
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_SelectionItemPatternId)
        if pattern:
            _AutomationClient.instance().dll.SelectionItemPatternSelect(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('SelectionItemPattern is not supported!', ConsoleColor.Yellow)

    def AddToSelection(self):
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_SelectionItemPatternId)
        if pattern:
            _AutomationClient.instance().dll.SelectionItemPatternAddToSelection(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('SelectionItemPattern is not supported!', ConsoleColor.Yellow)

    def RemoveFromSelection(self):
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_SelectionItemPatternId)
        if pattern:
            _AutomationClient.instance().dll.SelectionItemPatternRemoveFromSelection(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('SelectionItemPattern is not supported!', ConsoleColor.Yellow)

    def CurrentIsSelected(self):
        """Return bool"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_SelectionItemPatternId)
        if pattern:
            isSelect = _AutomationClient.instance().dll.SelectionItemPatternCurrentIsSelected(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return bool(isSelect)
        else:
            Logger.WriteLine('SelectionItemPattern is not supported!', ConsoleColor.Yellow)


class SelectionPattern:
    def IsSelectionPatternAvailable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_SelectionPatternId) != 0

    def GetCurrentSelection(self):
        """Return an IUIAutomationElementArray"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_SelectionPatternId)
        if pattern:
            pElementArray = _AutomationClient.instance().dll.SelectionPatternGetCurrentSelection(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return pElementArray
        else:
            Logger.WriteLine('SelectionPattern is not supported!', ConsoleColor.Yellow)


class RangeValuePattern:
    def IsRangeValuePatternAvailable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_RangeValuePatternId) != 0

    def RangeValuePatternCurrentValue(self):
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_RangeValuePatternId)
        if pattern:
            value = _AutomationClient.instance().dll.RangeValuePatternCurrentValue(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('RangeValuePattern is not supported!', ConsoleColor.Yellow)

    def RangeValuePatternSetValue(self, value):
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_RangeValuePatternId)
        if pattern:
            _AutomationClient.instance().dll.RangeValuePatternSetValue(pattern, value)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('RangeValuePattern is not supported!', ConsoleColor.Yellow)

    def CurrentMaximum(self):
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_RangeValuePatternId)
        if pattern:
            value = _AutomationClient.instance().dll.RangeValuePatternCurrentMaximum(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('RangeValuePattern is not supported!', ConsoleColor.Yellow)

    def CurrentMinimum(self):
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_RangeValuePatternId)
        if pattern:
            value = _AutomationClient.instance().dll.RangeValuePatternCurrentMinimum(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('RangeValuePattern is not supported!', ConsoleColor.Yellow)


class TableItemPattern:
    def IsTableItemPatternAvailable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_TableItemPatternId) != 0

    def CurrentRowHeaderItems(self):
        """return list of Control"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_TableItemPatternId)
        if pattern:
            lists = []
            iUIAutomationElementArray = _AutomationClient.instance().dll.TableItemPatternCurrentRowHeaderItems(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            if iUIAutomationElementArray:
                length = _AutomationClient.instance().dll.ElementArrayGetLength(iUIAutomationElementArray)
                for i in range(length):
                    lists.append(Control.CreateControlFromElement(_AutomationClient.instance().dll.ElementArrayGetElement(iUIAutomationElementArray, i)))
                _AutomationClient.instance().dll.ReleaseElementArray(iUIAutomationElementArray)
            return lists
        else:
            Logger.WriteLine('TableItemPattern is not supported!', ConsoleColor.Yellow)

    def CurrentColumnHeaderItems(self):
        """return list of Control"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_TableItemPatternId)
        if pattern:
            lists = []
            iUIAutomationElementArray = _AutomationClient.instance().dll.TableItemPatternCurrentColumnHeaderItems(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            if iUIAutomationElementArray:
                length = _AutomationClient.instance().dll.ElementArrayGetLength(iUIAutomationElementArray)
                for i in range(length):
                    lists.append(Control.CreateControlFromElement(_AutomationClient.instance().dll.ElementArrayGetElement(iUIAutomationElementArray, i)))
                _AutomationClient.instance().dll.ReleaseElementArray(iUIAutomationElementArray)
            return lists
        else:
            Logger.WriteLine('TableItemPattern is not supported!', ConsoleColor.Yellow)


class TablePattern:
    def IsTablePatternAvailable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_TablePatternId) != 0

    def CurrentRowHeaders(self):
        """return list of Control"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_TablePatternId)
        if pattern:
            lists = []
            iUIAutomationElementArray = _AutomationClient.instance().dll.TablePatternCurrentRowHeaders(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            if iUIAutomationElementArray:
                length = _AutomationClient.instance().dll.ElementArrayGetLength(iUIAutomationElementArray)
                for i in range(length):
                    lists.append(Control.CreateControlFromElement(_AutomationClient.instance().dll.ElementArrayGetElement(iUIAutomationElementArray, i)))
                _AutomationClient.instance().dll.ReleaseElementArray(iUIAutomationElementArray)
            return lists
        else:
            Logger.WriteLine('TablePattern is not supported!', ConsoleColor.Yellow)

    def CurrentColumnHeaders(self):
        """return list of Control"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_TablePatternId)
        if pattern:
            lists = []
            iUIAutomationElementArray = _AutomationClient.instance().dll.TablePatternCurrentColumnHeaders(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            if iUIAutomationElementArray:
                length = _AutomationClient.instance().dll.ElementArrayGetLength(iUIAutomationElementArray)
                for i in range(length):
                    lists.append(Control.CreateControlFromElement(_AutomationClient.instance().dll.ElementArrayGetElement(iUIAutomationElementArray, i)))
                _AutomationClient.instance().dll.ReleaseElementArray(iUIAutomationElementArray)
            return lists
        else:
            Logger.WriteLine('TablePattern is not supported!', ConsoleColor.Yellow)

    def CurrentRowOrColumnMajor(self):
        """return int, a value in RowOrColumnMajor"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_TablePatternId)
        if pattern:
            value = _AutomationClient.instance().dll.TablePatternCurrentRowOrColumnMajor(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('TablePattern is not supported!', ConsoleColor.Yellow)


class TextPattern:
    def IsTextPatternAvailable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_TextPatternId) != 0
    #todo


class TogglePattern:
    def IsTogglePatternAvailable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_TogglePatternId) != 0

    def Toggle(self, waitTime = OPERATION_WAIT_TIME):
        """Toggle or UnToggle the control"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_TogglePatternId)
        if pattern:
            _AutomationClient.instance().dll.TogglePatternToggle(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            time.sleep(waitTime)
        else:
            Logger.WriteLine('TogglePattern is not supported!', ConsoleColor.Yellow)

    def CurrentToggleState(self):
        """Return an integer of class ToggleState"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_TogglePatternId)
        if pattern:
            state = _AutomationClient.instance().dll.TogglePatternCurrentToggleState(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return state
        else:
            Logger.WriteLine('TogglePattern is not supported!', ConsoleColor.Yellow)


class TransformPattern:
    def IsTransformPatternAvailable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_TransformPatternId) != 0

    def Move(self, x, y):
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_TransformPatternId)
        if pattern:
            _AutomationClient.instance().dll.TransformPatternMove(pattern, x, y)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('TransformPattern is not supported!', ConsoleColor.Yellow)

    def Resize(self, width, height):
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_TransformPatternId)
        if pattern:
            _AutomationClient.instance().dll.TransformPatternResize(pattern, width, height)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('TransformPattern is not supported!', ConsoleColor.Yellow)

    def Rotate(self, degrees):
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_TransformPatternId)
        if pattern:
            _AutomationClient.instance().dll.TransformPatternRotate(pattern, degrees)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
        else:
            Logger.WriteLine('TransformPattern is not supported!', ConsoleColor.Yellow)


class TransformPattern2:
    def IsTransformPattern2Available(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_TransformPattern2Id) != 0


class ValuePattern:
    def IsValuePatternAvailable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_ValuePatternId) != 0

    def CurrentValue(self):
        """Return unicode string"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_ValuePatternId)
        if pattern:
            bstrValue = _AutomationClient.instance().dll.ValuePatternCurrentValue(pattern)
            if bstrValue:
                value = ctypes.c_wchar_p(bstrValue).value[:]
                _AutomationClient.instance().dll.ReleasePattern(pattern)
                _AutomationClient.instance().dll.FreeBSTR(bstrValue)
                return value
        else:
            Logger.WriteLine('ValuePattern is not supported!', ConsoleColor.Yellow)
        return ''

    def SetValue(self, value, waitTime = OPERATION_WAIT_TIME):
        """Set unicode string to control's value"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_ValuePatternId)
        if pattern:
            value = _AutomationClient.instance().dll.ValuePatternSetValue(pattern, ctypes.c_wchar_p(value))
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            time.sleep(waitTime)
        else:
            Logger.WriteLine('ValuePattern is not supported!', ConsoleColor.Yellow)

    def CurrentIsReadOnly(self):
        """Return bool"""
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_ValuePatternId)
        if pattern:
            isReadOnly = _AutomationClient.instance().dll.ValuePatternCurrentIsReadOnly(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return bool(isReadOnly)
        else:
            Logger.WriteLine('ValuePattern is not supported!', ConsoleColor.Yellow)


class WindowPattern:
    def IsWindowPatternAvailable(self):
        """Return bool"""
        return _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_WindowPatternId) != 0

    def CurrentWindowVisualState(self):
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_WindowPatternId)
        if pattern:
            value = _AutomationClient.instance().dll.WindowPatternCurrentWindowVisualState(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('WindowPattern is not supported!', ConsoleColor.Yellow)

    def SetWindowVisualState(self, value, waitTime = OPERATION_WAIT_TIME):
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_WindowPatternId)
        if pattern:
            _AutomationClient.instance().dll.WindowPatternSetWindowVisualState(pattern, value)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            time.sleep(waitTime)
        else:
            Logger.WriteLine('WindowPattern is not supported!', ConsoleColor.Yellow)

    def CurrentCanMaximize(self):
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_WindowPatternId)
        if pattern:
            value = _AutomationClient.instance().dll.WindowPatternCurrentCanMaximize(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('WindowPattern is not supported!', ConsoleColor.Yellow)

    def Maximize(self, waitTime = OPERATION_WAIT_TIME):
        if self.CurrentCanMaximize():
            self.SetWindowVisualState(WindowVisualState.Maximized, waitTime)

    def CurrentCanMinimize(self):
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_WindowPatternId)
        if pattern:
            value = _AutomationClient.instance().dll.WindowPatternCurrentCanMinimize(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return value
        else:
            Logger.WriteLine('WindowPattern is not supported!', ConsoleColor.Yellow)

    def Minimize(self, waitTime = OPERATION_WAIT_TIME):
        if self.CurrentCanMinimize():
            self.SetWindowVisualState(WindowVisualState.Minimized, waitTime)

    def Normal(self, waitTime = OPERATION_WAIT_TIME):
        self.SetWindowVisualState(WindowVisualState.Normal, waitTime)

    def IsMaximize(self):
        return self.CurrentWindowVisualState() == WindowVisualState.Maximized

    def IsMinimize(self):
        return self.CurrentWindowVisualState() == WindowVisualState.Minimized

    def CurrentIsModal(self):
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_WindowPatternId)
        if pattern:
            value = _AutomationClient.instance().dll.WindowPatternCurrentIsModal(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return bool(value)
        else:
            Logger.WriteLine('WindowPattern is not supported!', ConsoleColor.Yellow)

    def CurrentIsTopmost(self):
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_WindowPatternId)
        if pattern:
            value = _AutomationClient.instance().dll.WindowPatternCurrentIsTopmost(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            return bool(value)
        else:
            Logger.WriteLine('WindowPattern is not supported!', ConsoleColor.Yellow)

    def Close(self, waitTime = OPERATION_WAIT_TIME):
        pattern = _AutomationClient.instance().dll.GetElementPattern(self.Element, PatternId.UIA_WindowPatternId)
        if pattern:
            _AutomationClient.instance().dll.WindowPatternClose(pattern)
            _AutomationClient.instance().dll.ReleasePattern(pattern)
            time.sleep(waitTime)
        else:
            Logger.WriteLine('WindowPattern is not supported!', ConsoleColor.Yellow)


#see Control Pattern Mapping for UI Automation Clients
#https://msdn.microsoft.com/zh-cn/library/dd319586(v=vs.85).aspx

class AppBarControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.AppBarControl)


class ButtonControl(Control, ExpandCollapsePattern, InvokePattern, TogglePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ButtonControl)


class CalendarControl(Control, GridPattern, TablePattern, ScrollPattern, SelectionPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.CalendarControl)


class CheckBoxControl(Control, TogglePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.CheckBoxControl)


class ComboBoxControl(Control, ExpandCollapsePattern, SelectionPattern, ValuePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ComboBoxControl)

    def Select(self, name, waitTime = OPERATION_WAIT_TIME):
        """not support Qt's ComboBoxControl"""
        supportExpandCollapse = self.IsExpandCollapsePatternAvailable()
        if supportExpandCollapse:
            self.Expand()
        else:
            #Windows Form's ComboBoxControl doesn't support ExpandCollapsePattern
            self.Click(-10, 0.5, False)
        find = False
        listItemControl = self.ListItemControl(Name = name)
        if listItemControl.Exists(1):
            listItemControl.ScrollIntoView()
            listItemControl.Click(waitTime = waitTime)
            find = True
        else:
            #ComboBox's popup window is a child of root control
            listControl = ListControl(searchDepth= 1)
            if listControl.Exists(1):
                listItemControl = listControl.ListItemControl(Name = name)
                if listItemControl.Exists(0, 0):
                    listItemControl.Click(waitTime = waitTime)
                    find = True
        if not find:
            Logger.ColorfulWriteLine('Can\'t find <Color=Cyan>{}</Color> in ComboBoxControl'.format(name), ConsoleColor.Yellow)
            if supportExpandCollapse:
                self.Collapse(waitTime)
            else:
                self.Click(-10, 0.5, False, waitTime)


class CustomControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.CustomControl)


class DataGridControl(Control, GridPattern, ScrollPattern, SelectionPattern, TablePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.DataGridControl)


class DataItemControl(Control, SelectionItemPattern, ExpandCollapsePattern, GridItemPattern, ScrollItemPattern, TableItemPattern, TogglePattern, ValuePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.DataItemControl)


class DocumentControl(Control, TextPattern, ScrollPattern, ValuePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.DocumentControl)


class EditControl(Control, RangeValuePattern, TextPattern, ValuePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.EditControl)


class GroupControl(Control, ExpandCollapsePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.GroupControl)


class HeaderControl(Control, TransformPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.HeaderControl)


class HeaderItemControl(Control, InvokePattern, TransformPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.HeaderItemControl)


class HyperlinkControl(Control, InvokePattern, ValuePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.HyperlinkControl)


class ImageControl(Control, GridItemPattern, TableItemPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ImageControl)


class ListControl(Control, GridPattern, MultipleViewPattern, ScrollPattern, SelectionPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ListControl)

    def GetSelectedItems(self):
        """Return a list of children which are selected"""
        lists = []
        iUIAutomationElementArray = self.GetCurrentSelection()
        if iUIAutomationElementArray:
            length = _AutomationClient.instance().dll.ElementArrayGetLength(iUIAutomationElementArray)
            for i in range(length):
                lists.append(Control.CreateControlFromElement(_AutomationClient.instance().dll.ElementArrayGetElement(iUIAutomationElementArray, i)))
            _AutomationClient.instance().dll.ReleaseElementArray(iUIAutomationElementArray)
        return lists


class ListItemControl(Control, SelectionItemPattern, ExpandCollapsePattern, GridItemPattern, InvokePattern, ScrollItemPattern, TogglePattern, ValuePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ListItemControl)


class MenuControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.MenuControl)


class MenuBarControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.MenuBarControl)


class MenuItemControl(Control, ExpandCollapsePattern, InvokePattern, SelectionItemPattern, TogglePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.MenuItemControl)


class PaneControl(Control, DockPattern, ScrollPattern, TransformPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.PaneControl)


class ProgressBarControl(Control, RangeValuePattern, ValuePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ProgressBarControl)


class RadioButtonControl(Control, SelectionItemPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.RadioButtonControl)


class ScrollBarControl(Control, RangeValuePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ScrollBarControl)


class SemanticZoomControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.SemanticZoomControl)


class SeparatorControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.SeparatorControl)


class SliderControl(Control, RangeValuePattern, SelectionPattern, ValuePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.SliderControl)


class SpinnerControl(Control, RangeValuePattern, SelectionPattern, ValuePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.SpinnerControl)


class SplitButtonControl(Control, ExpandCollapsePattern, InvokePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.SplitButtonControl)


class StatusBarControl(Control, GridPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.StatusBarControl)


class TabControl(Control, SelectionPattern, ScrollPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.TabControl)


class TabItemControl(Control, SelectionItemPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.TabItemControl)


class TableControl(Control, GridPattern, GridItemPattern, TablePattern, TableItemPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.TableControl)


class TextControl(Control, GridItemPattern, TableItemPattern, TextPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.TextControl)


class ThumbControl(Control, TransformPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ThumbControl)


class TitleBarControl(Control):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.TitleBarControl)


class ToolBarControl(Control, DockPattern, ExpandCollapsePattern, TransformPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ToolBarControl)


class ToolTipControl(Control, TextPattern, WindowPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.ToolTipControl)


class TreeControl(Control, ScrollPattern, SelectionPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.TreeControl)


class TreeItemControl(Control, ExpandCollapsePattern, InvokePattern, ScrollItemPattern, SelectionItemPattern, TogglePattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.TreeItemControl)


class WindowControl(Control, TransformPattern, WindowPattern, DockPattern):
    def __init__(self, element = 0, searchFromControl = None, searchDepth = 0xFFFFFFFF, searchWaitTime = SEARCH_INTERVAL, foundIndex = 1, **searchPorpertyDict):
        Control.__init__(self, element, searchFromControl, searchDepth, searchWaitTime, foundIndex, **searchPorpertyDict)
        self.AddSearchProperty(ControlType = ControlType.WindowControl)

    def SetTopmost(self, isTopmost = True):
        return Win32API.SetWindowTopmost(self.Handle, isTopmost)

    def SwitchToThisWindow(self):
        Win32API.SwitchToThisWindow(self.Handle)

    def MoveToCenter(self):
        left, top, right, bottom = self.BoundingRectangle
        width, height = right - left, bottom - top
        screenWidth, screenHeight = Win32API.GetScreenSize()
        x, y = (screenWidth-width)//2, (screenHeight-height)//2
        if x < 0: x = 0
        if y < 0: y = 0
        return Win32API.SetWindowPos(self.Handle, SWP.HWND_TOP, x, y, 0, 0, SWP.SWP_NOSIZE)

    def MetroClose(self, waitTime = OPERATION_WAIT_TIME):
        """only works in Windows 8/8.1, if current window is Metro UI"""
        window = WindowControl(searchDepth = 1, ClassName = METRO_WINDOW_CLASS_NAME)
        if window.Exists(0, 0):
            screenWidth, screenHeight = Win32API.GetScreenSize()
            Win32API.MouseMoveTo(screenWidth//2, 0, waitTime = 0)
            Win32API.MouseDragDrop(screenWidth//2, 0, screenWidth//2, screenHeight, waitTime = waitTime)
        else:
            Logger.WriteLine('Window is not Metro!', ConsoleColor.Yellow)

    def SetActive(self, waitTime = OPERATION_WAIT_TIME):
        curState = self.CurrentWindowVisualState()
        if curState == WindowVisualState.Minimized:
            self.ShowWindow(ShowWindow.Restore)
        elif curState == WindowVisualState.Maximized:
            self.ShowWindow(ShowWindow.Maximize)
        else:
            self.ShowWindow(ShowWindow.Show)
        ret = Win32API.SetForegroundWindow(self.Handle)  #maybe fail if foreground windows's process is not python
        time.sleep(waitTime)
        return ret


ControlDict = {
            ControlType.AppBarControl : AppBarControl,
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


class Logger:
    FileName = '@AutomationLog.txt'
    LineSep = '\n'
    ColorName2Value = {
        "Black"       : ConsoleColor.Black          ,
        "DarkBlue"    : ConsoleColor.DarkBlue       ,
        "DarkGreen"   : ConsoleColor.DarkGreen      ,
        "DarkCyan"    : ConsoleColor.DarkCyan       ,
        "DarkRed"     : ConsoleColor.DarkRed        ,
        "DarkMagenta" : ConsoleColor.DarkMagenta    ,
        "DarkYellow"  : ConsoleColor.DarkYellow     ,
        "Gray"        : ConsoleColor.Gray           ,
        "DarkGray"    : ConsoleColor.DarkGray       ,
        "Blue"        : ConsoleColor.Blue           ,
        "Green"       : ConsoleColor.Green          ,
        "Cyan"        : ConsoleColor.Cyan           ,
        "Red"         : ConsoleColor.Red            ,
        "Magenta"     : ConsoleColor.Magenta        ,
        "Yellow"      : ConsoleColor.Yellow         ,
        "White"       : ConsoleColor.White          ,
    }

    @staticmethod
    def SetLogFile(path):
        Logger.FileName = path

    @staticmethod
    def Write(log, consoleColor = -1, writeToFile = True, printToStdout = True, logFile = None):
        """
        consoleColor: value in class ConsoleColor, such as ConsoleColor.DarkGreen
        if consoleColor == -1, use default color
        """
        if printToStdout and sys.stdout:
            isValidColor = (consoleColor >= ConsoleColor.Black and consoleColor <= ConsoleColor.White)
            if isValidColor:
                Win32API.SetConsoleColor(consoleColor)
            try:
                sys.stdout.write(log)
            except Exception as ex:
                Win32API.SetConsoleColor(ConsoleColor.Red)
                isValidColor = True
                sys.stdout.write(ex.__class__.__name__ + ': can\'t print the log!')
                if log.endswith(Logger.LineSep):
                    sys.stdout.write(Logger.LineSep)
            if isValidColor:
                Win32API.ResetConsoleColor()
            if sys.stdout:
                sys.stdout.flush()
        if not writeToFile:
            return
        fileName = logFile if logFile else Logger.FileName
        try:
            fout = None
            if IsPy3:
                fout = open(fileName, 'a+', encoding = 'utf-8')
            else:
                fout = codecs.open(fileName, 'a+', 'utf-8')
            fout.write(log)
        except Exception as ex:
            if sys.stdout:
                sys.stdout.write(ex.__class__.__name__ + ': can\'t write the log!')
        finally:
            if fout:
                fout.close()

    @staticmethod
    def WriteLine(log, consoleColor = -1, writeToFile = True, printToStdout = True, logFile = None):
        """
        consoleColor: value in class ConsoleColor, such as ConsoleColor.DarkGreen
        if consoleColor == -1, use default color
        """
        Logger.Write(log + Logger.LineSep, consoleColor, writeToFile, printToStdout, logFile)

    @staticmethod
    def ColorfulWrite(log, consoleColor = -1, writeToFile = True, printToStdout = True, logFile = None):
        """ColorfulWrite('Hello <Color=Green>Green</Color> !!!'), color name must in Logger.ColorName2Value"""
        text = []
        start = 0
        while True:
            index1 = log.find('<Color=', start)
            if index1 >= 0:
                if index1 > start:
                    text.append((log[start:index1], consoleColor))
                index2 = log.find('>', index1)
                colorName = log[index1+7:index2]
                index3 = log.find('</Color>', index2 + 1)
                text.append((log[index2+1:index3], Logger.ColorName2Value[colorName]))
                start = index3 + 8
            else:
                if start < len(log):
                    text.append((log[start:], consoleColor))
                break
        for t, c in text:
            Logger.Write(t, c, writeToFile, printToStdout, logFile)

    @staticmethod
    def ColorfulWriteLine(log, consoleColor = -1, writeToFile = True, printToStdout = True, logFile = None):
        """ColorfulWriteLine('Hello <Color=Green>Green</Color> !!!'), color name must in Logger.ColorName2Value"""
        Logger.ColorfulWrite(log + Logger.LineSep, consoleColor, writeToFile, printToStdout, logFile)

    @staticmethod
    def Log(log = '', consoleColor = -1, writeToFile = True, printToStdout = True, logFile = None):
        """
        consoleColor: value in class ConsoleColor, such as ConsoleColor.DarkGreen
        if consoleColor == -1, use default color
        """
        t = datetime.datetime.now()
        frame = sys._getframe(1)
        log = '{}-{:02}-{:02} {:02}:{:02}:{:02}.{:03} Function: {}, Line: {} -> {}{}'.format(t.year, t.month, t.day,
            t.hour, t.minute, t.second, t.microsecond // 1000, frame.f_code.co_name, frame.f_lineno, log, Logger.LineSep)
        Logger.Write(log, consoleColor, writeToFile, printToStdout, logFile)

    @staticmethod
    def ColorfulLog(log = '', consoleColor = -1, writeToFile = True, printToStdout = True, logFile = None):
        """
        consoleColor: value in class ConsoleColor, such as ConsoleColor.DarkGreen
        if consoleColor == -1, use default color
        """
        t = datetime.datetime.now()
        frame = sys._getframe(1)
        log = '{}-{:02}-{:02} {:02}:{:02}:{:02}.{:03} Function: {}, Line: {} -> {}{}'.format(t.year, t.month, t.day,
            t.hour, t.minute, t.second, t.microsecond // 1000, frame.f_code.co_name, frame.f_lineno, log, Logger.LineSep)
        Logger.ColorfulWrite(log, consoleColor, writeToFile, printToStdout, logFile)

    @staticmethod
    def DeleteLog():
        if os.path.exists(Logger.FileName):
            os.remove(Logger.FileName)


def SetGlobalSearchTimeOut(seconds):
    global TIME_OUT_SECOND
    TIME_OUT_SECOND = seconds


def GetClipboardText():
    return Win32API.GetClipboardText()


def SetClipboardText(text):
    Win32API.SetClipboardText(text)


def Click(x, y, waitTime = OPERATION_WAIT_TIME):
    Win32API.MouseClick(x, y, waitTime)


def RightClick(x, y, waitTime = OPERATION_WAIT_TIME):
    Win32API.MouseRightClick(x, y, waitTime)


def MiddleClick(x, y, waitTime = OPERATION_WAIT_TIME):
    Win32API.MouseMiddleClick(x, y, waitTime)


def MoveTo(x, y, waitTime = OPERATION_WAIT_TIME):
    Win32API.MouseMoveTo(x, y, waitTime)


def WheelDown(wheelTimes = 1, waitTime = OPERATION_WAIT_TIME):
    Win32API.MouseWheelDown(wheelTimes, waitTime)


def WheelUp(wheelTimes = 1, waitTime = OPERATION_WAIT_TIME):
    Win32API.MouseWheelUp(wheelTimes, waitTime)


def DragDrop(x1, y1, x2, y2, waitTime = OPERATION_WAIT_TIME):
    Win32API.MouseDragDrop(x1, y1, x2, y2, 1, waitTime)


def KeyDown(key, waitTime = OPERATION_WAIT_TIME):
    """key: a value in class Keys"""
    Win32API.KeyDown(key, waitTime)


def KeyUp(key, waitTime = OPERATION_WAIT_TIME):
    """key: a value in class Keys"""
    Win32API.KeyUp(key, waitTime)


def SendKey(key, waitTime = OPERATION_WAIT_TIME):
    """
    Simulate typing a key
    key: a value in class Keys
    example: SendKey(automation.Keys.VK_F)
    """
    Win32API.SendKey(key, waitTime)


def SendKeys(keys, interval=0.01, waitTime = OPERATION_WAIT_TIME, debug=False):
    """
    Simulate typing keys on keyboard
    keys: str, keys to type
    interval: double, seconds between keys
    debug: bool, if True, print the Keys
    example:
    {Ctrl}, {Delete} ... are special keys' name in Win32API.SpecialKeyDict
    SendKeys('{Ctrl}a{Delete}{Ctrl}v{Ctrl}s{Ctrl}{Shift}s{Win}e{PageDown}') #press Ctrl+a, Delete, Ctrl+v, Ctrl+s, Ctrl+Shift+s, Win+e, PageDown
    SendKeys('{Ctrl}(AB)({Shift}(123))') #press Ctrl+A+B, type (, press Shift+1+2+3, type ), if () follows a hold key, hold key won't release util )
    SendKeys('{Ctrl}{a 3}') #press Ctrl+a at the same time, release Ctrl+a, then type a 2 times
    SendKeys('{a 3}{B 5}') #type a 3 times, type B 5 times
    SendKeys('{{}Hello{}}abc {a}{b}{c} test{} 3}{!}{a} (){(}{)}') #type: {Hello}abc abc test}}}!a ()()
    SendKeys('0123456789{Enter}')
    SendKeys('ABCDEFGHIJKLMNOPQRSTUVWXYZ{Enter}')
    SendKeys('abcdefghijklmnopqrstuvwxyz{Enter}')
    SendKeys('`~!@#$%^&*()-_=+{Enter}')
    SendKeys('[]{{}{}}\\|;:\'\",<.>/?{Enter}')
    """
    Win32API.SendKeys(keys, interval, waitTime, debug)


def WaitForExist(control, timeout):
    return control.Exists(timeout, 1)


def WaitForDisappear(control, timeout):
    return control.Disappears(timeout, 1)


def WalkTree(top, getChildrenFunc = None, getFirstChildFunc = None, getNextSiblingFunc = None, includeTop = False, maxDepth = 0xFFFFFFFF):
    """
    walking a tree not using recursive algorithm
    if getChildrenFunc is valid, ignore getFirstChildFunc and getNextSiblingFunc
        yield 3 items tuple: (item, depth, remain children count in current depth)
    if getChildrenFunc is not valid, using getFirstChildFunc and getNextSiblingFunc
        yield 2 items tuple: (item, depth)
    example:
    def GetDirChildren(dir):
        if os.path.isdir(dir):
            return [os.path.join(dir, it) for it in os.listdir(dir)]

    for it, depth, count in WalkTree('D:\\', getChildrenFunc= GetDirChildren):
        print(it, depth, count)
    """
    if maxDepth <= 0:
        return
    depth = 0
    if getChildrenFunc:
        if includeTop:
            yield top, 0, 0
        children = getChildrenFunc(top)
        childList = [children]
        while depth >= 0:   #or while childList:
            lastItems = childList[-1]
            if lastItems:
                yield lastItems[0], depth + 1, len(lastItems) - 1
                if depth + 1 < maxDepth:
                    children = getChildrenFunc(lastItems[0])
                    if children:
                        depth += 1
                        childList.append(children)
                del lastItems[0]
            else:
                del childList[depth]
                depth -= 1
    elif getFirstChildFunc and getNextSiblingFunc:
        if includeTop:
            yield top, 0
        child = getFirstChildFunc(top)
        childList = [child]
        while depth >= 0:  #or while childList:
            lastItem = childList[-1]
            if lastItem:
                yield lastItem, depth + 1
                child = getNextSiblingFunc(lastItem)
                childList[depth] = child
                if depth + 1 < maxDepth:
                    child = getFirstChildFunc(lastItem)
                    if child:
                        depth += 1
                        childList.append(child)
            else:
                del childList[depth]
                depth -= 1


def ControlsAreSame(control1, control2):
    """return 1 if control1 and control2 are the same control, otherwise return 0"""
    return _AutomationClient.instance().dll.CompareElements(control1.Element, control2.Element)


def GetRootControl():
    global _rootControl
    if not _rootControl:
        _rootControl = Control.CreateControlFromElement(_AutomationClient.instance().dll.GetRootElement())
    return _rootControl


def GetFocusedControl():
    return Control.CreateControlFromElement(_AutomationClient.instance().dll.GetFocusedElement())


def GetForegroundControl():
    """return Foreground Window"""
    return ControlFromHandle(Win32API.GetForegroundWindow())
    #another implement
    #focusedControl = GetFocusedControl()
    #parentControl = focusedControl
    #controlList = []
    #while parentControl:
        #controlList.insert(0, parentControl)
        #parentControl = parentControl.GetParentControl()
    #if len(controlList) == 1:
        #parentControl = controlList[0]
    #else:
        #parentControl = controlList[1]
    #return parentControl


def GetConsoleWindow():
    """return console window that runs python"""
    title = Win32API.GetConsoleTitle()
    consoleWindow = WindowControl(searchDepth= 1, Name = title)
    if consoleWindow.Exists(0, 0):
        return consoleWindow


def ControlFromPoint(x, y):
    """use IUIAutomation ElementFromPoint x,y, may return 0 if mouse is over cmd's title bar icon"""
    element = _AutomationClient.instance().dll.ElementFromPoint(x, y)
    return Control.CreateControlFromElement(element)


def ControlFromPoint2(x, y):
    """use Win32API.WindowFromPoint x,y"""
    return Control.CreateControlFromElement(_AutomationClient.instance().dll.ElementFromHandle(Win32API.WindowFromPoint(x, y)))


def ControlFromCursor():
    x, y = Win32API.GetCursorPos()
    return ControlFromPoint(x, y)


def ControlFromCursor2():
    x, y = Win32API.GetCursorPos()
    return ControlFromPoint2(x, y)


def ControlFromHandle(handle):
    return Control.CreateControlFromElement(_AutomationClient.instance().dll.ElementFromHandle(handle))


def WalkControl(control, includeTop = False, maxDepth = 0xFFFFFFFF):
    """
    control: Control
    maxDepth: integer
    yield 2 items tuple(Control, depth)
    """
    if includeTop:
        yield control, 0
    if maxDepth <= 0:
        return
    depth = 0
    child = control.GetFirstChildControl()
    controlList = [child]
    while depth >= 0:
        lastControl = controlList[-1]
        if lastControl:
            yield lastControl, depth + 1
            child = lastControl.GetNextSiblingControl()
            controlList[depth] = child
            if depth + 1 < maxDepth:
                child = lastControl.GetFirstChildControl()
                if child:
                    depth += 1
                    controlList.append(child)
        else:
            del controlList[depth]
            depth -= 1


def LogControl(control, depth = 0, showAllName = True, showMore = False):
    """
    control: Control
    depth: integer
    showAllName: bool
    showMore: bool
    """
    def getKeyName(theDict, theValue):
        for key in theDict:
            if theValue == theDict[key]:
                return key
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
    Logger.Write('0x{0:X}({0})'.format(control.Handle), ConsoleColor.DarkGreen)
    Logger.Write('    Depth: ')
    Logger.Write(str(depth), ConsoleColor.DarkGreen)
    if ((isinstance(control, ValuePattern) and control.IsValuePatternAvailable())):
        Logger.Write('    Value: ')
        Logger.Write(control.CurrentValue(), ConsoleColor.DarkGreen)
    if ((isinstance(control, RangeValuePattern) and control.IsRangeValuePatternAvailable())):
        Logger.Write('    RangeValue: ')
        Logger.Write(str(control.RangeValuePatternCurrentValue()), ConsoleColor.DarkGreen)
    if isinstance(control, TogglePattern) and control.IsTogglePatternAvailable():
        Logger.Write('    CurrentToggleState: ')
        Logger.Write('ToggleState.' + getKeyName(ToggleState.__dict__, control.CurrentToggleState()), ConsoleColor.DarkGreen)
    if isinstance(control, SelectionItemPattern) and control.IsSelectionItemPatternAvailable():
        Logger.Write('    CurrentIsSelected: ')
        Logger.Write(str(control.CurrentIsSelected()), ConsoleColor.DarkGreen)
    if isinstance(control, ExpandCollapsePattern) and control.IsExpandCollapsePatternAvailable():
        Logger.Write('    CurrentExpandCollapseState: ')
        Logger.Write('ExpandCollapseState.' + getKeyName(ExpandCollapseState.__dict__, control.CurrentExpandCollapseState()), ConsoleColor.DarkGreen)
    if isinstance(control, ScrollPattern) and control.IsScrollPatternAvailable():
        Logger.Write('    CurrentHorizontalViewSize: ')
        Logger.Write(str(control.CurrentHorizontalViewSize()), ConsoleColor.DarkGreen)
        Logger.Write('    CurrentVerticalViewSize: ')
        Logger.Write(str(control.CurrentVerticalViewSize()), ConsoleColor.DarkGreen)
        Logger.Write('    CurrentHorizontalScrollPercent: ')
        Logger.Write(str(control.CurrentHorizontalScrollPercent()), ConsoleColor.DarkGreen)
        Logger.Write('    CurrentVerticalScrollPercent: ')
        Logger.Write(str(control.CurrentVerticalScrollPercent()), ConsoleColor.DarkGreen)
    if isinstance(control, GridPattern) and control.IsGridPatternAvailable():
        Logger.Write('    RowCount: ')
        Logger.Write(str(control.CurrentRowCount()), ConsoleColor.DarkGreen)
        Logger.Write('    ColumnCount: ')
        Logger.Write(str(control.CurrentColumnCount()), ConsoleColor.DarkGreen)
    if isinstance(control, GridItemPattern) and control.IsGridItemPatternAvailable():
        Logger.Write('    Row: ')
        Logger.Write(str(control.CurrentRow()), ConsoleColor.DarkGreen)
        Logger.Write('    Column: ')
        Logger.Write(str(control.CurrentColumn()), ConsoleColor.DarkGreen)
    if showMore:
        Logger.Write('    SupportedPattern:')
        for key in PatternDict:
            pattern = _AutomationClient.instance().dll.GetElementPattern(control.Element, key)
            if pattern:
                _AutomationClient.instance().dll.ReleasePattern(pattern)
                Logger.Write(' ' + PatternDict[key], ConsoleColor.DarkGreen)
    Logger.Write(Logger.LineSep)


def EnumAndLogControlAncestors(control, showAllName = True, showMore = False):
    lists = []
    while control:
        lists.insert(0, control)
        control = control.GetParentControl()
    for (i, control) in enumerate(lists):
        LogControl(control, i, showAllName, showMore)


def EnumAndLogControl(control, maxDepth = 0xFFFFFFFF, showAllName = True, showMore = False):
    """
    control: Control
    maxDepth: integer
    showAllName: bool
    showMore: bool
    """
    for c, d in WalkControl(control, True, maxDepth):
        LogControl(c, d, showAllName, showMore)


def FindControl(control, compareFunc, maxDepth = 0xFFFFFFFF, findFromSelf = False, foundIndex = 1):
    """
    control: Control
    compareFunc: compare function with parameters (control, depth), should return True or False
    maxDepth: integer
    findFromSelf: bool
    foundIndex: integer, value must be greater or equal to 1
    """
    foundCount = 0
    if not control:
        control = GetRootControl()
    for child, depth in WalkControl(control, findFromSelf, maxDepth):
        if compareFunc(child, depth):
            foundCount += 1
            if foundCount == foundIndex:
                return child


def ShowDesktop():
    """show the desktop by win + d"""
    SendKeys('{Win}d')
    time.sleep(1)
    #another implement
    #paneTray = PaneControl(searchDepth = 1, ClassName = 'Shell_TrayWnd')
    #if paneTray.Exists():
        #WM_COMMAND = 0x111
        #MIN_ALL = 419
        #MIN_ALL_UNDO = 416
        #Win32API.PostMessage(paneTray.Handle, WM_COMMAND, MIN_ALL, 0)
        #time.sleep(1)


def RunWithHotKey(keyFunctionDict, stopHotKey = None):
    """
    keyFunctionDict: hotkey, function dict, like {(automation.ModifierKey.MOD_CONTROL, automation.Keys.VK_1) : func}
    bind function with hotkey, the function will be run or stopped in another thread when the hotkey was pressed
    automation doesn't support multi thread, so you can't use UI Control in the function
    you can call another script that uses UI Control

    def main(stopEvent):
        n = 0
        while True:
            if stopEvent.is_set(): # must check stopEvent.is_set() if you want to stop when stop hot key was pressed
                break
            print(n)
            n += 1
            stopEvent.wait(1)
        print('main exit')
        print(automation.GetRootControl())      # will raise exception, can't use UI Control, todo

    automation.RunHotKey({(automation.ModifierKey.MOD_CONTROL, automation.Keys.VK_1) : main}
                        , (automation.ModifierKey.MOD_CONTROL | automation.ModifierKey.MOD_SHIFT, automation.Keys.VK_2))
    """
    stopHotKeyId = 1
    exitHotKeyId = 2
    hotKeyId = 3
    registed = True
    def getModName(theDict, theValue):
        name = ''
        for key in theDict:
            if isinstance(theDict[key], int) and theValue & theDict[key]:
                if name:
                    name += '|'
                name += key
        return name
    def getKeyName(theDict, theValue):
        for key in theDict:
            if theValue == theDict[key]:
                return key
    def releaseAllKey():
        for key, value in Keys.__dict__.items():
            if isinstance(value, int) and key.startswith('VK'):
                if Win32API.IsKeyPressed(value):
                    Win32API.ReleaseKey(value)

    id2HotKey = {}
    id2Function = {}
    id2Thread = {}
    id2Name = {}
    for hotkey in keyFunctionDict:
        id2HotKey[hotKeyId] = hotkey
        id2Function[hotKeyId] = keyFunctionDict[hotkey]
        id2Thread[hotKeyId] = None
        modName = getModName(ModifierKey.__dict__, hotkey[0])
        keyName = getKeyName(Keys.__dict__, hotkey[1])
        id2Name[hotKeyId] = str((modName, keyName))
        if ctypes.windll.user32.RegisterHotKey(0, hotKeyId, hotkey[0], hotkey[1]):
            Logger.ColorfulWriteLine('Register hotKey <Color=DarkGreen>{}</Color> succeed'.format((modName, keyName)), writeToFile = False)
        else:
            registed = False
            Logger.ColorfulWriteLine('Register hotKey <Color=DarkGreen>{}</Color> failed, maybe it was allready registered by another program'.format((modName, keyName)), writeToFile = False)
        hotKeyId += 1
    if stopHotKey and len(stopHotKey) == 2:
        modName = getModName(ModifierKey.__dict__, stopHotKey[0])
        keyName = getKeyName(Keys.__dict__, stopHotKey[1])
        if ctypes.windll.user32.RegisterHotKey(0, stopHotKeyId, stopHotKey[0], stopHotKey[1]):
            Logger.ColorfulWriteLine('Register stop hotKey <Color=Yellow>{}</Color> succeed'.format((modName, keyName)), writeToFile = False)
        else:
            registed = False
            Logger.ColorfulWriteLine('Register stop hotKey <Color=Yellow>{}</Color> failed, maybe it was allready registered by another program'.format((modName, keyName)), writeToFile = False)
    if not registed:
        return
    if ctypes.windll.user32.RegisterHotKey(0, exitHotKeyId, ModifierKey.MOD_CONTROL, Keys.VK_D):
        Logger.ColorfulWriteLine('Register <Color=Yellow>Ctrl+D</Color> succeed, for exiting current script', writeToFile = False)
    else:
        Logger.ColorfulWriteLine('Register <Color=Yellow>Ctrl+D</Color> failed', writeToFile = False)
    from threading import Thread, Event
    funcThread = None
    stopEvent = Event()
    msg = MSG()
    def threadFunc(function, stopEvent):
        function(stopEvent)
        releaseAllKey()  #need to release keys if some keys were pressed
    while ctypes.windll.user32.GetMessageW(ctypes.byref(msg), 0, 0, 0) != 0:
        if msg.message == 0x0312: # WM_HOTKEY=0x0312
            if msg.wParam in id2HotKey:
                if msg.lParam&0x0000FFFF == id2HotKey[msg.wParam][0] and msg.lParam>>16&0x0000FFFF == id2HotKey[msg.wParam][1]:
                    Logger.ColorfulWriteLine('----------hotkey <Color=DarkGreen>{}</Color> pressed----------'.format(id2Name[msg.wParam]), writeToFile = False)
                    if not id2Thread[msg.wParam]:
                        stopEvent.clear()
                        funcThread=Thread(None, threadFunc, args = (id2Function[msg.wParam], stopEvent))
                        funcThread.setDaemon(True)
                        funcThread.start()
                        id2Thread[msg.wParam] = funcThread
                    else:
                        if not id2Thread[msg.wParam].isAlive():
                            stopEvent.clear()
                            funcThread=Thread(None, threadFunc, args = (id2Function[msg.wParam], stopEvent))
                            funcThread.setDaemon(True)
                            funcThread.start()
                            id2Thread[msg.wParam] = funcThread
                        else:
                            Logger.WriteLine('There is a thread that had already run for hotkey {}'.format(id2Name[msg.wParam]), ConsoleColor.Yellow, writeToFile = False)
            elif stopHotKeyId == msg.wParam:
                if msg.lParam&0x0000FFFF == stopHotKey[0] and msg.lParam>>16&0x0000FFFF == stopHotKey[1]:
                    Logger.WriteLine('----------stop hotkey pressed----------', writeToFile = False)
                    stopEvent.set()
                    for id in id2Thread:
                        id2Thread[id] = None
            elif exitHotKeyId == msg.wParam:
                if msg.lParam&0x0000FFFF == ModifierKey.MOD_CONTROL and msg.lParam>>16&0x0000FFFF == Keys.VK_D:
                    Logger.WriteLine('Ctrl+D pressed. Exit', ConsoleColor.Yellow, False)
                    break
