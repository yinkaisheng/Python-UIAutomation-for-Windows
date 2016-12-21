#!python3
# -*- coding: utf-8 -*-

import time
import uiautomation as automation

def main():
    firefoxWindow = automation.WindowControl(searchDepth = 1, ClassName = 'MozillaWindowClass')
    if not firefoxWindow.Exists(0):
        automation.Logger.WriteLine('please run Firefox first', automation.ConsoleColor.Yellow)
        return
    firefoxWindow.ShowWindow(automation.ShowWindow.Maximize)
    firefoxWindow.SetActive()
    time.sleep(1)
    tab = firefoxWindow.TabControl()
    newTabButton = tab.ButtonControl(searchDepth= 1)
    newTabButton.Click()
    edit = firefoxWindow.EditControl()
    # edit.Click()
    edit.SendKeys('http://global.bing.com/?rb=0&setmkt=en-us&setlang=en-us{Enter}')
    time.sleep(2)
    searchEdit = automation.FindControl(firefoxWindow,
                           lambda c:
                           (isinstance(c, automation.EditControl) or isinstance(c, automation.ComboBoxControl)) and c.Name == 'Enter your search term'
                           )
    # searchEdit.Click()
    searchEdit.SendKeys('Python-UIAutomation-for-Windows site:github.com{Enter}', 0.05)
    link = firefoxWindow.HyperlinkControl(SubName = 'yinkaisheng/Python-UIAutomation-for-Windows')
    automation.Win32API.PressKey(automation.Keys.VK_CONTROL)
    link.Click()  #press control to open the page in a new tab
    automation.Win32API.ReleaseKey(automation.Keys.VK_CONTROL)
    newTab = tab.TabItemControl(SubName = 'yinkaisheng/Python-UIAutomation-for-Windows')
    newTab.Click()
    starButton = firefoxWindow.ButtonControl(Name = 'Star this repository')
    if starButton.Exists(5):
        automation.GetConsoleWindow().SetActive()
        automation.Logger.WriteLine('Star Python-UIAutomation-for-Windows after 2 seconds', automation.ConsoleColor.Yellow)
        time.sleep(2)
        firefoxWindow.SetActive()
        time.sleep(1)
        starButton.Click()
        time.sleep(2)

if __name__ == '__main__':
    main()
    automation.GetConsoleWindow().SetActive()
    input('press Enter to exit')
