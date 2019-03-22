#!python3
# -*- coding: utf-8 -*-
import os
import sys
import time

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as automation

def WalkDesktop():
    def GetFirstChild(control):
        return control.GetFirstChildControl()

    def GetNextSibling(control):
        return control.GetNextSiblingControl()

    desktop = automation.GetRootControl()
    for control, depth in automation.WalkTree(desktop, getFirstChild=GetFirstChild, getNextSibling=GetNextSibling, includeTop=True, maxDepth=2):
        print(' ' * depth * 4 + str(control))

def WalkQueens(queenCount = 8):
    def GetNextQueenColumns(place):
        notValid = set()
        nextRow = len(place)
        for i in range(nextRow):
            leftDiagonal = place[i] - (nextRow - i)
            if leftDiagonal >= 0:
                notValid.add(leftDiagonal)
            notValid.add(place[i])
            rightDiagonal = place[i] + (nextRow - i)
            if rightDiagonal < queenCount:
                notValid.add(rightDiagonal)
        return [x for x in range(queenCount) if x not in notValid]

    def GetNextQueens(place):
        places = []
        if len(place) == 0:
            for i in range(queenCount):
                places.append([i])
        else:
            valid = GetNextQueenColumns(place)
            for x in valid:
                newPlace = place[:]
                newPlace.append(x)
                places.append(newPlace)
        return places

    count = 0
    for queens, depth, remain in automation.WalkTree([], GetNextQueens, yieldCondition=lambda c, d: d == queenCount):
        count += 1
        automation.Logger.WriteLine(count, automation.ConsoleColor.Cyan)
        for x in queens:
            automation.Logger.ColorfullyWriteLine('o ' * x + '<Color=Cyan>x </Color>' + 'o ' * (queenCount - x - 1))

def WalkPermutations(uniqueItems):
    def NextPermutations(aTuple):
        left, permutation = aTuple
        ret = []
        for i, item in enumerate(left):
            nextLeft = left[:]
            del nextLeft[i]
            nextPermutation = permutation + [item]
            ret.append((nextLeft, nextPermutation))
        return ret

    n = len(uniqueItems)
    count = 0
    for (left, permutation), depth, remain in automation.WalkTree((uniqueItems, []), NextPermutations, yieldCondition=lambda c, d: d == n):
        count += 1
        print(count, permutation)

def main():
    automation.Logger.WriteLine('\nwalk desktop for 2 depth', automation.ConsoleColor.Cyan)
    time.sleep(1)
    WalkDesktop()
    automation.Logger.WriteLine('\nwalk 8 queens', automation.ConsoleColor.Cyan)
    time.sleep(1)
    WalkQueens(8)
    automation.Logger.WriteLine('\nwalk permutations for "abc"', automation.ConsoleColor.Cyan)
    time.sleep(1)
    WalkPermutations(list("abc"))

if __name__ == '__main__':
    main()
    input('press enter to exit')
