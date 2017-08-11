#!python3
# -*- coding: utf-8 -*-
#author: yinkaisheng@foxmail.com
"""
How to use:
Wireshark version must >= 2.0
run Wireshark and start capture
use VLC to play a rtsp url
stop capture
run this script
"""

import os
import sys
import time
import copy

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as automation


class PacketInfo():
    """Class Packet Info"""

    def __init__(self):
        """Constructor"""
        self.No = 0
        self.Time = 0.0
        self.Source = ''
        self.Destination = ''
        self.Protocol = ''
        self.Length = 0
        self.Info = ''
        self.TimeStamp = 0
        self.Seq = 0


def AnalyzeUI(sampleRate = 90000, payload = 96, beginNo = 0, maxPackets = 0xFFFFFFFF, calculateLost = False):
    """Wireshark version must >= 2.0"""
    wireSharkWindow = automation.WindowControl(searchDepth= 1, ClassName = 'Qt5QWindowIcon')
    if wireSharkWindow.Exists(0, 0):
        wireSharkWindow.SetActive()
    else:
        automation.Logger.WriteLine('can not find wireshark', automation.ConsoleColor.Yellow)
        return
    tree = wireSharkWindow.TreeControl(searchDepth= 4, SubName = 'Packet list')
    left, top, right, bottom = tree.BoundingRectangle
    tree.Click(10, 30)
    automation.SendKeys('{Home}{Ctrl}{Alt}4')
    time.sleep(0.5)
    tree.Click(10, 30)
    headers = []
    headerFunctionDict = {'No': int,
                          'Time': float,
                          'Source': str,
                          'Destination': str,
                          'Protocol': str,
                          'Length': int,
                          'Info': str,}
    index = 0
    payloadStr = 'PT=DynamicRTP-Type-' + str(payload)
    packets = []
    for item, depth in automation.WalkTree(tree
                                           , getFirstChildFunc= lambda c:c.GetFirstChildControl()
                                           , getNextSiblingFunc= lambda c:c.GetNextSiblingControl()):
        if isinstance(item, automation.HeaderControl):
            headers.append(item.Name.rstrip('.'))
        elif isinstance(item, automation.TreeItemControl):
            if index == 0:
                if len(packets) >= maxPackets:
                    break
                packet = PacketInfo()
            name = item.Name
            packet.__dict__[headers[index]] = headerFunctionDict[headers[index]](name)
            if headers[index] == 'Info':
                findPayload = True
                while findPayload:
                    ptIndex = name.find('PT=DynamicRTP-Type', 1)
                    if ptIndex > 0:
                        info = name[:ptIndex]
                        name = name[ptIndex:]
                        findPayload = True
                    else:
                        info = name
                        findPayload = False
                    packet.Info = info
                    if info.find(payloadStr) >= 0:
                        packet = copy.copy(packet)
                        startIndex = info.find('Seq=')
                        if startIndex > 0:
                            endIndex = info.find(' ', startIndex)
                            packet.Seq = int(info[startIndex+4:endIndex].rstrip(','))
                        startIndex = info.find('Time=', startIndex)
                        if startIndex > 0:
                            endIndex = info.find(' ', startIndex)
                            packet.TimeStamp = int(info[startIndex+5:endIndex].rstrip(','))
                            if packet.No >= beginNo:
                                packets.append(packet)
                                automation.Logger.WriteLine('No: {0[No]:<10}, Time: {0[Time]:<10}, Protocol: {0[Protocol]:<6}, Length: {0[Length]:<6}, Info: {0[Info]:<10},'.format(packet.__dict__))
            index = (index + 1) % len(headers)
            if item.BoundingRectangle[3] >= bottom:
                automation.SendKeys('{PageDown}')
                time.sleep(0.1)
    AnalyzePackets(packets, sampleRate, calculateLost)

def AnalyzeCsvFile(csvFile, sampleRate = 90000, payload = 96, beginNo = 0, maxPackets = 0xFFFFFFFF, calculateLost = False):
    """Analyze Wireshark export csv file"""
    headers = []
    headerFunctionDict = {'No': int,
                          'Time': float,
                          'Source': str,
                          'Destination': str,
                          'Protocol': str,
                          'Length': int,
                          'Info': str,}
    index = 0
    payloadStr = 'PT=DynamicRTP-Type-' + str(payload)
    packets = []
    with open(csvFile) as fin:
        for line in fin:
            line = line.strip().strip('"')
            if line.startswith('No'):
                headers = line.split('","')
                headers[0] = headers[0].rstrip('.')
            elif line:
                for name in line.split('","'):
                    if index == 0:
                        if len(packets) >= maxPackets:
                            break
                        packet = PacketInfo()
                    packet.__dict__[headers[index]] = headerFunctionDict[headers[index]](name)
                    if headers[index] == 'No' and name == '20':
                        findBreakPoint = 1  #for debug
                    if headers[index] == 'Info':
                        findPayload = True
                        while findPayload:
                            ptIndex = name.find('PT=DynamicRTP-Type', 1)
                            if ptIndex > 0:
                                info = name[:ptIndex]
                                name = name[ptIndex:]
                                findPayload = True
                            else:
                                info = name
                                findPayload = False
                            packet.Info = info
                            if info.find(payloadStr) >= 0:
                                packet = copy.copy(packet)
                                startIndex = info.find('Seq=')
                                if startIndex > 0:
                                    endIndex = info.find(' ', startIndex)
                                    packet.Seq = int(info[startIndex+4:endIndex].rstrip(','))
                                startIndex = info.find('Time=', startIndex)
                                if startIndex > 0:
                                    endIndex = info.find(' ', startIndex)
                                    packet.TimeStamp = int(info[startIndex+5:endIndex].rstrip(','))
                                    if packet.No >= beginNo:
                                        packets.append(packet)
                                        automation.Logger.WriteLine('No: {0[No]:<10}, Time: {0[Time]:<10}, Protocol: {0[Protocol]:<6}, Length: {0[Length]:<6}, Info: {0[Info]:<10},'.format(packet.__dict__))
                    index = (index + 1) % len(headers)
    AnalyzePackets(packets, sampleRate, calculateLost)

def AnalyzePackets(packets, sampleRate, calculateLost):
    automation.Logger.WriteLine('\n----------\nAnalyze Result:')
    seq = packets[0].Seq - 1
    lostSeqs = []
    framePackets = []
    lastTimeStamp = -1
    for p in packets:
        seq += 1
        if seq != p.Seq:
            lostSeqs.extend(range(seq, p.Seq))
            seq = p.Seq
        if lastTimeStamp < 0:
            framePackets.append(p)
        else:
            if lastTimeStamp == p.TimeStamp:
                framePackets[-1] = p
            else:
                framePackets.append(p)
        lastTimeStamp = p.TimeStamp

    lastTimeStamp = -1
    lastTime = -1
    totalDiff = 0
    frameCount = 0
    for p in framePackets:
        if lastTimeStamp < 0:
            automation.Logger.WriteLine('No: {0[No]:<8}, Seq: {0[Seq]:<8}, Time: {0[Time]:<10}, Protocol: {0[Protocol]:<4}, Length: {0[Length]:<6}, TimeStamp: {0[TimeStamp]:<15}'.format(p.__dict__))
        else:
            diffTimeStamp = p.TimeStamp - lastTimeStamp
            frameCount += 1
            totalDiff += diffTimeStamp
            sumTime = p.Time - framePackets[0].Time
            sumTimeFromTimeStamp = (p.TimeStamp - framePackets[0].TimeStamp) / sampleRate
            automation.Logger.WriteLine('No: {0[No]:<8}, Seq: {0[Seq]:<8}, Time: {0[Time]:<10}, Protocol: {0[Protocol]:<4}, Length: {0[Length]:<6}, TimeStamp: {0[TimeStamp]:<15}, 帧时间戳差值: {1:<6}, 帧实际时间差值: {2:<10.6f}, 接收总时间: {3:<10.6f}, 时间戳总时间: {4:<10.6f}, 提前时间：{5:<10.6f}'.format(
                p.__dict__, diffTimeStamp, p.Time-lastTime, sumTime, sumTimeFromTimeStamp, sumTimeFromTimeStamp - sumTime))
        lastTime = p.Time
        lastTimeStamp = p.TimeStamp
    if frameCount:
        averageDiff = totalDiff // frameCount
        frameCount += 1
        seqCount = packets[-1].Seq - packets[0].Seq + 1
        totalTimeFromTimeStamp = (framePackets[-1].TimeStamp - framePackets[0].TimeStamp) / sampleRate
        realTime = framePackets[-1].Time - framePackets[0].Time
        if calculateLost:
            automation.Logger.WriteLine('\n包总数: {0}, 帧数: {1}, 实际帧率:{2:.2f}, 平均时间戳: {3}, 接收总时间: {4:.6f}, 时间戳总时间: {5:.6f}, 丢包率: {6}({7:.2f}%), 丢包序号: \n{8}'.format(
                                    len(packets), frameCount, frameCount / realTime, averageDiff, realTime, totalTimeFromTimeStamp, len(lostSeqs), len(lostSeqs) / seqCount, '\n'.join(lostSeqs)))
        else:
            automation.Logger.WriteLine('\n包总数: {0}, 帧数: {1}, 实际帧率:{2:.2f}, 平均时间戳: {3}, 接收总时间: {4:.6f}, 时间戳总时间: {5:.6f}'.format(
                                    len(packets), frameCount, frameCount / realTime, averageDiff, realTime, totalTimeFromTimeStamp))


if __name__ == '__main__':
    if not automation.IsPy3:
        input = raw_input
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('-f', '--file', type = str, dest = 'file', default = '',
                        help = 'wireshark export csv file')
    parser.add_argument('-s', '--sample', type = int, dest = 'sampleRate', default = 0,
                      help = 'sample rate')
    parser.add_argument('-p', '--payload', type = int, dest = 'payload', default = 0,
                        help = 'payload type')
    parser.add_argument('-b', '--begin', type = int, dest = 'beginNo', default = 0,
                      help = 'begin no')
    parser.add_argument('-n', '--num', type = int, dest = 'maxPacket', default = 0xFFFFFFFF,
                      help = 'read packets count')
    parser.add_argument('-l', '--lost', type = bool, dest = 'calculateLost', default = False,
                          help = 'calculate lost packets')
    args = parser.parse_args()
    cmdWindow = automation.GetConsoleWindow()
    if cmdWindow:  #if run by a debugger, maybe no console window
        cmdWindow.SetTopmost()
    if 0 == args.sampleRate:
        args.sampleRate = int(input('please input sample rate: '))
    if 0 == args.payload:
        args.payload = int(input('please input playload type: '))
    if args.file:
        AnalyzeCsvFile(args.file, args.sampleRate, args.payload, args.beginNo, args.maxPacket, args.calculateLost)
    else:
        AnalyzeUI(args.sampleRate, args.payload, args.beginNo, args.maxPacket, args.calculateLost)
    if cmdWindow:
        cmdWindow.SetActive()
    input('press enter to exit')
