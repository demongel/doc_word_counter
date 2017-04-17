#!/usr/bin/env python
# -*- coding: utf-8-*-

import os
import win32com.client as win32

# create list of files (exclude temporarly word files (starts with '~$')
def walk(dir):
    for name in os.listdir(dir):
        path = os.path.join(dir, name)
        fileslist = open('fileslist.txt', 'a+')
        if os.path.isfile(path):
            if name.startswith('~$'):
                os.remove(path)
            else:
#                 print isinstance(path,str)  # 判断类型
#                 print path.decode('gbk').encode("utf-8");# str 进行decode  unicode进行encode
                fileslist.write(path.decode('gbk').encode("utf-8")) #存入正确的中文路径
                fileslist.write('\n')
        else:
            walk(path)


# extract statistics from .doc files
def wrdreader(path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    word1 = word.Documents.Open(path)
    word1.Visible = False
    text = word1.Content
    # counts.txt - a list of lines, where a line have the next format: file with filename + number of word statistic
    filescount = open('counts.txt', 'a+')
    # numbers.txt - a line of numbers, separated by a space
    numberscount = open('numbers.txt', 'a+')
    # Below we are used .ComputeStatistics() method with '5' as an argument. The next variations can be used: '0' - StatisticWords, '1' - StatisticLines', '2' - StatisticPages, '3' - StatisticCharacters, '4' - StatisticParagraphs, '5' - StatisticCharactersWithSpaces
    # 参数3 不含空格  参数5 含空格
    filescount.write(path.decode('gbk').encode("utf-8") + ' ' + str(text.ComputeStatistics(3)) + '\n')
    numberscount.write(str(text.ComputeStatistics(3)) + ' ')
    print(path.decode('gbk').encode("utf-8"))  #打印正确的中文
    print(text.ComputeStatistics(3))
    filescount.close
    numberscount.close
    word1.Close()
  
  

def countr():
    numberscount = open('numbers.txt')
    arrnumbers = numberscount.read().strip()
    summ = 0
    for i in arrnumbers.split(' '):
        if type(int(i)) == int:
            summ += int(i)   
    numberscount.close
    numberscount = open('numbers.txt', 'a+')
    numberscount.write('\nTotal: ' + str(summ))
    print(summ)

  
def main():
# for the correct operation of the program need to specify the directory of word files (program can correctly process, including the situation when files are placed in subfolders)
    path = 'E:\count'   # 需要统计的文件夹
# remove temp files for earlier execution of this app
    for file in ['counts.txt', 'fileslist.txt', 'numbers.txt']:
        if os.path.isfile(file): 
            os.remove(file)
    walk(path)  # 将路径写入一个文本中。。？
    fileslist = open('fileslist.txt', 'r')  # 读取该文件  统计信息
    links_to_all_files = fileslist.read()
    paths = links_to_all_files.split('\n')
    for path in paths:
        if path != '':
            wrdreader(path.decode('utf-8').encode("gbk")) # 读取时 重新转换成gbk编码，否则不能正确读取路径。
    countr()

if __name__ == '__main__':
    main()
