#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys, xlrd, signal, random, os, time

def signal_handler(signal, frame):
  print("\nYou pressed Ctrl+C!")
  sys.exit(0)

signal.signal(signal.SIGINT,signal_handler)

def clear_terminal(t):
  time.sleep(t)
  if os.name == "posix":  # Unix-like
    os.system("clear")
  else:
    os.system("cls")

def read_excel():
  # check the arguments
  idx, rflag, arglen = 0, False, len(sys.argv)
  if arglen < 2:
    print("Usage:", sys.argv[0], "excelfile [sheet_index=0] [random=False]")
    exit(1)
  elif arglen > 3:
    rflag = bool(int(sys.argv[3]))
  elif arglen > 2:
    idx = int(sys.argv[2]) - 1

  # open excel file and read the specified sheet
  ExcelFile = xlrd.open_workbook(sys.argv[1]) 
  sheet = ExcelFile.sheet_by_index(idx)

  # print sheet's name and rows
  print("Sheet.Name:", sheet.name, " Sheet.Rows:", sheet.nrows, " Random:", rflag)

  # use random shuttle or not
  nrange = range(sheet.nrows)
  if rflag:
    nrange = list(nrange)
    random.shuffle(nrange)

  # print content if the column 2 is not empty
  erows = []
  for i in nrange:
    # ignore the first row
    if i == 0:
      continue
    # if input is wrong word, put it into the loop cycle array
    if sheet.row_values(i)[1] != "":
      print(sheet.row_values(i)[3], end='')
      kdc = input(" : ")
      if sheet.row_values(i)[1] == kdc:
        print("Good!")
      else:
        erows.append(sheet.row_values(i))
        print("No!!!", erows[-1][1], erows[-1][2])

  # repeat until everything is ok
  i = 0
  clear_terminal(0)
  while len(erows) > 0:
    print(erows[i][3], end='')
    kdc = input(" : ")
    if erows[i][1] == kdc:
      erows.pop(i)
      print("Good!")
    else:
      print("No!!!", erows[i][1], erows[i][2])
      i += 1
  
    if i >= len(erows):
      i = 0
      clear_terminal(1)
      

  print("All done! Good job!!!")

if __name__ == '__main__':
  read_excel()

