#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys, xlrd 

def read_excel():
  # open excel file and read the first sheet
  ExcelFile = xlrd.open_workbook(sys.argv[1]) 
  sheet = ExcelFile.sheet_by_index(0)

  # print content if the column 2 is not empty
  erows = []
  for i in range(sheet.nrows):
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

  print("All done! Good job!!!")

if __name__ == '__main__':
  read_excel()

