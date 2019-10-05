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

def show_subject(qstr, astr, cidx, ttl):
  t_w = os.get_terminal_size().columns
  # '%' means formatted string or Arithmetical complement
  if ((len("(%d/%d)" % (cidx, ttl)) + len(qstr) + 3) % t_w + len(astr)) > t_w:
    print("(%d/%d)" % (cidx, ttl) + qstr + ' :')
  else:
    print("(%d/%d)" % (cidx, ttl) + qstr + ' : ', end='')

def usage_help():
  print("Usage  :", sys.argv[0], "excelfile [sheet_index=1] [random=0] [start_index=1] [word_amount]")
  print("Example:", sys.argv[0], "~/kdc.xls 2 1 20 15")
  exit(1)

def read_excel():
  # check the arguments
  idx, rflag, arglen, rstart, ramount = 0, False, len(sys.argv), 0, 0
  if arglen < 2:
    usage_help()
  
  if arglen > 5:
    ramount = int(sys.argv[5])
  
  if arglen > 4:
    rstart = int(sys.argv[4])
  
  if arglen > 3:
    rflag = bool(int(sys.argv[3]))
  
  if arglen > 2:
    idx = int(sys.argv[2]) - 1

  # open excel file and read the specified sheet
  ExcelFile = xlrd.open_workbook(sys.argv[1]) 
  sheet = ExcelFile.sheet_by_index(idx)

  # filter the range, ignore the 1st title row
  if rstart > 0 and ramount > 0 and (rstart + ramount) < sheet.nrows:
    nrange = range(rstart, rstart + ramount)
  elif rstart > 0 and rstart < sheet.nrows:
    ramount = sheet.nrows - rstart
    nrange = range(rstart, sheet.nrows)
  else:
    rstart = 1
    ramount = sheet.nrows - 1
    nrange = range(sheet.nrows)
  
  # use random shuttle or not
  if rflag:
    nrange = list(nrange)
    random.shuffle(nrange)

  # print sheet's name, rows, order and range
  clear_terminal(0)
  print("Sheet.Name:", sheet.name, " Sheet.Rows:", sheet.nrows, " Random:", rflag, " Range:", "%d-%d" % (rstart, rstart + ramount - 1))

  # print content if the column 2 is not empty
  seq, erows = 1, []
  for i in nrange:
    # ignore the first row
    if i == 0:
      continue
    # wrap line according to len of string
    if sheet.row_values(i)[1] != "":
      show_subject(sheet.row_values(i)[3], sheet.row_values(i)[1], seq, ramount)
      kdc = input()
      if sheet.row_values(i)[1] == kdc:
        print("Good!", sheet.row_values(i)[2])
      else:
        # if input is wrong word, put it into the loop cycle array
        erows.append(sheet.row_values(i))
        print("No!!!", erows[-1][1], erows[-1][2])
      # increasing sequnce
      seq += 1

  # repeat until everything is ok
  i, seq, ramount = 0, 1, len(erows)
  clear_terminal(0)
  while len(erows) > 0:
    show_subject(erows[i][3], erows[i][1], seq, ramount)
    kdc = input()
    if erows[i][1] == kdc:
      print("Good!", erows[i][2])
      erows.pop(i)
    else:
      print("No!!!", erows[i][1], erows[i][2])
      i += 1
    # increasing sequnce
    seq += 1
  
    if i >= len(erows):
      i, seq, ramount = 0, 1, len(erows)
      clear_terminal(0.3)
      

  print("All done! Good job!!!")

if __name__ == '__main__':
  read_excel()

