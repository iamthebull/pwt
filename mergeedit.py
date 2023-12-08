import os
import sys
import csv

(cwd, scriptname) = os.path.split(sys.argv[0])

filenames = os.listdir(cwd)

datafiles = []
for fullname in filenames:
    (fname, fext) = os.path.splitext(fullname)
    if fname[-4:] + fext.lower() == "Edit.txt":
        datafiles += [fullname]

numdatafiles = len(datafiles)
datafiles.sort()

fileindex = 1
wpindex = 1
for datafile in datafiles:
    if fileindex == 1:
        wpath = os.path.join(cwd, 'merged.txt')
        wfile = open(wpath, 'w', newline="")
        fwriter = csv.writer(wfile, delimiter=';')
        fwriter.writerow(["Waypoint","Elevation","Latitude","Longitude","Date","Time"])
    fileindex += 1

    dpath = os.path.join(cwd, datafile)
    dfile = open(dpath, "r")
    freader = csv.reader(dfile, delimiter=';')

    rowindex = 1
    for row in freader:
        if rowindex > 1:
            row[0] = "%5i" % wpindex
            fwriter.writerow(row)
            wpindex += 1
        rowindex += 1
