import os
import csv
from datetime import datetime
from buildxml import buildGPXTrk, buildGPXWp, buildKMLTrk, buildKMLWp, combineKML
import subprocess

rawdateformat = "%d.%m.%y"
rawtimeformat = "%H:%M:%S"
gpsdateformat = "%Y-%m-%dT"
gpstimeformat = "%H:%M:%SZ"
gpsdatetimeformat = gpsdateformat + gpstimeformat

class invalidrow(Exception): pass

pdir = os.path.dirname(os.path.realpath(__file__))

now = datetime.now()

logfile = open(os.path.join(pdir, "process.log"), "a")
def writelog(line='', logfile=logfile): logfile.write(line + '\n')
logdiv = "*"*60
writelog(logdiv)
writelog("%s - Processing - %s" % (os.path.basename(__file__), now.strftime("%Y.%m.%d %H:%M:%S")))
writelog()

configdata = {}
configfile = os.path.join(pdir, "config.txt")
if os.path.exists(configfile):
	writelog("Reading config - %s" % configfile)
	writelog()
	with open(configfile, 'r', newline='') as cfgfile:
		creader = csv.reader(cfgfile, delimiter='=')
		for row in creader:
			if len(row) == 2 and row[0][0] != "#":
				configdata[row[0].strip()] = row[1].strip()
else:
	writelog("Config file not found - using defaults")
	writelog()
msg = "Working directory: ~"
if not "dir" in configdata:
	cwd = pdir
else:
	if configdata["dir"] == "":
		cwd = pdir
	else:
		cwd = configdata["dir"]
		if not os.path.exists(cwd):
			os.makedirs(cwd)
			msg += " was created"
writelog(msg.replace("~", cwd))
overwrite = configdata["overwrite"].lower() == 'true' if "overwrite" in configdata else False
writedeleted = configdata["writedeleted"].lower() == 'true' if "writedeleted" in configdata else False
outputGPXTrk = configdata["outputGPXTrk"].lower() == 'true' if "outputGPXTrk" in configdata else True
outputGPXWp = configdata["outputGPXWp"].lower() == 'true' if "outputGPXWp" in configdata else True
outputKMLTrk = configdata["outputKMLTrk"].lower() == 'true' if "outputKMLTrk" in configdata else True
outputKMLWp = configdata["outputKMLWp"].lower() == 'true' if "outputKMLWp" in configdata else True

writelog()

gammadir = os.path.join(cwd, "Gamma")
gpxdir = os.path.join(cwd, "GPX")
kmldir = os.path.join(cwd, "KML")
rapdir = os.path.join(cwd, "RAP")
surfdir = os.path.join(cwd, "Surfer")
for dir in [gammadir, gpxdir, kmldir, rapdir, surfdir]:
	if not os.path.exists(dir):
		os.makedirs(dir)

filenames = os.listdir(gammadir)

datafiles = []
for fullname in filenames:
	(fname, fext) = os.path.splitext(fullname)
	if fext.lower() == ".txt" and not any([fname[-4:] == fn for fn in ["Edit", "_del"]]):
		datafiles += [fullname]
datafiles.sort()

writelog("Found %s data files:" % len(datafiles))
for f in datafiles:
	writelog("    %s" % os.path.join(gammadir, f))
writelog()

for datafile in datafiles:
	writelog("Processing data file: %s" % os.path.join(gammadir, datafile))
	filename = os.path.splitext(datafile)[0]
	wpath = os.path.join(gammadir, '%s_Edit.txt' % filename)
	if os.path.exists(wpath) and not overwrite:
		writelog("Found edit file: %s" % wpath)
		writelog()
		continue
	else:
		writelog("Generating edit file: %s" % wpath)

	rawrowcount = 0
	rowcount = 0
	rawdatetime = now
	validrowcount = 0
	invalidrowcount = 0
	prevvalidrow = 0
	prevwp = 0
	prevlon = prevlat = 0.0
	stats = {}
	addpoints = []

	dpath = os.path.join(gammadir, datafile)
	dfile = open(dpath, "r")
	freader = csv.reader(dfile, delimiter=';')

	wfile = open(wpath, 'w', newline="")
	fwriter = csv.writer(wfile, delimiter=';')
	fwriter.writerow(["Waypoint","Elevation","Latitude","Longitude","Date","Time"])

	if writedeleted:
		ipath = os.path.join(gammadir, '%s_del.txt' % filename)
		ifile = open(ipath, 'w')
		writelog("Generating delete file: %s" % ipath)

	for row in freader:
		rawrowcount += 1
		rowlen = len(row)
		if rowlen == 0:
			pass # ignore empty rows
		elif rowlen == 1:
			writelog('%s' % row[0])
			if rawrowcount == 4:
				rawdatestr = row[0][1:]
			if rawrowcount == 5:
				rawtimestr = row[0][1:]
				try:
					rawdatetime = datetime.strptime(rawdatestr + rawtimestr, rawdateformat + rawtimeformat)
				except:
					rawdatetime = now
		else:
			try:
				rowcount += 1
				if rowcount == 1:
					RowLength = len(row)
				try:
					wp = int(row[0])
				except:
					reason = 'Invalid waypoint'
					raise invalidrow()
					
				if rowlen < RowLength:
					reason = 'Missing values'
					raise invalidrow()
				
				if rowlen > RowLength:
					reason = 'Additional values'
					raise invalidrow()

				try:
					alt = float(row[1])
				except:
					reason = 'Invalid elevation'
					raise invalidrow()

				try:
					lat = float(row[2])
				except:
					reason = 'Invalid latitue'
					raise invalidrow()
				if rowcount > 1 and (abs((lat - reflat)*2/(lat + reflat)) > 0.1):
					reason = 'Latitude out of range'
					raise invalidrow()
				
				try:
					lon = float(row[3])
				except:
					reason = 'Invalid longitude'
					raise invalidrow()
				if rowcount > 1 and (abs((lon - reflon)*2/(lon + reflon)) > 0.01):
					reason = 'Longitude out of range'
					raise invalidrow()
				
				if RowLength == 6:
					try:
						datetm = datetime.strptime(row[4] + row[5], rawdateformat + rawtimeformat)
					except:
						reason = 'Invalid datetime'
						raise invalidrow()
				else:
					datetm = rawdatetime
					row += rawdatestr
					row += rawtimestr

				validrowcount += 1
				point = {"wp":wp, "alt":alt, "lat":lat, "lon":lon, "datetm":datetm}
				if lon != prevlon or lat != prevlat:
					addpoints += [point.copy()]

				prevlon = lon
				prevlat = lat

				row = ["%5i" % validrowcount, "%6.1f" % alt, "%10.6f" % lat, "%10.6f" % lon, "%8s" % datetm.strftime(rawdateformat), "%8s" % datetm.strftime(rawtimeformat)]
				fwriter.writerow(row)

				if validrowcount == 1:
					reflat = stats["minlat"] = stats["maxlat"] = lat
					reflon = stats["minlon"] = stats["maxlon"] = lon
					stats["begintime"] = datetm
				else:
					if lat < stats["minlat"] : stats["minlat"] = lat
					if lat > stats["maxlat"] : stats["maxlat"] = lat
					if lon < stats["minlon"] : stats["minlon"] = lon
					if lon > stats["maxlon"] : stats["maxlon"] = lon
					stats["endtime"] = datetm

				validgap = rowcount - prevvalidrow -1
				if validgap > 0:
					writelog('%s points deleted' % validgap)
				else:
					wpgap = wp - prevwp - 1
					if wpgap > 0:
						writelog('Waypoint gap - %s points (%s to %s)' % (wpgap, prevwp, wp))

				prevwp = wp
				prevvalidrow = rowcount

			except invalidrow:
				invalidrowcount += 1
				if writedeleted:
					ifile.write('%-55s - %s\n' % (';'.join([item for item in row]), reason))
			
	# write stats to logfile
	validrowspercent = validrowcount / rowcount * 100
	invalidrowspercent = invalidrowcount / rowcount * 100
	writelog('Data points: %s' % rowcount)
	writelog('Valid data points: %s (%.1f%%)' % (validrowcount, validrowspercent))
	writelog('Invalid data points: %s (%.1f%%)' % (invalidrowcount, invalidrowspercent))
	writelog('Latitude (min, max, diff): %.6f, %.6f, %.6f' %(stats["minlat"], stats["maxlat"], abs(stats["maxlat"] - stats["minlat"])))
	writelog('Longitude (min, max, diff): %.6f, %.6f, %.6f' %(stats["minlon"], stats["maxlon"], abs(stats["maxlon"] - stats["minlon"])))
	writelog()

	dfile.close()
	wfile.close()
	if writedeleted: ifile.close()

	if outputGPXTrk:
		gpxfile = os.path.join(gpxdir, '%s_Edit_trk.gpx' % filename)
		if os.path.exists(gpxfile) and not overwrite:
			writelog("Found GPX file: %s" % gpxfile)
		else:
			writelog("Generating GPX file: %s" % gpxfile)
			buildGPXTrk(gpxfile, rawdatetime, addpoints, stats, gpsdatetimeformat)

	if outputGPXWp:
		gpxfile = os.path.join(gpxdir, '%s_Edit_wp.gpx' % filename)
		if os.path.exists(gpxfile) and not overwrite:
			writelog("Found GPX file: %s" % gpxfile)
		else:
			writelog("Generating GPX file: %s" % gpxfile)
			buildGPXWp(gpxfile, rawdatetime, addpoints, gpsdatetimeformat)
	
	if outputKMLTrk:
		kmlfile = os.path.join(kmldir, '%s_Edit_trk.kml' % filename)
		if os.path.exists(kmlfile) and not overwrite:
			writelog("Found KML file: %s" % kmlfile)
		else:
			writelog("Generating KML file: %s" % kmlfile)
			buildKMLTrk(kmlfile, addpoints, stats, gpsdatetimeformat)

	if outputKMLWp:
		kmlfile = os.path.join(kmldir, '%s_Edit_wp.kml' % filename)
		if os.path.exists(kmlfile) and not overwrite:
			writelog("Found KML file: %s" % kmlfile)
		else:
			writelog("Generating KML file: %s" % kmlfile)
			buildKMLWp(kmlfile, addpoints, stats, gpsdatetimeformat)
	
	writelog()

# combine KML files
combineKML(kmldir, overwrite=overwrite, writelog=writelog)

writelog()
writelog("Processing completed - %s" % datetime.now().strftime("%Y.%m.%d %H:%M:%S"))
writelog(logdiv)
logfile.close()

if "scripter" in configdata:
	subprocess.run([configdata["scripter"], "-x", os.path.join(pdir,"process.bas"), cwd])
