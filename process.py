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

class FileWriter:
	def __init__(self, path, filename, mode='w'):
		self.path = path
		self.filename = filename
		self.mode = mode
		self.file = None

	def open(self):
		if not self.file:
			self.file = open(os.path.join(self.path, self.filename), self.mode)
	
	def close(self):
		if self.file:
			self.file.close()
			self.file = None

	def write(self, line=""):
		self.file.write(line)

	def writeln(self, line=""):
		self.file.write(line + '\n')

	def setmode(self, mode):
		self.mode = mode

pdir = os.path.dirname(os.path.realpath(__file__))

now = datetime.now()

lfw = FileWriter(pdir, "process.log")
lfw.open()

logdiv = "*"*60
lfw.writeln(logdiv)
lfw.writeln("%s - Processing - %s" % (os.path.basename(__file__), now.strftime("%Y.%m.%d %H:%M:%S")))
lfw.writeln()

configdata = {}
configfile = os.path.join(pdir, "config.txt")
if os.path.exists(configfile):
	lfw.writeln("Reading config - %s" % configfile)
	lfw.writeln()
	with open(configfile, 'r', newline='') as cfgfile:
		creader = csv.reader(cfgfile, delimiter='=')
		for row in creader:
			if len(row) == 2 and row[0][0] != "#":
				configdata[row[0].strip()] = row[1].strip()
else:
	lfw.writeln("Config file not found - using defaults")
	lfw.writeln()
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
lfw.writeln(msg.replace("~", cwd))

overwrite = configdata["overwrite"].lower() == 'true' if "overwrite" in configdata else False
writedeleted = configdata["writedeleted"].lower() == 'true' if "writedeleted" in configdata else False

laterror = float(configdata["laterror"]) if "laterror" in configdata else 0.1
lonerror = float(configdata["lonerror"]) if "lonerror" in configdata else 0.1

outputGPXTrk = configdata["outputGPXTrk"].lower() == 'true' if "outputGPXTrk" in configdata else True
outputGPXWp = configdata["outputGPXWp"].lower() == 'true' if "outputGPXWp" in configdata else True
outputKMLTrk = configdata["outputKMLTrk"].lower() == 'true' if "outputKMLTrk" in configdata else True
outputKMLWp = configdata["outputKMLWp"].lower() == 'true' if "outputKMLWp" in configdata else True

lfw.writeln()

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

lfw.writeln("Found %s data files:" % len(datafiles))
for f in datafiles:
	lfw.writeln("    %s" % os.path.join(gammadir, f))
lfw.writeln()

for datafile in datafiles:
	lfw.writeln("Processing data file: %s" % os.path.join(gammadir, datafile))
	filename = os.path.splitext(datafile)[0]
	wpath = os.path.join(gammadir, '%s_Edit.txt' % filename)
	if os.path.exists(wpath) and not overwrite:
		lfw.writeln("Found edit file: %s" % wpath)
		lfw.writeln()
		continue
	else:
		lfw.writeln("Generating edit file: %s" % wpath)

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
		lfw.writeln("Generating delete file: %s" % ipath)
	lfw.writeln()

	for row in freader:
		rawrowcount += 1
		rowlen = len(row)
		if rowlen == 0:
			pass # ignore empty rows
		elif rowlen == 1:
			lfw.writeln('%s' % row[0])
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
				if rowcount > 1 and abs(lat - prevlat) > laterror:
					reason = 'Latitude jump detected'
					raise invalidrow()

				try:
					lon = float(row[3])
				except:
					reason = 'Invalid longitude'
					raise invalidrow()
				if rowcount > 1 and abs(lon - prevlon) > lonerror:
						reason = 'Longitude jump detected'
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

				row = ["%5i" % wp, "%6.1f" % alt, "%10.6f" % lat, "%10.6f" % lon, "%8s" % datetm.strftime(rawdateformat), "%8s" % datetm.strftime(rawtimeformat)]
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
					lfw.writeln('%s points deleted' % validgap)
				else:
					wpgap = wp - prevwp - 1
					if wpgap > 0:
						lfw.writeln('Waypoint gap - %s points (%s to %s)' % (wpgap, prevwp, wp))

				prevwp = wp
				prevvalidrow = rowcount

			except invalidrow:
				invalidrowcount += 1
				if writedeleted:
					ifile.write('%-55s - %s\n' % (';'.join([item for item in row]), reason))
			
	# write stats to logfile
	validrowspercent = validrowcount / rowcount * 100
	invalidrowspercent = invalidrowcount / rowcount * 100
	lfw.writeln('Data points: %s' % rowcount)
	lfw.writeln('Valid data points: %s (%.1f%%)' % (validrowcount, validrowspercent))
	lfw.writeln('Invalid data points: %s (%.1f%%)' % (invalidrowcount, invalidrowspercent))
	lfw.writeln('Latitude (min, max, diff): %.6f, %.6f, %.6f' %(stats["minlat"], stats["maxlat"], abs(stats["maxlat"] - stats["minlat"])))
	lfw.writeln('Longitude (min, max, diff): %.6f, %.6f, %.6f' %(stats["minlon"], stats["maxlon"], abs(stats["maxlon"] - stats["minlon"])))
	lfw.writeln()

	dfile.close()
	wfile.close()
	if writedeleted: ifile.close()

	if outputGPXTrk:
		gpxfile = os.path.join(gpxdir, '%s_Edit_trk.gpx' % filename)
		if os.path.exists(gpxfile) and not overwrite:
			lfw.writeln("Found GPX file: %s" % gpxfile)
		else:
			lfw.writeln("Generating GPX file: %s" % gpxfile)
			buildGPXTrk(gpxfile, rawdatetime, addpoints, stats, gpsdatetimeformat)

	if outputGPXWp:
		gpxfile = os.path.join(gpxdir, '%s_Edit_wp.gpx' % filename)
		if os.path.exists(gpxfile) and not overwrite:
			lfw.writeln("Found GPX file: %s" % gpxfile)
		else:
			lfw.writeln("Generating GPX file: %s" % gpxfile)
			buildGPXWp(gpxfile, rawdatetime, addpoints, gpsdatetimeformat)
	
	if outputKMLTrk:
		kmlfile = os.path.join(kmldir, '%s_Edit_trk.kml' % filename)
		if os.path.exists(kmlfile) and not overwrite:
			lfw.writeln("Found KML file: %s" % kmlfile)
		else:
			lfw.writeln("Generating KML file: %s" % kmlfile)
			buildKMLTrk(kmlfile, addpoints, stats, gpsdatetimeformat)

	if outputKMLWp:
		kmlfile = os.path.join(kmldir, '%s_Edit_wp.kml' % filename)
		if os.path.exists(kmlfile) and not overwrite:
			lfw.writeln("Found KML file: %s" % kmlfile)
		else:
			lfw.writeln("Generating KML file: %s" % kmlfile)
			buildKMLWp(kmlfile, addpoints, stats, gpsdatetimeformat)
	
	lfw.writeln()

lfw.close()

if "surferdir" in configdata:
	subprocess.run([os.path.join(configdata["surferdir"],"scripter.exe"), "-x", os.path.join(pdir,"process.bas"), cwd])

lfw.setmode('a')
lfw.open()

# combine KML files
combineKML(kmldir, overwrite=overwrite, writelog=lfw.writeln)

lfw.writeln()
lfw.writeln("Processing completed - %s" % datetime.now().strftime("%Y.%m.%d %H:%M:%S"))
lfw.writeln(logdiv)
lfw.close()

# if "surferdir" in configdata:
# 	subprocess.run([os.path.join(configdata["surferdir"],"scripter.exe"), "-x", os.path.join(pdir,"process.bas"), cwd])
