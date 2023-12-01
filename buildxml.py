import os
# import xml.etree.ElementTree as ETree
from lxml import etree as ETree

kmlns = {
        "":"http://www.opengis.net/kml/2.2",
        "gx":"http://www.google.com/kml/ext/2.2",
        # "kml":"http://www.opengis.net/kml/2.2",
        "atom":"http://www.w3.org/2005/Atom"
        }

def buildGPXTrk(filepath, rawdatetime, points, stats, dtformat):
    root = ETree.Element("gpx", version="1.0", creator="PGamma")
    ETree.SubElement(root, "time").text = rawdatetime.strftime(dtformat)
    bounds = ETree.SubElement(root, "bounds", minlat=str(stats["minlat"]), minlon=str(stats["minlon"]), maxlat=str(stats["maxlat"]), maxlon=str(stats["maxlon"]))
    trk = ETree.SubElement(root, "trk")
    trkseg = ETree.SubElement(trk, "trkseg")
    for point in points:
        trkpt = ETree.SubElement(trkseg, "trkpt", lat=str(point["lat"]), lon=str(point["lon"]))
        ETree.SubElement(trkpt, "ele").text = str(point["alt"])
        ETree.SubElement(trkpt, "time").text = point["datetm"].strftime(dtformat)
    tree = ETree.ElementTree(root)
    ETree.indent(tree)
    tree.write(filepath, encoding="utf-8", xml_declaration=True)

def buildGPXWp(filepath, rawdatetime, points, dtformat):
    root = ETree.Element("gpx", version="1.1", creator="PGamma")
    meta = ETree.SubElement(root, "metadata")
    ETree.SubElement(meta, "time").text = rawdatetime.strftime(dtformat)
    index = 0
    for point in points:
        wpt = ETree.SubElement(root, "wpt", lat=str(point["lat"]), lon=str(point["lon"]))
        ETree.SubElement(wpt, "ele").text = str(point["alt"])
        ETree.SubElement(wpt, "time").text = point["datetm"].strftime(dtformat)
        ETree.SubElement(wpt, "name").text = "%s" % index
        ETree.SubElement(wpt, "sym").text = "Flag, Blue"
        index += 1
    tree = ETree.ElementTree(root)
    ETree.indent(tree)
    tree.write(filepath, encoding="utf-8", xml_declaration=True)

# def buildKMLRoot():
#     return ETree.Element("kml")

def buildKMLRoot():
    root = ETree.Element("kml")
    root.set("xmlns","http://www.opengis.net/kml/2.2")
    root.set("xmlns:gx", "http://www.google.com/kml/ext/2.2")
    root.set("xmlns:kml", "http://www.opengis.net/kml/2.2")
    root.set("xmlns:atom", "http://www.w3.org/2005/Atom")
    return root

# def buildKMLAllRoot(outputTrk, outputWp):
#     root = buildKMLRoot()
#     allfolder = ETree.SubElement(root, "Folder")
#     ETree.SubElement(allfolder, "name").text = "all"
#     ETree.SubElement(allfolder, "open").text = "1"
#     if outputTrk:
#         trkfolder = ETree.SubElement(allfolder, "Folder")
#         ETree.SubElement(trkfolder, "name").text = "tracks"
#         ETree.SubElement(trkfolder, "open").text = "1"
#     else:
#         trkfolder = None
#     if outputWp:
#         wpfolder = ETree.SubElement(allfolder, "Folder")
#         ETree.SubElement(wpfolder, "name").text = "waypoints"
#         ETree.SubElement(wpfolder, "open").text = "1"
#     else:
#         wpfolder = None
#     return (root, trkfolder, wpfolder)

def buildKMLAllRoot():
    root = buildKMLRoot()
    allfolder = ETree.SubElement(root, "Folder")
    ETree.SubElement(allfolder, "name").text = "all"
    ETree.SubElement(allfolder, "open").text = "1"
    return (root)

def buildKMLTrk(filepath, points, stats, dtformat):
    filename = os.path.splitext(os.path.basename(filepath))[0]
    root = buildKMLRoot()

    doc = ETree.SubElement(root, "Document")
    ETree.SubElement(doc, "name").text = filename
    lookat = ETree.SubElement(doc, "LookAt")
    gxt = ETree.SubElement(lookat, "gx:TimeSpan")
    ETree.SubElement(gxt, "begin").text = stats["begintime"].strftime(dtformat)
    ETree.SubElement(gxt, "end").text = stats["endtime"].strftime(dtformat)
    ETree.SubElement(lookat, "longitude").text = "%0.6f" % ((stats["maxlon"] + stats["minlon"])/2)
    ETree.SubElement(lookat, "latitude").text = "%0.6f" % ((stats["maxlat"] + stats["minlat"])/2)
    ETree.SubElement(lookat, "altitude").text = "0"
    ETree.SubElement(lookat, "heading").text = "0"
    ETree.SubElement(lookat, "tilt").text = "0"
    ETree.SubElement(lookat, "range").text = "100"

    stylemap = ETree.SubElement(doc, "StyleMap", id="lineStyle")
    pair = ETree.SubElement(stylemap, "Pair")
    ETree.SubElement(pair, "key").text = "normal"
    ETree.SubElement(pair, "styleUrl").text = "#lineStyle_n"
    pair = ETree.SubElement(stylemap, "Pair")
    ETree.SubElement(pair, "key").text = "highlight"
    ETree.SubElement(pair, "styleUrl").text = "#lineStyle_h"
    style = ETree.SubElement(doc, "Style", id="lineStyle_h")
    lstyle = ETree.SubElement(style, "LineStyle")
    ETree.SubElement(lstyle, "color").text = "99ffac59"
    ETree.SubElement(lstyle, "width").text = "6"
    style = ETree.SubElement(doc, "Style", id="lineStyle_n")
    lstyle = ETree.SubElement(style, "LineStyle")
    ETree.SubElement(lstyle, "color").text = "99ffac59"
    ETree.SubElement(lstyle, "width").text = "6"

    stylemap = ETree.SubElement(doc, "StyleMap", id="track-none")
    pair = ETree.SubElement(stylemap, "Pair")
    ETree.SubElement(pair, "key").text = "normal"
    ETree.SubElement(pair, "styleUrl").text = "#track-none_n"
    pair = ETree.SubElement(stylemap, "Pair")
    ETree.SubElement(pair, "key").text = "highlight"
    ETree.SubElement(pair, "styleUrl").text = "#track-none_h"
    style = ETree.SubElement(doc, "Style", id="track-none_h")
    iconstyle = ETree.SubElement(style, "IconStyle")
    ETree.SubElement(iconstyle, "scale").text = "1.2"
    ETree.SubElement(iconstyle, "heading").text = "0"
    icon = ETree.SubElement(iconstyle, "Icon")
    ETree.SubElement(icon, "href").text = "https://earth.google.com/images/kml-icons/track-directional/track-none.png"
    style = ETree.SubElement(doc, "Style", id="track-none_n")
    iconstyle = ETree.SubElement(style, "IconStyle")
    ETree.SubElement(iconstyle, "scale").text = "0.5"
    ETree.SubElement(iconstyle, "heading").text = "0"
    icon = ETree.SubElement(iconstyle, "Icon")
    ETree.SubElement(icon, "href").text = "https://earth.google.com/images/kml-icons/track-directional/track-none.png"
    labelstyle = ETree.SubElement(style, "LabelStyle")
    ETree.SubElement(labelstyle, "scale").text = "0"

    tracksfolder = ETree.SubElement(doc, "Folder")
    ETree.SubElement(tracksfolder, "name").text = "Tracks"
    subfolder = ETree.SubElement(tracksfolder, "Folder")
    ETree.SubElement(subfolder, "name").text = "Items"
    timespan = ETree.SubElement(subfolder, "TimeSpan")

    pointsfolder = ETree.SubElement(subfolder, "Folder")
    ETree.SubElement(pointsfolder, "name").text = "Points"

    index = 0
    coordsall = ""
    for point in points:
        placemark = ETree.SubElement(pointsfolder, "Placemark")
        ETree.SubElement(placemark, "name").text = "%s" % index
        description = ETree.SubElement(placemark, "description")
        ETree.SubElement(description, "p").text = "Longitude: %0.6f" % point["lon"]
        ETree.SubElement(description, "p").text = "Latitude: %0.6f" % point["lat"]
        ETree.SubElement(description, "p").text = "Altitude: %0.1f meters" % point["alt"]
        ETree.SubElement(description, "p").text = "Time: %s" % point["datetm"].strftime(dtformat)
        lookatpm = ETree.SubElement(placemark, "LookAt")
        ETree.SubElement(lookatpm, "longitude").text = str(point["lon"])
        ETree.SubElement(lookatpm, "latitude").text = str(point["lat"])
        ETree.SubElement(lookatpm, "altitude").text = str(point["alt"])
        ETree.SubElement(lookatpm, "heading").text = "0"
        ETree.SubElement(lookatpm, "tilt").text = "0"
        ETree.SubElement(lookatpm, "range").text = "100"
        timestamp = ETree.SubElement(placemark, "TimeStamp")
        ETree.SubElement(timestamp, "when").text = point["datetm"].strftime(dtformat)
        ETree.SubElement(placemark, "styleUrl").text = "#track-none"
        pt = ETree.SubElement(placemark, "Point")
        coords = "%0.6f,%0.6f,%0.1f" % (point["lon"],point["lat"],point["alt"])
        ETree.SubElement(pt, "coordinates").text = coords
        coordsall += " %s" % coords
        index += 1

    placemark2 = ETree.SubElement(subfolder, "Placemark")
    ETree.SubElement(placemark2, "name").text = "Path"
    ETree.SubElement(placemark2, "styleUrl").text = "#lineStyle"
    linestring = ETree.SubElement(placemark2, "LineString")
    ETree.SubElement(linestring, "tessellate").text = "1"
    ETree.SubElement(linestring, "coordinates").text = coordsall

    # for (prefix, uri) in kmlns.items(): ETree.register_namespace(prefix, uri)
    tree = ETree.ElementTree(root)
    ETree.indent(tree)
    tree.write(filepath, encoding="utf-8", xml_declaration=True)

def buildKMLWp(filepath, points, stats, dtformat):
    filename = os.path.splitext(os.path.basename(filepath))[0]
    root = buildKMLRoot()

    doc = ETree.SubElement(root, "Document")
    ETree.SubElement(doc, "name").text = filename
    lookat = ETree.SubElement(doc, "LookAt")
    gxt = ETree.SubElement(lookat, "gx:TimeSpan")
    ETree.SubElement(gxt, "begin").text = stats["begintime"].strftime(dtformat)
    ETree.SubElement(gxt, "end").text = stats["endtime"].strftime(dtformat)
    ETree.SubElement(lookat, "longitude").text = "%0.6f" % ((stats["maxlon"] + stats["minlon"])/2)
    ETree.SubElement(lookat, "latitude").text = "%0.6f" % ((stats["maxlat"] + stats["minlat"])/2)
    ETree.SubElement(lookat, "altitude").text = "0"
    ETree.SubElement(lookat, "heading").text = "0"
    ETree.SubElement(lookat, "tilt").text = "0"
    ETree.SubElement(lookat, "range").text = "100"

    stylemap = ETree.SubElement(doc, "StyleMap", id="waypoint")
    pair = ETree.SubElement(stylemap, "Pair")
    ETree.SubElement(pair, "key").text = "normal"
    ETree.SubElement(pair, "styleUrl").text = "#waypoint_n"
    pair = ETree.SubElement(stylemap, "Pair")
    ETree.SubElement(pair, "key").text = "highlight"
    ETree.SubElement(pair, "styleUrl").text = "#waypoint_h"
    style = ETree.SubElement(doc, "Style", id="waypoint_h")
    iconstyle = ETree.SubElement(style, "IconStyle")
    ETree.SubElement(iconstyle, "scale").text = "1.2"
    icon = ETree.SubElement(iconstyle, "Icon")
    ETree.SubElement(icon, "href").text = "https://maps.google.com/mapfiles/kml/pal4/icon61.png"
    style = ETree.SubElement(doc, "Style", id="waypoint_n")
    iconstyle = ETree.SubElement(style, "IconStyle")
    ETree.SubElement(iconstyle, "scale").text = "1"
    icon = ETree.SubElement(iconstyle, "Icon")
    ETree.SubElement(icon, "href").text = "https://maps.google.com/mapfiles/kml/pal4/icon61.png"

    wpfolder = ETree.SubElement(doc, "Folder")
    ETree.SubElement(wpfolder, "name").text = "Waypoints"
    
    index = 0
    for point in points:
        placemark = ETree.SubElement(wpfolder, "Placemark")
        ETree.SubElement(placemark, "name").text = str(index)
        timestamp = ETree.SubElement(placemark, "TimeStamp")
        ETree.SubElement(timestamp, "when").text = point["datetm"].strftime(dtformat)
        ETree.SubElement(placemark, "styleUrl").text = "#waypoint"
        pt = ETree.SubElement(placemark, "Point")
        ETree.SubElement(pt, "coordinates").text = "%0.6f,%0.6f,%0.1f" % (point["lon"], point["lat"], point["alt"])
        index += 1

    # for (prefix, uri) in kmlns.items(): ETree.register_namespace(prefix, uri)
    tree = ETree.ElementTree(root)
    ETree.indent(tree)
    tree.write(filepath, encoding="utf-8", xml_declaration=True)

def combineKML(kmldir, overwrite=False):
    # for (prefix, uri) in kmlns.items():
    #     # print(prefix, uri)
    #     ETree.register_namespace(prefix, uri)
    # ETree.register_namespace("", "http://www.opengis.net/kml/2.2")
    # ns = kmlns
    allfile = "all"
    file = os.path.join(kmldir, "%s.kml" % allfile)
    if os.path.exists(file):
        tree = ETree.parse(file)
        root = tree.getroot()
        # print('exists')
    else:
        root = buildKMLAllRoot()
        print('building')
    docs = {}
    for folder in root.findall('./Folder/Folder', kmlns):
        # print(folder)
        foldername = folder.find('name', kmlns).text
        docs[foldername] = {}
        for doc in folder.findall('Document', kmlns):
            docname = doc.find('name', kmlns).text
            # print(docname)
            docs[foldername][docname] = doc
            print(docname, doc)
    # print(docs)
    print()
    kmlfiles = [os.path.splitext(kmlfile)[0] for kmlfile in os.listdir(kmldir) if os.path.splitext(kmlfile)[1].lower() == ".kml"]
    if allfile in kmlfiles: kmlfiles.remove(allfile)
    for kmlfile in kmlfiles:
        kmltype = kmlfile.rsplit("_", maxsplit=1)[1]
        if kmltype == "trk":
            # print(kmlfile, "is trk")
            if not "tracks" in docs:
                docs["tracks"] = {}
                print("creating tracks")
            else:
                print('tracks found')
            if not kmlfile in docs["tracks"] or overwrite:
                docs["tracks"][kmlfile] = ETree.parse(os.path.join(kmldir, "%s.kml" % kmlfile)).getroot().find('Document', kmlns)
                print("%s creating" % kmlfile)
            else:
                print("%s exists" % kmlfile)
            print()
        elif kmltype == "wp":
            # print(kmlfile, "is wp")
            if not "waypoints" in docs:
                docs["waypoints"] = {}
                print("creating waypoints")
            else:
                print('waypoints found')
            if not kmlfile in docs["waypoints"] or overwrite:
                docs["waypoints"][kmlfile] = ETree.parse(os.path.join(kmldir, "%s.kml" % kmlfile)).getroot().find('Document', kmlns)
                print("%s creating" % kmlfile)
            else:
                print("%s exists" % kmlfile)
            print()
            print()
        elif kmltype == "grid":
            # print(kmlfile, "is grid")
            print()
        else:
            pass
            # print(kmlfile, "is other")
    # build new tree
    print("building new root")
    root2 = buildKMLAllRoot()
    all2 = findKMLAllFolder(root2)
    print(all2)
    print(sorted(list(docs)))
    for folder in sorted(list(docs)):
        print(folder)
        folderElement = findKMLFolder(root2, folder)
        print(folderElement, "folder element")
        if not folderElement:
            print("%s folder added" % folder)
            folderElement = addKMLFolder(root2, folder)
        for docname in sorted(docs[folder]):
            folderElement.append(docs[folder][docname])
            print("%s doc added" % docname)

    tree2 = ETree.ElementTree(root2)
    ETree.indent(tree2)
    tree2.write(os.path.join(kmldir, "temp", "all.kml"), encoding="utf-8", xml_declaration=True)

def findKMLFolder(root, foldername):
    return root.find('./Folder/Folder[name="%s"]' % foldername)

def addKMLFolder(root, foldername):
    folder = findKMLFolder(root, foldername)
    if not folder:
        allfolder = findKMLAllFolder(root)
        folder = ETree.SubElement(allfolder, "Folder")
        ETree.SubElement(folder, "name").text = foldername
        ETree.SubElement(folder, "open").text = "1"
    return folder

def findKMLAllFolder(root):
    return root.find('./Folder[name="all"]')

# kmlpath = "C:\\Projects\\Testing\\KML"
# combineKML(kmlpath)
