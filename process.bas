' script for creating surfer plots
Sub Main

	' get the current dir
	cwd = CurDir
	configArray = readConfigFile(cwd)

	' configuration settings
	pdir = getConfigVal("dir", configArray, "str")
	titletext = getConfigVal("titletext", configArray, "str")
	operator = getConfigVal("operator", configArray, "str")
	overwrite = getConfigVal("overwrite", configArray, "bool")
	background = getConfigVal("background", configArray, "bool")
	leaveopen = getConfigVal("leaveopen", configArray, "bool")
	pagewidth = getConfigVal("pagewidth", configArray, "val")
	pageheight = getConfigVal("pageheight", configArray, "val")
	numcolsperpoint = getConfigVal("numcolsperpoint", configArray, "val")
	numrows = getConfigVal("numrows", configArray, "val")

	' map size
	mapheight = getConfigVal("mapheight", configArray, "val")
	mapbottommargin = getConfigVal("mapbottommargin", configArray, "val")
	mapsidemargin = getConfigVal("mapsidemargin", configArray, "val")
	widescale = getConfigVal("widescale", configArray, "val")
	' scale size
	scaletopmargin = getConfigVal("scaletopmargin", configArray, "val")
	scalesidemargin = getConfigVal("scalesidemargin", configArray, "val")
	scaleoffset = getConfigVal("scaleoffset", configArray, "val")
	' space size
	spacecellwidth = getConfigVal("spacecellwidth", configArray, "val")

	'axis labels
	showtopaxislabels = getConfigVal("showtopaxislabels", configArray, "bool")
	showbottomaxislabels = getConfigVal("showbottomaxislabels", configArray, "bool")
	showleftaxislabels = getConfigVal("showleftaxislabels", configArray, "bool")
	showrightaxislabels = getConfigVal("showrightaxislabels", configArray, "bool")
	' axis titles
	showtopaxistitle = getConfigVal("showtopaxistitle", configArray, "bool")
	showbottomaxistitle = getConfigVal("showbottomaxistitle", configArray, "bool")
	showleftaxistitle = getConfigVal("showleftaxistitle", configArray, "bool")
	showrightaxistitle = getConfigVal("showrightaxistitle", configArray, "bool")
	axistitleoffset = getConfigVal("axistitleoffset", configArray, "val")

	If showtopaxistitle Then
		topaxistitle = "[meters]"
	Else
		topaxistitle = ""
	End If
	If showbottomaxistitle Then
		bottomaxistitle = "[meters]"
	Else
		bottomaxistitle = ""
	End If
	If showleftaxistitle Then
		leftaxistitle = "Depth [meters]"
	Else
		leftaxistitle = ""
	End If
	If showrightaxistitle Then
		rightaxistitle = "Depth [feet]"
	Else
		rightaxistitle = ""
	End If

	' set directories
	If Command$ <> "" Then
		rootdir = Command$
	ElseIf pdir <> "" Then
		rootdir = pdir
	Else
		rootdir = cwd
	End If
	datafiledir = rootdir + "\Gamma\"
	surfdir = rootdir + "\Surfer\"
	kmldir = rootdir + "\KML\"
	rapdir = rootdir + "\RAP\"

	Debug.Clear

	logf = FreeFile
	logdiv = String(60, "*")
	Open cwd + "\process.log" For Append As #logf
	Print #logf, logdiv
	Print #logf, "process.bas - Processing - " + getNowStr()
	Print #logf, ""

	Print #logf, "Working directory: " + rootdir
	Print #logf, ""

	Debug.Print "Processing gamma"
	Debug.Print ""

    Print #logf, "Processing gamma"
    Print #logf, ""

	Set surf = CreateObject("surfer.application")
	surf.Visible = Not background

	datafiles = GetFiles(datafiledir, "*_Edit.txt")
	For Each datafile In datafiles
		Debug.Print "Found datafile: " + datafiledir + datafile
		Print #logf, "Found datafile: " + datafiledir + datafile
		filename = Split(datafile, ".")(0)
		gridfile = datafiledir + filename + ".grd"
		If Dir(gridfile) = "" or overwrite Then
			Debug.Print "Creating grid file: " + gridfile
			Print #logf, "Creating grid file: " + gridfile
			' create grid file
			surf.GridData6(DataFile:= datafiledir + datafile,  xCol:=4, yCol:=3, zCol:=2, Algorithm:=srfKriging, ShowReport:=False, OutGrid:=gridfile)
		Else
			Debug.Print "Found gridfile: " + gridfile
			Print #logf, "Found gridfile: " + gridfile
		End If

		plotfile = surfdir + filename + ".srf"
		If Dir(plotfile) = "" or overwrite Then
			Debug.Print "Creating plotfile: " + plotfile
			Print #logf, "Creating plotfile: " + plotfile
			' create plot object
			Set Plot = surf.Documents.Add(srfDocPlot)
			Set PageSetup = Plot.PageSetup
			PageSetup.Width = pagewidth
			PageSetup.Height = pageheight
	
			' create shapes object
			Set Shapes = Plot.Shapes
	
			' add contour plot
			Set MapFrame = Shapes.AddContourMap(GridFileName:=gridfile)
			' configure contour plot
			Set ContourLayer = MapFrame.Overlays(1)
			With ContourLayer
				.ShowMajorLabels = False
				.ShowMinorLabels = False
				.FillForegroundColorMap.LoadFile(surf.Path + "\Samples\Rainbow.clr")
				.FillContours = True
				.ShowColorScale = True
				With .ColorScale
					.Visible = True
					.Title = "Impulses\n Per Second\n (IPS)"
					.TitlePosition = srfColorScaleTitlePositionTop
					.TitleAngle = 0
					.TitleFont.hAlign = srfTACenter
				End With
			End With

			kmlfile = kmldir + filename + "_grid.kml"
			If Dir(kmlfile) = "" or overwrite Then
				Debug.Print "Creating kmlfile: " + kmlfile
				Print #logf, "Creating kmlfile: " + kmlfile
				' hide axis for kml export
				For Each axis In  MapFrame.Axes
					axis.Visible = False
				Next
				ContourLayer.ColorScale.Visible = False
				' export KML file
				Plot.Export2(FileName:= kmldir + filename + "_grid.kml", SelectionOnly:=False)
				' unhide axis and color scale
				For Each axis In MapFrame.Axes
					axis.Visible = True
				Next
				ContourLayer.ColorScale.Visible = True
			Else
				Debug.Print "Found kmlfile: " + kmlfile
				Print #logf, "Found kmlfile: " + kmlfile
			End If

			' add post map
			Set PostFrame = Shapes.AddPostMap(DataFileName:=datafiledir + datafile, xCol:=4, yCol:=3)
			Set PostLayer = PostFrame.Overlays(1)
			PostLayer.Symbol.Index = 12
			PostLayer.Symbol.Size = 0.05

			MapFrame.Selected = True
			PostFrame.Selected = True
			Set NewFrame = plot.Selection.OverlayMaps

			With NewFrame.Axes
				With .Item("Bottom Axis")
					.Title = "Longitude"
					.TitleFont.Size = 12
				End With
				With .Item("Left Axis")
					.Title = "Latitude"
					.TitleFont.Size = 12
				End With
			End With

			' add title text box
			title = "\fs200 " + titletext + "\n "
			Set TitleBox = Shapes.AddText(x:=pagewidth/2, y:=pageheight-0.25, Text:=title)
			TitleBox.Name = "Title"
			TitleBox.Font.hAlign = srfTACenter

			boxtext = "Gamma Survey\n " + operator + "\n " + getNowStr("dd.mm.yyyy")
			Set StampBox = Shapes.AddText(x:=4.5, y:=pageheight-1, Text:=boxtext)
			StampBox.Name = "Stamp"
			StampBox.Font.Size = 14
			StampBox.Font.hAlign = srfTACenter

			' save the plot
			Plot.SaveAs(FileName:=plotfile)
		Else
			Debug.Print "Found plotfile: " + plotfile
			Print #logf, "Found plotfile: " + plotfile
		End If

		datafile = Dir()
		Debug.Print ""
		Print #logf, ""

	Next

	Debug.Print "Gamma processing completed"
    Debug.Print ""
	Print #logf, "Gamma processing completed"
    Print #logf, ""

	Debug.Print "Processing RAP"
	Debug.Print ""
    Print #logf, "Processing RAP"
    Print #logf, ""

	folders = GetDirs(rapdir)
	For Each folder In folders
		Debug.Print "Found folder: " + folder
		Print #logf, "Found folder: " + folder
		datadir = rapdir + folder + "\Data_Cor\"
		datafiles = GetFiles(datadir, "*frn.txt")

		For Each datafile In datafiles
			Debug.Print "Found datafile: " + datafile
			Print #logf, "Found datafile: " + datafile

			filename = Split(datafile, ".")(0)
			gridfile = datadir + filename + ".grd"
			If (Dir(gridfile) = "") or overwrite Then
				creategridfile = True
			Else
				creategridfile = False
				Debug.Print "Found gridfile: " + gridfile
				Print #logf, "Found gridfile: " + gridfile
			End If

			plotfile = surfdir + folder + " " + Split(filename, "frn")(0) + ".srf"
			If (Dir(plotfile) = "") or overwrite Then
				createplotfile = True
			Else
				createplotfile = False
				Debug.Print "Found plotfile: " + plotfile
				Print #logf, "Found plotfile: " + plotfile
			End If

			If creategridfile or createplotfile Then
				' open datafile
				Set wks = surf.Documents.Open(datadir + datafile)
				rowCount = wks.Columns(1).RowCount
				samples = wks.cells(Row:=1,Col:=1).Value
				xMin = wks.cells(2,1).Value
				xMax = wks.cells(rowCount,Col:=1).Value
				xInt = wks.cells(samples+3,1).Value
				yMin = wks.cells(2,2).Value
				yMax = wks.cells(samples+1,2).Value
				yInt = wks.cells(3,2).Value - yMin

				' close datafile
				wks.Close
			End If

			If creategridfile Then
				Debug.Print "Creating gridfile: " + gridfile
				Print #logf, "Creating gridfile: " + gridfile
				NumCols = numcolsperpoint * (xMax - xMin) / xInt + 1
				surf.GridData6(DataFile:= datadir + datafile,  xCol:=1, yCol:=2, zCol:=3, NumCols:=NumCols, NumRows:=numrows, xMin:=xMin, xMax:=xMax, yMin:=yMin, yMax:=yMax, Algorithm:=srfKriging, ShowReport:=False, OutGrid:=gridfile)
			End If

			If createplotfile Then
				Debug.Print "Creating plotfile: " + plotfile
				Print #logf, "Creating plotfile: " + plotfile
				' create plot object
				Set Plot = surf.Documents.Add(srfDocPlot)
				Set PageSetup = Plot.PageSetup
				PageSetup.Width = pagewidth
				PageSetup.Height = pageheight
				scalingratio = (yMax - yMin) / mapheight

				' create shapes object
				Set Shapes = Plot.Shapes

				' add contour map
				Set MapFrame1 = Shapes.AddContourMap(GridFileName:=gridfile)
				With MapFrame1
					.Name = "Contour Map"
					With .Axes("Top Axis")
						.SetScale(MajorInterval:=xInt, LastMajorTick:=xMax)
						.ShowLabels = showtopaxislabels
						.Title = topaxistitle
					End With
					With .Axes("Bottom Axis")
						.SetScale(MajorInterval:=xInt, LastMajorTick:=xMax)
						.ShowLabels = showbottomaxislabels
						.Title = bottomaxistitle
						.TitleFont.hAlign = srfTACenter
						.TitleOffset2 = axistitleoffset
					End With
					With .Axes("Left Axis")
						.Reverse = True
						.SetScale(LastMajorTick:=yMax)
						.ShowLabels = showleftaxislabels
						.Title = leftaxistitle
					End With
					With .Axes("Right Axis")
						.SetScale(LastMajorTick:=yMax)
						.ShowLabels = showrightaxislabels
						.Title = rightaxistitle
					End With
					.xMapPerPU = scalingratio
					.yMapPerPU = scalingratio
				End With
				Set ContourLayer = MapFrame1.Overlays(1)
				With ContourLayer
					.ShowMajorLabels = False
					.ShowMinorLabels = False
					.FillForegroundColorMap.LoadPreset("rap")
					.FillForegroundColorMap.SetDataLimits(0, 1500)
					.FillContours = True
					.SetSimpleLevels(Min:=0, Max:=1500, Interval:=50)
					.ShowColorScale = True
				End With
				Set ColorScale1 = ContourLayer.ColorScale
				With ColorScale1
					.Name = "Contour Map Color Scale"
					.LabelFrequency = 5
					.Title = "\fs120 Contour Map\n ---------------------------\n \fs100 Relative Mechanical\n Strength (Yellow)\n Or\n Weakness (Violet)\n of Strata"
					.TitlePosition = srfColorScaleTitlePositionTop
					.TitleAngle = 0
					.TitleFont.hAlign = srfTACenter
					.TitleFont.Size = 12
				End With

				mf1height = MapFrame1.Height
				mf1width = MapFrame1.Width
				cs1height = ColorScale1.Height
				cs1width = ColorScale1.Width

				' add color releif map
				Set MapFrame2 = Plot.Shapes.AddColorReliefMap(GridFileName:=gridfile)
				With MapFrame2
					.Name = "Color Relief Map"
					With .Axes("Top Axis")
						.SetScale(MajorInterval:=xInt, LastMajorTick:=xMax)
						.ShowLabels = showtopaxislabels
						.Title = topaxistitle
					End With
					With .Axes("Bottom Axis")
						.SetScale(MajorInterval:=xInt, LastMajorTick:=xMax)
						.ShowLabels = showbottomaxislabels
						.Title = bottomaxistitle
						.TitleFont.hAlign = srfTACenter
						.TitleOffset2 = axistitleoffset
					End With
					With .Axes("Left Axis")
						.Reverse = True
						.SetScale(LastMajorTick:=yMax)
						.ShowLabels = showleftaxislabels
						.Title = leftaxistitle
					End With
					With .Axes("Right Axis")
						.SetScale(LastMajorTick:=yMax)
						.ShowLabels = showrightaxislabels
						.Title = rightaxistitle
					End With
					.xMapPerPU = scalingratio
					.yMapPerPU = scalingratio
				End With
				Set ColorReliefLayer = MapFrame2.Overlays(1)
				With ColorReliefLayer
					.TerrainRepresentation = srfTerrainRepresentationColorOnly
					.ColorMap.LoadPreset("Terrain")
					.ColorMap.SetDataLimits(0, 1500)
					.ShowColorScale = True
				End With
				Set ColorScale2 = ColorReliefLayer.ColorScale
				With ColorScale2
					.Name = "Color Relief Map Color Scale"
					.Title = "\fs120 Color Relief Map\n ---------------------------\n \fs100 Relative Mechanical\n Strength (Green)\n Or\n Weakness (Blue)\n of Strata"
					.TitlePosition = srfColorScaleTitlePositionTop
					.TitleAngle = 0
					.TitleFont.hAlign = srfTACenter
					.TitleFont.Size = 12
				End With

				mf2height = MapFrame2.Height
				mf2width = MapFrame2.Width
				cs2height = ColorScale2.Height
				cs2width = ColorScale2.Width

				' make a wide copy of color relief map
				Plot.Selection.DeselectAll
				MapFrame2.Selected = True
				Plot.Selection.Copy
				Plot.Shapes.Paste
				For Each shape In Plot.Selection
					If shape.Name = "Color Relief Map" Then
						Set MapFrame3 = shape
						shape.Name = "Color Relief Map  Wide"
						shape.xMapPerPU = scalingratio / widescale
						shape.yMapPerPU = scalingratio
					End If
				Next

				mf3height = mapframe3.Height
				mf3width = MapFrame3.Width

				maptop = mf1height + mapbottommargin
				mapwidth = (xMax - xMin) / scalingratio
				mapcellwidth = mf1width + mapsidemargin
				widemapcellwidth = mf3width + mapsidemargin
				scaletop = maptop - scaletopmargin
				scalecellwidth = cs1width + scalesidemargin
				cellswidth = 2 * mapcellwidth + 2 * scalecellwidth + widemapcellwidth + spacecellwidth
				cellsleft = (pagewidth - cellswidth) / 2

				scalecell1left = cellsleft
				scalecell1center = scalecell1left + scalecellwidth / 2
				scalecell1right = scalecell1left + scalecellwidth

				mapcell1left = scalecell1right
				mapcell1center = mapcell1left + mapcellwidth / 2
				mapcell1right = mapcell1left + mapcellwidth

				mapcell2left = mapcell1right + spacecellwidth
				mapcell2center = mapcell2left + mapcellwidth / 2
				mapcell2right = mapcell2left + mapcellwidth

				scalecell2left = mapcell2right
				scalecell2center = scalecell2left + scalecellwidth / 2
				scalecell2right = scalecell2left + scalecellwidth

				mapcell3left = scalecell2right
				mapcell3center = mapcell3left + widemapcellwidth / 2
				mapcell3right = mapcell3left + widemapcellwidth

				Plot.Selection.DeselectAll
				map1left = mapcell1center - mf1width / 2
				With MapFrame1
					.Selected = True
					Plot.Selection.Left = map1left
					Plot.Selection.Top = maptop
					.Selected = False
				End With

				scale1left = scalecell1center - cs1width / 2 + scaleoffset
				With ColorScale1
					.Selected = True
					Plot.Selection.Left = scale1left
					Plot.Selection.Top = scaletop
					.Selected = False
				End With

				map2left = mapcell2center - mf2width / 2
				With MapFrame2
					.Selected = True
					Plot.Selection.Left = map2left
					Plot.Selection.Top = maptop
					.Selected = False
				End With

				scale2left = scalecell2center - cs2width / 2 + scaleoffset
				With ColorScale2
					.Selected = True
					Plot.Selection.Left = scale2left
					Plot.Selection.Top = scaletop
					.Selected = False
				End With

				' position color relief wide
				map3left = mapcell3center - mf3width / 2
				With MapFrame3
					.Selected = True
					Plot.Selection.Left = map3left
					Plot.Selection.Top = maptop
					.Selected = False
				End With

				' add title text box
				title = "\fs200 " + titletext + "\n "
				title = title + "\fs100   \n \fs160 " + folder + " -" + Str(xMax - xMin) + "m @" + Str(xInt) + "m spacing - Depth:" + Str(yMax) + "m"
				Set TitleBox = Shapes.AddText(x:=pagewidth/2, y:=pageheight-0.25, Text:=title)
				TitleBox.Name = "Title"
				TitleBox.Font.hAlign = srfTACenter

				boxtext = "Passive Seismic Profiles\n " + operator + "\n " + getNowStr("dd.mm.yyyy")
				Set StampBox = Shapes.AddText(x:=3.5, y:=pageheight-1, Text:=boxtext)
				StampBox.Name = "Stamp"
				StampBox.Font.Size = 14
				StampBox.Font.hAlign = srfTACenter

				' save the plot
				Plot.SaveAs(FileName:=plotfile)

			End If

		Next

		Debug.Print ""
		Print #logf, ""

	Next

    Debug.Print "RAP processing completed"
	Print #logf, "RAP processing completed"
	Print #logf
	Print #logf, "End Log - " + getNowStr()
	Print #logf, logdiv

	' close logfile
	Close #logf

	If Not leaveopen Then
		surf.Quit
	End If

End Sub

' returns the value of the config parameter from the config array
Function getConfigVal(key As String, configArray As Variant, dtype As String) As Variant
	getConfigVal = Null
	For Each pair In configArray
		If pair(0) = key Then
			If dtype = "str" Then
				getConfigVal = pair(1)
			ElseIf dtype = "val" Then
				getConfigVal = Val(pair(1))
			ElseIf dtype = "bool" Then
				getConfigVal = CBool(pair(1))
			End If
			Exit For
		End If
	Next
End Function

' returns an array containing the values from the config file
Function readConfigFile(cwd As String) As Variant
	Dim configArray() As Variant
	index = 0

    ' Specify the path to the configuration file 
    configFile = cwd + "\config.txt"

    ' Check if the file exists 
    If Dir(configFile) = "" Then
        MsgBox "Config file not found!" 
        Exit Function
    End If 

    ' Open and read file
    cfile = FreeFile
    Open configFile For Input As #cfile
    Do While Not EOF(cfile)
        ' Read the entire line into the variable 
        Line Input #cfile, linevals

        ' Split the line into key-value pairs 
        keyValue = Split(linevals, "=")

        ' Process key-value pairs 
        If UBound(keyValue) = 1 Then
			ReDim Preserve configArray(index)
			configArray(index) = Array(Trim(keyValue(0)), Trim(keyValue(1)))
			index = index + 1
        End If
    Loop
	Set readConfigFile = configArray
End Function

' returns the current date/time as a formatted string
Function getNowStr(Optional fmt As String = "yyyy.mm.dd hh:nn:ss")
	getNowStr = Format(Now, fmt)
End Function

' returns an array of the folder names in path
Function GetDirs(path As String) As Variant
	Dim dirs() As String
	index = 0
	file = Dir(path, 16)
	While file <> ""
		If InStr(file, ".") = 0 Then
			ReDim Preserve dirs(index)
			dirs(index) = file
			index = index + 1
		End If
		file = Dir()
	Wend
	GetDirs = dirs
End Function

' returns and array of the file names found in path matching pattern
Function GetFiles(path As String , Optional pattern As String = "") As Variant
	Dim files() As String
	index = 0
	file = Dir(path + pattern)
	While file <> ""
		ReDim Preserve files(index)
		files(index) = file
		index = index + 1
		file = Dir()
	Wend
	GetFiles = files
End Function
