' script for creating surfer plots
Sub Main

	' configuration settings
	background = True
	leaveopen = True
	pageWidth = 17
	pageHeight = 11
    defaultDir = "C:\_Working\"
	' map size
	mapheight = 8
	maptop = mapheight + 1
	mapmargin = 1
	widescale = 4
	' scale size
	scaletop = maptop - 1
	scalecellwidth = 1.75
	scaleoffset = 0.2
	' space size
	spacecellwidth = 1

	'axis labels
	showtopaxislabels = True
	showbottomaxislabels = True
	showleftaxislabels = True
	showrightaxislabels = False
	' axis titles
	showtopaxistitle = False
	showbottomaxistitle = True
	showleftaxistitle = True
	showrightaxistitle = False
	axistitleoffset = -0.05
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

	Debug.Clear

	Open CurDir() + "\logfile.txt" For Append As #1
	Print #1, "************************************************"
	Print #1, "Start Log - " + Str(Now)
	Print #1, ""

	Set surf = CreateObject("surfer.application")
	surf.Visible = Not background

	If Command$ = "" Then
		rootdir = defaultDir
	Else
		rootdir = Command$
	End If
	Print #1, "Processing files in: " + rootdir
	Print #1, ""
	datafiledir = rootdir + "Gamma\"
	surfdir = rootdir + "Surfer\"
	kmldir = rootdir + "KML\"
	rapdir = rootdir + "RAP\"

	Debug.Print "Processing gamma"
	Debug.Print ""

    Print #1, "Processing gamma"
    Print #1, ""

	datafiles = GetFiles(datafiledir, "*_Edit.txt")
	For Each datafile In datafiles
		Debug.Print "Found datafile: " + datafiledir + datafile
		Print #1, "Found datafile: " + datafiledir + datafile
		filename = Split(datafile, ".")(0)
		gridfile = datafiledir + filename + ".grd"
		If Dir(gridfile) <> "" Then
			Debug.Print "Found gridfile: " + gridfile
			Print #1, "Found gridfile: " + gridfile
		Else
			Debug.Print "Creating grid file: " + gridfile
			Print #1, "Creating grid file: " + gridfile
			' create grid file
			surf.GridData(DataFile:= datafiledir + datafile,  xCol:=4, yCol:=3, zCol:=2, Algorithm:=srfKriging, ShowReport:=False, OutGrid:=gridfile)
		End If

		plotfile = surfdir + filename + ".srf"
		If Dir(plotfile) <> "" Then
			Debug.Print "Found plotfile: " + plotfile
			Print #1, "Found plotfile: " + plotfile
		Else
			Debug.Print "Creating plotfile: " + plotfile
			Print #1, "Creating plotfile: " + plotfile
			' create plot object
			Set Plot = surf.Documents.Add(srfDocPlot)
			Set PageSetup = Plot.PageSetup
			PageSetup.Width = pageWidth
			PageSetup.Height = pageHeight
	
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
			If Dir(kmlfile) <> "" Then
				Debug.Print "Found kmlfile: " + kmlfile
				Print #1, "Found kmlfile: " + kmlfile
			Else
				Debug.Print "Creating kmlfile: " + kmlfile
				Print #1, "Creating kmlfile: " + kmlfile
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

			' save the plot
			Plot.SaveAs(FileName:=plotfile)
		End If

		datafile = Dir()
		Debug.Print ""
		Print #1, ""

	Next

	Debug.Print "Gamma processing completed"
    Debug.Print ""
	Print #1, "Gamma processing completed"
    Print #1, ""

	Debug.Print "Processing RAP"
	Debug.Print ""
    Print #1, "Processing RAP"
    Print #1, ""

	folders = GetDirs(rapdir)
	For Each folder In folders
		Debug.Print "Found folder: " + folder
		Print #1, "Found folder: " + folder
		datadir = rapdir + folder + "\Data_Cor\"
		datafiles = GetFiles(datadir, "*frn.txt")

		For Each datafile In datafiles
			Debug.Print "Found datafile: " + datafile
			Print #1, "Found datafile: " + datafile

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

			scalingratio = (yMax - yMin) / mapheight
			mapwidth = (xMax - xMin) / scalingratio
			mapcellwidth = mapwidth + mapmargin
			widemapwidth = mapwidth * widescale
			widemapcellwidth = widemapwidth + mapmargin
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

			' close datafile
			wks.Close

			filename = Split(datafile, ".")(0)
			gridfile = datadir + filename + ".grd"
			If Dir(gridfile) <> "" Then
				Debug.Print "Found gridfile: " + gridfile
				Print #1, "Found gridfile: " + gridfile
			Else
				Debug.Print "Creating gridfile: " + gridfile
				Print #1, "Creating gridfile: " + gridfile
				NumCols = 2 * (xMax - xMin) / xInt + 1
				surf.GridData6(DataFile:= datadir + datafile,  xCol:=1, yCol:=2, zCol:=3, NumCols:=NumCols, NumRows:=101, xMin:=xMin, xMax:=xMax, yMin:=yMin, yMax:=yMax, Algorithm:=srfKriging, ShowReport:=False, OutGrid:=gridfile)
			End If

			plotfile = surfdir + folder + " " + Split(filename, "frn")(0) + ".srf"
			If Dir(plotfile) <> "" Then
				Debug.Print "Found plotfile: " + plotfile
				Print #1, "Found plotfile: " + plotfile
			Else
				Debug.Print "Creating plotfile: " + plotfile
				Print #1, "Creating plotfile: " + plotfile
				' create plot object
				Set Plot = surf.Documents.Add(srfDocPlot)
				Set PageSetup = Plot.PageSetup
				PageSetup.Width = pagewidth
				PageSetup.Height = pageheight
		
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

				' position contour map
				mf1width = MapFrame1.Width
				map1left = mapcell1center - mf1width / 2
				With MapFrame1
					.Selected = True
					Plot.Selection.Left = map1left
					Plot.Selection.Top = maptop
					.Selected = False
				End With

				' position color scale
				cs1width = ColorScale1.Width
				scale1left = scalecell1center - cs1width / 2 + scaleoffset
				With ColorScale1
					.Selected = True
					Plot.Selection.Left = scale1left
					Plot.Selection.Top = scaletop
					.Selected = False
				End With

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

				' position color relief map
				mf2width = MapFrame2.Width
				map2left = mapcell2center - mf2width / 2
				With MapFrame2
					.Selected = True
					Plot.Selection.Left = map2left
					Plot.Selection.Top = maptop
					.Selected = False
				End With

				' position color scale
				cs2width = ColorScale2.Width
				scale2left = scalecell2center - cs2width / 2 + scaleoffset
				With ColorScale2
					.Selected = True
					Plot.Selection.Left = scale2left
					Plot.Selection.Top = scaletop
					.Selected = False
				End With

				' make a wide copy of color relief map
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

				' position color relief wide
				mf3width = MapFrame3.Width
				map3left = mapcell3center - mf3width / 2
				Plot.Selection.Left = map3left
				Plot.Selection.Top = maptop
				Plot.Selection.DeselectAll

				' add title text box
				title = "\fs180 " + projectName + "\n "
				title = title + "\fs160 " + customerName + "\n "
				title = title + "\fs140 " + folder + " -" + Str(xMax - xMin) + "m @" + Str(xInt) + "m spacing - Depth:" + Str(yMax) + "m"
				Set TitleBox = Shapes.AddText(x:=8.5, y:=10.75, Text:=title)
				TitleBox.Font.hAlign = srfTACenter

				' save the plot
				Plot.SaveAs(FileName:=plotfile)

			End If

		Next

		Debug.Print ""
		Print #1, ""

	Next

    Debug.Print "RAP processing completed"
	Print #1, "RAP processing completed"
	Print #1, "End Log - " + Str(Now)
	Print #1, "************************************************"

	' close logfile
	Close #1

	If Not leaveopen Then
		surf.Quit
	End If

End Sub

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