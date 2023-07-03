Sub Main()

	Dim partsFolder As String = "D:\01 Academia BIM\BIM A+7\Parametrizacija"
	Dim oDoc As AssemblyDocument = ThisApplication.ActiveDocument
	Dim oPattern As RectangularOccurrencePattern

	Dim PlotWidth As Integer
	Dim PlotLength As Integer
	'Dim WidthofTunnel As Integer
	Dim NumberofTunnels As Integer
	Dim GreenhouseWidth As Integer

	Dim SideColumns As Integer
	Dim CentralColumns As Integer

	' Define List to store the inputs
	Dim list As New List(Of Integer)()
	Dim index As Integer
	index = 0
	' Show remaining items of the list	
	Dim sResult As String = ""

	' BASE NAME 
	Dim baseNameColumn As String = "PG808024700 - Pilar Galvanizado Quadrado 80x2.0mm"

	Dim baseNameCapitel As String = "ABG0131 - Abraçadeira Galvanizada Dupla De V Para Pilar  80mmx80mm"

	Dim baseNamePattern As String = "RectPatternOuter"
	Dim baseNamePatternInner As String = "RectPatternInner"

	Dim baseNamePatternCapitel As String = "RectPatternCapitelDouble"


	Dim baseNamePatternRoof As String = "Roof"
	Dim baseNamePatternRoofSimple As String = "RoofSimple"

	Dim baseNamePatternGutter As String = "Gutter"

	' CURRENT INDEX
	Dim currentIndexColumn As Integer
	Dim currentIndexColumnInner As Integer

	Dim currentIndexPattern As Integer

	Dim currentIndexCapitel As Integer



	' LAST ELEMENT 
	Dim lastColumn As String
	Dim lastColumnInner As String

	Dim lastCapitelDouble As String

	' PATTERN NAME
	Dim patternName As String
	Dim patternNameInner As String

	Dim patternNameCapitelDouble As String

	Dim patternNameRoof As String
	Dim patternNameRoofSimple As String

	Dim patternNameGutter As String



	'	 COLUMN PLACEMENT ******************************************************************************************************

	Dim ColumnOuter = Components.Add("PG808024700 - Pilar Galvanizado Quadrado 80x2.0mm", partsFolder & "\PG808024700 - Pilar Galvanizado Quadrado 80x2.0mm.ipt", position:=Nothing, grounded:=True, visible:=True, appearance:=Nothing)


	'	 COLUMN DIMENSIONS ******************************************************************************************************
	Parameter("PG808024700 - Pilar Galvanizado Quadrado 80x2.0mm", "ColumnWidth") = ColumnLength2
	Parameter("PG808024700 - Pilar Galvanizado Quadrado 80x2.0mm", "ColumnLength") = ColumnLength1
	Parameter("PG808024700 - Pilar Galvanizado Quadrado 80x2.0mm", "ColumnHeight") = ColumnHeight
	Parameter("PG808024700 - Pilar Galvanizado Quadrado 80x2.0mm", "ColumnThickness") = ColumnThickness


	iLogicVb.RunRule("PG808024700 - Pilar Galvanizado Quadrado 80x2.0mm", "ColumnDimensions")
	ColumnThickness = 3

	'ColumnLength2 = InputListBox("Import Column Width", MultiValue.List("ColumnLength2"), ColumnLength2, Title := "Column Width", ListName := "ColumnLength2")
	'ColumnLength1 = InputListBox("Import Column Length", MultiValue.List("ColumnLength1"), ColumnLength1, Title := "Column Length", ListName := "ColumnLength1")
	'ColumnThickness = InputListBox("Import Column Thickness", MultiValue.List("ColumnThickness"), ColumnThickness, Title := "Column Thickness", ListName := "ColumnThickness")
	'ColumnHeight = InputListBox("Import Column Height", MultiValue.List("ColumnHeight"), ColumnHeight, Title := "ColumnHeight", ListName := "ColumnHeight")
	'ColumnSection = InputListBox("Choose column section", MultiValue.List("ColumnSection"), ColumnSection, Title := "Column Section", ListName := "ColumnSection")


	'If ColumnSection = "80 x 80" Then 
	'	ColumnLength2 = 80
	'	ColumnLength1 = 80
	'	Else
	'	ColumnLength2 = 60
	'	ColumnLength1 = 60
	'End If


	iLogicVb.UpdateWhenDone = True




	'	 PLOT inputs from the user ********************************************************************************************

	PlotWidth = InputBox("Import Width of the Plot in mm", " Width of the Plot", "Default Entry")
	PlotLength = InputBox("Import Length of the Plot in mm", " Length of the Plot", "Default Entry")
	MessageBox.Show("Plot Width = " & PlotWidth & vbCrLf & "Plot Lengths = " & PlotLength, "Value of Plot Width and Length")


	'	 TUNNELS inputs from the user ********************************************************************************************

	Do Until PlotWidth >= GreenhouseWidth And GreenhouseWidth >= PlotWidth - 8000

		MessageBox.Show("Greenhouse Width=" & GreenhouseWidth, "Title")

		WidthofTunnel = InputListBox("Choose width of the Tunnel", MultiValue.List("WidthofTunnel"), WidthofTunnel, Title:="WidthofTunnel", ListName:="WidthofTunnel")
		NumberofTunnels = InputBox("Enter Number of Tunnels", "Number of Tunnels", "Default Entry")

		GreenhouseWidth = GreenhouseWidth + WidthofTunnel * NumberofTunnels

		' Error message - Greenhouse larger then Plot
		If PlotWidth < GreenhouseWidth Then
			GreenhouseWidth = 0
			MessageBox.Show("Enter Values within the plot", "Warning")
		End If

		list.Add(WidthofTunnel)
		list.Add(NumberofTunnels)

	Loop

	MessageBox.Show("Greenhouse Width=" & GreenhouseWidth, "Greenhouse Width")



	' Showing stored list values   
	For Each elem As String In list
		sResult &= elem & vbCrLf
	Next

	MessageBox.Show(sResult, "Tunnels")


	SideColumns = Floor(CDbl(PlotLength / 2500))
	CentralColumns = Floor(CDbl(PlotLength / 5000) - 1)




	'	1 TUNNEL WIDTH ********************************************************************************************

	If list.Count = 2 Then

		Call FirstPatternTotal(patternNameGutter, baseNamePatternGutter, patternNameRoofSimple, baseNamePatternRoofSimple, patternNameRoof, baseNamePatternRoof, patternNameCapitelDouble, baseNamePatternCapitel, oDoc, list, patternNameInner, baseNamePatternInner, CentralColumns, SideColumns, PlotLength, patternName, baseNamePattern, currentIndexPattern, partsFolder)
		currentIndexColumnInner = 1
		'		Call Y2(lastColumn ,oDoc ,list ,patternName,baseNameColumn,currentIndexColumn,currentIndexColumnInner, SideColumns)

	End If

	'	2 TUNNEL WIDTHS ********************************************************************************************	

	If list.Count = 4 Then

		Call FirstPatternTotal(patternNameGutter, baseNamePatternGutter, patternNameRoofSimple, baseNamePatternRoofSimple, patternNameRoof, baseNamePatternRoof, patternNameCapitelDouble, baseNamePatternCapitel, oDoc, list, patternNameInner, baseNamePatternInner, CentralColumns, SideColumns, PlotLength, patternName, baseNamePattern, currentIndexPattern, partsFolder)
		Call SuppressColumnLast(patternName, oDoc)
		Call SuppressColumnInner(patternNameInner, CentralColumns, oDoc)
		'		Call SuppressCapitel(SideColumns,patternNameCapitelDouble,oDoc )
		Call SecondPatternTotal(baseNameCapitel, lastCapitelDouble, currentIndexCapitel, patternNameCapitelDouble, baseNamePatternCapitel, lastColumnInner, lastColumn, patternNameInner, patternName, currentIndexColumnInner, currentIndexColumn, currentIndexPattern, baseNamePatternInner, baseNamePattern, baseNameColumn, list, CentralColumns, SideColumns, partsFolder, oDoc, PlotLength)
		'		Call Y2(lastColumn ,oDoc ,list ,patternName,baseNameColumn,currentIndexColumn,currentIndexColumnInner, SideColumns)

	End If

	'	3+ TUNNEL WIDTHS ********************************************************************************************	

	If list.Count >= 6 Then

		Call FirstPatternTotal(patternNameGutter, baseNamePatternGutter, patternNameRoofSimple, baseNamePatternRoofSimple, patternNameRoof, baseNamePatternRoof, patternNameCapitelDouble, baseNamePatternCapitel, oDoc, list, patternNameInner, baseNamePatternInner, CentralColumns, SideColumns, PlotLength, patternName, baseNamePattern, currentIndexPattern, partsFolder)
		Call SuppressColumnLast(patternName, oDoc)
		Call SuppressColumnInner(patternNameInner, CentralColumns, oDoc)
		'		Call SuppressCapitel(SideColumns,patternNameCapitelDouble,oDoc )
		Call SecondPatternTotal(baseNameCapitel, lastCapitelDouble, currentIndexCapitel, patternNameCapitelDouble, baseNamePatternCapitel, lastColumnInner, lastColumn, patternNameInner, patternName, currentIndexColumnInner, currentIndexColumn, currentIndexPattern, baseNamePatternInner, baseNamePattern, baseNameColumn, list, CentralColumns, SideColumns, partsFolder, oDoc, PlotLength)
		Call SuppressColumnLast(patternName, oDoc)
		Call SuppressColumnInner(patternNameInner, CentralColumns, oDoc)
		'		Call SuppressCapitel(SideColumns,patternNameCapitelDouble,oDoc )
		Call ThirdPatternTotal(currentIndexCapitel, baseNameCapitel, lastCapitelDouble, patternNameCapitelDouble, baseNamePatternCapitel, lastColumnInner, lastColumn, patternNameInner, patternName, currentIndexColumnInner, currentIndexColumn, currentIndexPattern, baseNamePatternInner, baseNamePattern, baseNameColumn, list, CentralColumns, SideColumns, partsFolder, oDoc, PlotLength)
		'		Call Y2(lastColumn ,oDoc ,list ,patternName,baseNameColumn,currentIndexColumn,currentIndexColumnInner, SideColumns)

	End If




	'	FIRST LINE Y COLUMN ********************************************************************************************		


	Dim posY1 = ThisDoc.Geometry.Point(0, 2500, 0)
	Components.Add("PG808024700 - Pilar Galvanizado Quadrado 80x2.0mmY", partsFolder & "\PG808024700 - Pilar Galvanizado Quadrado 80x2.0mm.ipt", position:=posY1, grounded:=True, visible:=True, appearance:=Nothing)

	Dim rectPatternY = Patterns.AddRectangular("Y1", "PG808024700 - Pilar Galvanizado Quadrado 80x2.0mmY", SideColumns - 1, 2500, Nothing, "Y Axis", columnNaturalDirection:=True)


	'	CLAMPS HORIZONTAL PROFILES ********************************************************************************************	

	'		Dim posClamps = ThisDoc.Geometry.Point(0, 0, 200)
	'				Components.Add("Clampse80x80", partsFolder & "\Clamps for Horizontal Profiles 80x80.iam", position := posClamps, grounded:=False, visible:=True, appearance:=Nothing)

	'		Constraints.AddMate("Mate:6", "PG808024700 - Pilar Galvanizado Quadrado 80x2.0mm",
	'                            "Face1",
	'                            {"Clampse80x80", "clamps for horizontal profiles:1", "ABM0038 - Abraçadeira Magnélis Lateral Para Calha EM3540 Quadrado 80mm:2"},
	'                            "Face0")
	'        Constraints.AddMate("Mate:7", "PG808024700 - Pilar Galvanizado Quadrado 80x2.0mm",
	'                            "Face0",
	'                            {"Clampse80x80", "clamps for horizontal profiles:1", "ABM0038 - Abraçadeira Magnélis Lateral Para Calha EM3540 Quadrado 80mm:1"},
	'                            "Face1")
	'  Constraints.AddMate("Mate1", "Part1:1", "Work Plane1", "Part2:1", "Work Plane1",
	'                    offset := 0.0, e1InferredType := InferredTypeEnum.kNoInference, e2InferredType := InferredTypeEnum.kNoInference,
	'                    solutionType := MateConstraintSolutionTypeEnum.kNoSolutionType,
	'                    biasPoint1 := Nothing, biasPoint2 := Nothing)


	'Select Case ColumnHeight
	'	Case 5000
	'	Dim rectPattern = Patterns.AddRectangular("patternClamps80x80", "Clampse80x80", SideColumns + 1, 2500, Nothing, "Y Axis", columnNaturalDirection := True, rowCount := 5, rowOffset := 1000, rowEntityName := "Z Axis", rowNaturalDirection := True)

	'	Case 4500
	'	Dim rectPattern = Patterns.AddRectangular("patternClamps80x80", "Clampse80x80", SideColumns + 1, 2500, Nothing, "Y Axis", columnNaturalDirection := True, rowCount := 4, rowOffset := 1120, rowEntityName := "Z Axis", rowNaturalDirection := True)

	'	Case 4000
	'	Dim rectPattern = Patterns.AddRectangular("patternClamps80x80", "Clampse80x80", SideColumns + 1, 2500, Nothing, "Y Axis", columnNaturalDirection := True, rowCount := 4, rowOffset := 1000, rowEntityName := "Z Axis", rowNaturalDirection := True)

	'	Case 3500
	'	Dim rectPattern = Patterns.AddRectangular("patternClamps80x80", "Clampse80x80", SideColumns + 1, 2500, Nothing, "Y Axis", columnNaturalDirection := True, rowCount := 3, rowOffset := 1150, rowEntityName := "Z Axis", rowNaturalDirection := True)

	'	Case 3000
	'	Dim rectPattern = Patterns.AddRectangular("patternClamps80x80", "Clampse80x80", SideColumns + 1, 2500, Nothing, "Y Axis", columnNaturalDirection := True, rowCount := 3, rowOffset := 1000, rowEntityName := "Z Axis", rowNaturalDirection := True)

	'	Case 2500
	'	Dim rectPattern = Patterns.AddRectangular("patternClamps80x80", "Clampse80x80", SideColumns+1, 2500, Nothing, "Y Axis", columnNaturalDirection := True, rowCount := 2, rowOffset := 1200, rowEntityName := "Z Axis", rowNaturalDirection := True)



	'End Select



End Sub




Function FirstPatternTotal(patternNameGutter As String, baseNamePatternGutter As String, patternNameRoofSimple As String, baseNamePatternRoofSimple As String, patternNameRoof As String, baseNamePatternRoof As String, ByRef patternNameCapitelDouble As String, baseNamePatternCapitel As String, oDoc As AssemblyDocument, list As List(Of Integer), ByRef patternNameInner As String, baseNamePatternInner As String, CentralColumns As Integer, SideColumns As Integer, PlotLength As Integer, ByRef patternName As String, ByVal baseNamePattern As String, ByRef currentIndexPattern As Integer, partsFolder As String)


	' COLUMN OUTER /////////////////////////////////////////////////////////////////////////////////////////////////
	' Pattern Name - Index 		
	currentIndexPattern = list(0)
	' New pattern name 
	patternName = Nothing
	patternName = baseNamePattern & ":" & currentIndexPattern.ToString()
	MessageBox.Show("Pattern name = " & patternName, "Pattern name")

	' Pattern		
	Dim rectPattern2 = Patterns.AddRectangular(patternName, "PG808024700 - Pilar Galvanizado Quadrado 80x2.0mm", list(1) + 1, list(0), Nothing, "X Axis", columnNaturalDirection:=True, rowCount:=2, rowOffset:=PlotLength, rowEntityName:="Y Axis", rowNaturalDirection:=True)

	'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

	' COLUMN INNER /////////////////////////////////////////////////////////////////////////////////////////////////
	Dim pos = ThisDoc.Geometry.Point(list(0), 5000, 0)
	Components.Add("PG808024700 - Pilar Galvanizado Quadrado 80x2.0mmInner", partsFolder & "\PG808024700 - Pilar Galvanizado Quadrado 80x2.0mm.ipt", position:=pos, grounded:=True, visible:=True, appearance:=Nothing)
	ColumnThicknessInner = 2
	'		Parameter("PG808024700 - Pilar Galvanizado Quadrado 80x2.0mmInner", "ColumnThickness") = ColumnThicknessInner

	'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

	' Pattern Name - Index 		
	currentIndexPattern = list(0)
	' New pattern name 
	patternNameInner = Nothing
	patternNameInner = baseNamePatternInner & ":" & currentIndexPattern.ToString()
	MessageBox.Show("Pattern name = " & patternNameInner, "Pattern name")

	If list.Count = 2 Then
		Dim rectPatternInnerColumn = Patterns.AddRectangular(patternNameInner, "PG808024700 - Pilar Galvanizado Quadrado 80x2.0mmInner", list(1) - 1, list(0), Nothing, "X Axis", columnNaturalDirection:=True, rowCount:=CentralColumns, rowOffset:=5000, rowEntityName:="Y Axis", rowNaturalDirection:=True)
	Else
		Dim rectPatternInnerColumn = Patterns.AddRectangular(patternNameInner, "PG808024700 - Pilar Galvanizado Quadrado 80x2.0mmInner", list(1), list(0), Nothing, "X Axis", columnNaturalDirection:=True, rowCount:=CentralColumns, rowOffset:=5000, rowEntityName:="Y Axis", rowNaturalDirection:=True)
	End If

	'		' CAPITEL /////////////////////////////////////////////////////////////////////////////////////////////////
	'		Dim posDoubleCapitel = ThisDoc.Geometry.Point(list(0), 0, ColumnHeight)
	'		Components.Add("CapitelDouble80x80", partsFolder & "\ABG0131 - Abraçadeira Galvanizada Dupla De V Para Pilar  80mmx80mm.iam", position := posDoubleCapitel, grounded:=False, visible:=True, appearance:=Nothing)

	'		'Constraints Capitel
	'		Constraints.AddMate("393a2a5b-bdd9-4a6b-802c-b7b2d0910217",
	'		                    "PG808024700 - Pilar Galvanizado Quadrado 80x2.0mm:3",
	'		                    "Edge1",
	'		                    {"CapitelDouble80x80", "ABB0078 - Abraçadeira em Bruto Simples De V 80mm -  Dupla:1"},
	'		                    "Edge0",
	'		                    e1InferredType := InferredTypeEnum.kInferredPoint,
	'		                    e2InferredType := InferredTypeEnum.kInferredPoint)


	'		' Pattern Name - Index 		
	'		currentIndexPattern = list (0)
	'		' New pattern name 
	'		patternNameCapitelDouble = Nothing
	'		patternNameCapitelDouble = baseNamePatternCapitel & ":" & currentIndexPattern.ToString()
	'		MessageBox.Show("Pattern name = " & patternNameCapitelDouble, "Pattern name")

	'		If list.Count=2 Then 
	'		Dim rectPatternCapitelDouble = Patterns.AddRectangular(patternNameCapitelDouble, "CapitelDouble80x80", list(1)-1, list(0), Nothing, "X Axis", columnNaturalDirection := True, rowCount := SideColumns + 1, rowOffset := 2500, rowEntityName := "Y Axis", rowNaturalDirection := True)
	'		Else
	'		Dim rectPatternCapitelDouble = Patterns.AddRectangular(patternNameCapitelDouble, "CapitelDouble80x80", list(1), list(0), Nothing, "X Axis", columnNaturalDirection := True, rowCount := SideColumns + 1, rowOffset := 2500, rowEntityName := "Y Axis", rowNaturalDirection := True)
	'		End If

	'		' ROOF STRUCTURE /////////////////////////////////////////////////////////////////////////////////////////////////

	''		Select Case WidthofTunnel 
	''		Case 10000
	'		Dim posRoof = ThisDoc.Geometry.Point(10000, 5000, ColumnHeight)
	'		Components.Add("Roof10", partsFolder & "\Estufa 10 M Macho Fêmea Pilares Quadrado 80mm com Travamento em M-T.iam", position := posRoof, grounded:=False, visible:=True, appearance:=Nothing)

	'		Constraints.AddMate("504680b1-f17e-4ecd-b39d-21a090c2e7a5",
	'                            {"Roof10", "AR601511000 - Arco Sendzimir Ø60x1.5mm c11000mm:1"},
	'                            "Face0",
	'                            {"ABG0131 - Abraçadeira Galvanizada Dupla De V Para Pilar  80mmx80mm:3", "ABB0078 - Abraçadeira em Bruto Simples De V 80mm -  Dupla:1"},
	'                            "Face1",
	'                            e1InferredType := InferredTypeEnum.kInferredLine,
	'                            e2InferredType := InferredTypeEnum.kInferredLine,
	'                            solutionType := MateConstraintSolutionTypeEnum.kOpposedSolutionType)



	'		' Pattern Name - Index 		
	'		currentIndexPattern = list (0)
	'		' New pattern name 
	'		patternNameRoof = Nothing
	'		patternNameRoof = baseNamePatternRoof & ":" & currentIndexPattern.ToString()
	'		MessageBox.Show("Pattern name = " & patternNameRoof, "Pattern name")

	''	    End Select
	'		Dim rectPatternRoof = Patterns.AddRectangular(patternNameRoof, "Roof10", list(1), list(0), Nothing, "X Axis", columnNaturalDirection := True, rowCount :=Floor(CDbl(SideColumns /2)-1)  , rowOffset := 5000, rowEntityName := "Y Axis", rowNaturalDirection := True)


	'		' ROOF STRUCTURE SIMPLE /////////////////////////////////////////////////////////////////////////////////////////////////

	'		'Select Case WidthofTunnel 
	''		Case 10000
	'		Dim posRoofSimple = ThisDoc.Geometry.Point(10000, 2500, ColumnHeight)
	'		Components.Add("Roof10Simple", partsFolder & "\Estufa 10 M Macho Fêmea Pilares Quadrado 80mm Frontal.iam", position := posRoofSimple, grounded:=False, visible:=True, appearance:=Nothing)
	'		Constraints.AddMate("Mate:5", {"Roof10Simple", "AR601511000 - Arco Sendzimir Ø60x1.5mm c11000mm:1"},
	'                            "Face1",
	'                            {"ABG0131 - Abraçadeira Galvanizada Dupla De V Para Pilar  80mmx80mm:2", "ABB0078 - Abraçadeira em Bruto Simples De V 80mm -  Dupla:1"},
	'                            "Face1",
	'                            e1InferredType := InferredTypeEnum.kInferredLine,
	'                            e2InferredType := InferredTypeEnum.kInferredLine,
	'                            solutionType := MateConstraintSolutionTypeEnum.kAlignedSolutionType)


	'		' Pattern Name - Index 		
	'		currentIndexPattern = list (0)
	'		' New pattern name 
	'		patternNameRoofSimple = Nothing
	'		patternNameRoofSimple = baseNamePatternRoofSimple & ":" & currentIndexPattern.ToString()
	'		MessageBox.Show("Pattern name = " & patternNameRoofSimple, "Pattern name")

	''	    End Select
	'		Dim rectPatternRoofSimple = Patterns.AddRectangular(patternNameRoofSimple, "Roof10Simple", list(1), list(0), Nothing, "X Axis", columnNaturalDirection := True, rowCount :=Floor(CDbl(SideColumns /2))  , rowOffset := 5000, rowEntityName := "Y Axis", rowNaturalDirection := True)


	'		'	FIRST LINE Y CAPITEL********************************************************************************************


	'		Dim posSingleCapitel = ThisDoc.Geometry.Point(0, 0, ColumnHeight)
	'		Components.Add("CapitelSingle80x80", partsFolder & "\ABG0135 - Abraçadeira Galvanizada Lateral Simples De V Para Pilar 80mmx80mm.iam", position := posSingleCapitel, grounded:=False, visible:=True, appearance:=Nothing)

	'		'Constraints Capitel
	'		Constraints.AddMate("Mate:3", "PG808024700 - Pilar Galvanizado Quadrado 80x2.0mm",
	'                            "Edge1",
	'                            {"CapitelSingle80x80", "ABBC0036 - Abraçadeira Simples De V 80mm - Direita:1"},
	'                            "Edge0",
	'                            e1InferredType := InferredTypeEnum.kInferredPoint,
	'                            e2InferredType := InferredTypeEnum.kInferredPoint)
	'        Constraints.AddMate("Mate:4", "PG808024700 - Pilar Galvanizado Quadrado 80x2.0mm",
	'                            "Edge2",
	'                            {"CapitelSingle80x80", "ABBC0036 - Abraçadeira Simples De V 80mm - Direita:1"},
	'                            "Edge1",
	'                            e1InferredType := InferredTypeEnum.kInferredPoint,
	'                            e2InferredType := InferredTypeEnum.kInferredPoint)

	'		Dim rectPatternCapitelSingle = Patterns.AddRectangular("CapitelY1", "CapitelSingle80x80", SideColumns + 1, 2500, Nothing, "Y Axis", columnNaturalDirection := True)


	'		 ' GUTTER  /////////////////////////////////////////////////////////////////////////////////////////////////
	'		 Dim posGutter = ThisDoc.Geometry.Point(0, 0, ColumnHeight)
	'		Components.Add("Gutter", partsFolder & "\CLZ205120 - Caleiro Zincado  2.00mm com 5120mm.ipt", position := posGutter, grounded := False, visible := True, appearance := Nothing)
	'		Constraints.AddInsert("Insert:1", "Gutter", "Edge1",
	'		                      {"CapitelSingle80x80", "ABBC0119 - V de 125mm:1"},
	'		                      "Edge3", axesOpposed := False)


	'		' Pattern Name - Index 		
	'		currentIndexPattern = list (0)
	'		' New pattern name 
	'		patternNameGutter = Nothing
	'		patternNameGutter = baseNamePatternGutter & ":" & currentIndexPattern.ToString()
	'		MessageBox.Show("Pattern name = " & baseNamePatternGutter, "Pattern name")

	'		Dim rectPattern = Patterns.AddRectangular(patternNameGutter, "Gutter", list(1), list(0), Nothing, "X Axis", columnNaturalDirection := True, rowCount := Floor(CDbl(PlotLength / 5000)), rowOffset := 5000, rowEntityName := "Y Axis", rowNaturalDirection := True)



End Function



Function SecondPatternTotal(baseNameCapitel As String, ByRef lastCapitelDouble As String, ByRef currentIndexCapitel As Integer, ByRef patternNameCapitelDouble As String, baseNamePatternCapitel As String, ByRef lastColumnInner As String, ByRef lastColumn As String, ByRef patternNameInner As String, ByRef patternName As String, ByRef currentIndexColumnInner As Integer, ByRef currentIndexColumn As Integer, ByRef currentIndexPattern As Integer, ByRef baseNamePatternInner As String, ByRef baseNamePattern As String, ByRef baseNameColumn As String, ByRef list As List(Of Integer), ByRef CentralColumns As Integer, ByRef SideColumns As Integer, ByRef partsFolder As String, ByRef oDoc As AssemblyDocument, PlotLength As Integer)

	' COLUMN OUTER /////////////////////////////////////////////////////////////////////////////////////////////////
	'  Column index 
	currentIndexColumn = 1 + (list(1) * 2)
	' Last Column of the pattern 
	lastColumn = Nothing
	lastColumn = baseNameColumn & ":" & currentIndexColumn.ToString()
	MessageBox.Show("Last Column = " & lastColumn, "Last column in pattern")

	' Calling Pattern
	oPattern = oDoc.ComponentDefinition.OccurrencePatterns.Item(patternName)

	' Making Column  independant - Outer Pattern
	For k = 1 To oPattern.OccurrencePatternElements.Count
		If k = oPattern.OccurrencePatternElements.Count - 1 Then
			oPattern.OccurrencePatternElements.Item(k).Independent = True
		End If
	Next


	' COLUMN INNER /////////////////////////////////////////////////////////////////////////////////////////////////
	' Column index Inner
	currentIndexColumnInner = currentIndexColumn + 1 + CentralColumns * (list(1) - 1) + 1

	' Last Column of the pattern 
	lastColumnInner = Nothing
	lastColumnInner = baseNameColumn & ":" & currentIndexColumnInner.ToString()
	MessageBox.Show("Last Column = " & lastColumnInner, "Last column in pattern")

	' Calling Pattern Inner
	oPattern = oDoc.ComponentDefinition.OccurrencePatterns.Item(patternNameInner)

	' Making Column  independant - Inner Pattern
	For k = 1 To oPattern.OccurrencePatternElements.Count
		If k = oPattern.OccurrencePatternElements.Count - CentralColumns + 1 Then
			oPattern.OccurrencePatternElements.Item(k).Independent = True
		End If
	Next


	'		' CAPITEL /////////////////////////////////////////////////////////////////////////////////////////////////
	'		' Capitel index 
	'		currentIndexCapitel  = ((list(1) - 1) * (SideColumns + 1)) + 1
	'		' Last Column of the pattern 
	'		lastCapitelDouble  = Nothing
	'		lastCapitelDouble = baseNameCapitel & ":" & currentIndexCapitel.ToString()
	'		MessageBox.Show("Last Capitel = " & lastCapitelDouble, "Last capitel in pattern")

	'		' Calling Pattern
	'		oPattern = oDoc.ComponentDefinition.OccurrencePatterns.Item(patternNameCapitelDouble)

	'	    ' Making Capitel  independant 
	'		For k = 1 To oPattern.OccurrencePatternElements.Count
	'			If k = oPattern.OccurrencePatternElements.Count - SideColumns Then
	'				oPattern.OccurrencePatternElements.Item(k).Independent = True 
	'			End If
	'		Next

	'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


	' Delete used items of the list	
	If list.Count > 2 Then
		list.RemoveRange(0, 2)
	End If
	' Show remaining items of the list	
	sResult = Nothing
	For Each elem As String In list
		sResult &= elem & vbCrLf
	Next
	MessageBox.Show(sResult, "Remaining Tunnels")



	' COLUMN OUTER PATTERN  /////////////////////////////////////////////////////////////////////////////////////////////////		
	' Pattern Index 		
	currentIndexPattern = list(0)
	' New pattern name 
	patternName = Nothing
	patternName = baseNamePattern & ":" & currentIndexPattern.ToString()
	MessageBox.Show("Pattern name = " & patternName, "Pattern name")

	Dim rectPattern2 = Patterns.AddRectangular(patternName, lastColumn, list(1) + 1, list(0), Nothing, "X Axis", columnNaturalDirection:=True, rowCount:=2, rowOffset:=PlotLength, rowEntityName:="Y Axis", rowNaturalDirection:=True)


	' COLUMN INNER PATTERN /////////////////////////////////////////////////////////////////////////////////////////////////
	' Pattern Name INNER		
	currentIndexPattern = list(0)
	' New pattern name 
	patternNameInner = Nothing
	patternNameInner = baseNamePatternInner & ":" & currentIndexPattern.ToString()
	MessageBox.Show("Pattern name = " & patternNameInner, "Pattern name")

	If list.Count = 2 Then
		Dim rectPattern = Patterns.AddRectangular(patternNameInner, lastColumnInner, list(1), list(0), Nothing, "X Axis", columnNaturalDirection:=True, rowCount:=CentralColumns, rowOffset:=5000, rowEntityName:="Y Axis", rowNaturalDirection:=True)
	Else
		Dim rectPattern = Patterns.AddRectangular(patternNameInner, lastColumnInner, list(1) + 1, list(0), Nothing, "X Axis", columnNaturalDirection:=True, rowCount:=CentralColumns, rowOffset:=5000, rowEntityName:="Y Axis", rowNaturalDirection:=True)
	End If


	'		' CAPITEL PATTERN /////////////////////////////////////////////////////////////////////////////////////////////////
	'		' Pattern Name CAPITEL		
	'		' Pattern Name - Index 		
	'		currentIndexPattern = list (0)
	'		' New pattern name 
	'		patternNameCapitelDouble = Nothing
	'		patternNameCapitelDouble = baseNamePatternCapitel & ":" & currentIndexPattern.ToString()
	'		MessageBox.Show("Pattern name = " & patternNameCapitelDouble, "Pattern name")

	'		If list.Count=2 Then 
	'		Dim rectPatternCapitelDouble = Patterns.AddRectangular(patternNameCapitelDouble, lastCapitelDouble, list(1), list(0), Nothing, "X Axis", columnNaturalDirection := True, rowCount := SideColumns+1, rowOffset := 2500, rowEntityName := "Y Axis", rowNaturalDirection := True)
	'		Else
	'		Dim rectPattern3 = Patterns.AddRectangular(patternNameCapitelDouble, lastCapitelDouble, list(1)+1, list(0), Nothing, "X Axis", columnNaturalDirection := True, rowCount := SideColumns+1 , rowOffset := 2500, rowEntityName := "Y Axis", rowNaturalDirection := True)
	'		End If


	'		Return rectPattern
	'		Return rectPattern2

End Function
Function ThirdPatternTotal(ByRef currentIndexCapitel As Integer, baseNameCapitel As String, ByRef lastCapitelDouble As String, patternNameCapitelDouble As String, baseNamePatternCapitel As String, ByRef lastColumnInner As String, ByRef lastColumn As String, ByRef patternNameInner As String, ByRef patternName As String, ByRef currentIndexColumnInner As Integer, ByRef currentIndexColumn As Integer, ByRef currentIndexPattern As Integer, ByRef baseNamePatternInner As String, ByRef baseNamePattern As String, ByRef baseNameColumn As String, ByRef list As List(Of Integer), ByRef CentralColumns As Integer, ByRef SideColumns As Integer, ByRef partsFolder As String, ByRef oDoc As AssemblyDocument, PlotLength As Integer)

	' COLUMN OUTER /////////////////////////////////////////////////////////////////////////////////////////////////
	Do While list.Count > 2
		'  Column index 
		currentIndexColumn = currentIndexColumnInner + list(1) * 2
		' Last Column of the pattern 
		lastColumn = Nothing
		lastColumn = baseNameColumn & ":" & currentIndexColumn.ToString()
		MessageBox.Show("Last Column = " & lastColumn, "Last column in pattern")

		' Calling Pattern
		oPattern = oDoc.ComponentDefinition.OccurrencePatterns.Item(patternName)

		' Making Column  independant - Outer Pattern
		For k = 1 To oPattern.OccurrencePatternElements.Count
			If k = oPattern.OccurrencePatternElements.Count - 1 Then
				oPattern.OccurrencePatternElements.Item(k).Independent = True
			End If
		Next


		' COLUMN INNER /////////////////////////////////////////////////////////////////////////////////////////////////
		' Column index Inner
		currentIndexColumnInner = currentIndexColumn + 1 + CentralColumns * list(1)
		' Last Column of the pattern 
		lastColumnInner = Nothing
		lastColumnInner = baseNameColumn & ":" & currentIndexColumnInner.ToString()
		MessageBox.Show("Last Column = " & lastColumnInner, "Last column in pattern")

		' Calling Pattern Inner
		oPattern = oDoc.ComponentDefinition.OccurrencePatterns.Item(patternNameInner)

		' Making Column  independant - Inner Pattern
		For k = 1 To oPattern.OccurrencePatternElements.Count
			If k = oPattern.OccurrencePatternElements.Count - CentralColumns + 1 Then
				oPattern.OccurrencePatternElements.Item(k).Independent = True
			End If
		Next


		'		' CAPITEL /////////////////////////////////////////////////////////////////////////////////////////////////
		'		' Capitel index 
		'		currentIndexCapitel = currentIndexCapitel + (SideColumns + 1)*list(1)
		'		' Last Column of the pattern 
		'		lastCapitelDouble  = Nothing
		'		lastCapitelDouble = baseNameCapitel & ":" & currentIndexCapitel.ToString()
		'		MessageBox.Show("Last Capitel = " & lastCapitelDouble, "Last capitel in pattern")

		'		' Calling Pattern
		'		oPattern = oDoc.ComponentDefinition.OccurrencePatterns.Item(patternNameCapitelDouble)

		'	    ' Making Capitel  independant 
		'		For k = 1 To oPattern.OccurrencePatternElements.Count
		'			If k = oPattern.OccurrencePatternElements.Count - SideColumns Then
		'				oPattern.OccurrencePatternElements.Item(k).Independent = True 
		'			End If
		'		Next


		' Delete used items of the list	
		If list.Count > 2 Then
			list.RemoveRange(0, 2)
		End If
		' Show remaining items of the list	
		sResult = Nothing
		For Each elem As String In list
			sResult &= elem & vbCrLf
		Next
		MessageBox.Show(sResult, "Remaining Tunnels")


		' COLUMN OUTER PATTERN  /////////////////////////////////////////////////////////////////////////////////////////////////	
		' Pattern Index 		
		currentIndexPattern = list(0)
		' New pattern name 
		patternName = Nothing
		patternName = baseNamePattern & ":" & currentIndexPattern.ToString()
		MessageBox.Show("Pattern name = " & patternName, "Pattern name")

		Dim rectPattern2 = Patterns.AddRectangular(patternName, lastColumn, list(1) + 1, list(0), Nothing, "X Axis", columnNaturalDirection:=True, rowCount:=2, rowOffset:=PlotLength, rowEntityName:="Y Axis", rowNaturalDirection:=True)

		' COLUMN INNER PATTERN /////////////////////////////////////////////////////////////////////////////////////////////////		
		' Pattern Name INNER		
		currentIndexPattern = list(0)
		' New pattern name 
		patternNameInner = Nothing
		patternNameInner = baseNamePatternInner & ":" & currentIndexPattern.ToString()
		MessageBox.Show("Pattern name = " & patternNameInner, "Pattern name")

		If list.Count = 2 Then
			Dim rectPattern1 = Patterns.AddRectangular(patternNameInner, lastColumnInner, list(1), list(0), Nothing, "X Axis", columnNaturalDirection:=True, rowCount:=CentralColumns, rowOffset:=5000, rowEntityName:="Y Axis", rowNaturalDirection:=True)
		Else
			Dim rectPattern = Patterns.AddRectangular(patternNameInner, lastColumnInner, list(1) + 1, list(0), Nothing, "X Axis", columnNaturalDirection:=True, rowCount:=CentralColumns, rowOffset:=5000, rowEntityName:="Y Axis", rowNaturalDirection:=True)
			Call SuppressColumnLast(patternName, oDoc)
			Call SuppressColumnInner(patternNameInner, CentralColumns, oDoc)

		End If


		'		' CAPITEL PATTERN /////////////////////////////////////////////////////////////////////////////////////////////////
		'		' Pattern Name CAPITEL		
		'		' Pattern Name - Index 		
		'		currentIndexPattern = list (0)
		'		' New pattern name 
		'		patternNameCapitelDouble = Nothing
		'		patternNameCapitelDouble = baseNamePatternCapitel & ":" & currentIndexPattern.ToString()
		'		MessageBox.Show("Pattern name = " & patternNameCapitelDouble, "Pattern name")

		'		If list.Count=2 Then 
		'		Dim rectPatternCapitelDouble = Patterns.AddRectangular(patternNameCapitelDouble, lastCapitelDouble, list(1), list(0), Nothing, "X Axis", columnNaturalDirection := True, rowCount := SideColumns+1, rowOffset := 2500, rowEntityName := "Y Axis", rowNaturalDirection := True)
		'		Else
		'		Dim rectPatternCapitelDouble = Patterns.AddRectangular(patternNameCapitelDouble, lastCapitelDouble, list(1)+1, list(0), Nothing, "X Axis", columnNaturalDirection := True, rowCount := SideColumns+1 , rowOffset := 2500, rowEntityName := "Y Axis", rowNaturalDirection := True)
		'		Call SuppressCapitel(SideColumns, patternNameCapitelDouble, oDoc)

		'		End If
	Loop

End Function


Function SuppressColumnLast(ByRef patternName As String, oDoc As AssemblyDocument)
	' Calling Pattern
	oPattern = oDoc.ComponentDefinition.OccurrencePatterns.Item(patternName)

	' Making Column before last independant
	For k = 1 To oPattern.OccurrencePatternElements.Count
		If k = oPattern.OccurrencePatternElements.Count Then
			oPattern.OccurrencePatternElements.Item(k).Suppressed = True
		End If
	Next
End Function
Function SuppressColumnInner(ByRef patternNameInner As String, ByRef CentralColumns As Integer, oDoc As AssemblyDocument)
	' Suppressing Overlaping columns Inner Pattern
	' Calling Pattern
	oPattern = oDoc.ComponentDefinition.OccurrencePatterns.Item(patternNameInner)

	' Supress
	For k = oPattern.OccurrencePatternElements.Count - CentralColumns + 2 To oPattern.OccurrencePatternElements.Count
		oPattern.OccurrencePatternElements.Item(k).Suppressed = True
	Next

End Function
Function SuppressCapitel(ByRef SideColumns As Integer, ByRef patternNameCapitelDouble As String, oDoc As AssemblyDocument)
	' Calling Pattern
	oPattern = oDoc.ComponentDefinition.OccurrencePatterns.Item(patternNameCapitelDouble)


	For k = oPattern.OccurrencePatternElements.Count - SideColumns + 1 To oPattern.OccurrencePatternElements.Count
		oPattern.OccurrencePatternElements.Item(k).Suppressed = True
	Next
End Function

Function Y2(ByRef lastColumn As String, oDoc As AssemblyDocument, list As List(Of Integer), ByRef patternName As String, ByRef baseNameColumn As String, ByRef currentIndexColumn As Integer, ByRef currentIndexColumnInner As Integer, ByRef SideColumns As Integer)
	'	Last Line Y 

	' Calling Pattern
	Dim oPattern As OccurrencePattern
	oPattern = oDoc.ComponentDefinition.OccurrencePatterns.Item(patternName)

	' Making Column before last independant
	For k = 1 To oPattern.OccurrencePatternElements.Count
		If k = oPattern.OccurrencePatternElements.Count - 1 Then
			oPattern.OccurrencePatternElements.Item(k).Independent = True
		End If
	Next

	'  Column index 
	currentIndexColumn = currentIndexColumnInner + (list(1) * 2)

	' Last Column of the pattern 
	lastColumn = Nothing
	lastColumn = baseNameColumn & ":" & currentIndexColumn.ToString()
	MessageBox.Show("Last Column = " & lastColumn, "Last column in pattern")

	Dim rectPatternY2 = Patterns.AddRectangular("Y2", lastColumn, SideColumns, 2500, Nothing, "Y Axis", columnNaturalDirection:=True)

End Function