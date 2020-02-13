Attribute VB_Name = "Verbatim"
Option Base 1
Option Explicit

Dim i, j, k, l As Long
Public wb As Workbook

Public wsVerb As Worksheet
Public wsData As Worksheet
Public wsFrame As Worksheet
Public wsUpdate As Worksheet
Public wsInfo As Worksheet
Public wsCSV As Worksheet


Public aVerb As Variant
Public aData As Variant
Public aCoding As Variant
Public aFrame As Variant
Public aInfo As Variant
Public acsv As Variant


Sub TransposeVerbatim()
    
    Call TurnOffStuf
    
    
    Set wb = ActiveWorkbook
    Set wsVerb = wb.Sheets("Verbatim")
    Set wsInfo = wb.Sheets("Info")
       
    wsVerb.Cells.ClearFormats
    
    aVerb = wsVerb.UsedRange
    aInfo = wsInfo.Range("A3").CurrentRegion
          
    'Cek Duplicate Header
    wsVerb.Activate
    For i = 1 To UBound(aVerb, 2)
        On Error Resume Next
        If WorksheetFunction.CountIf(wsVerb.Range(Cells(1, 1), Cells(1, UBound(aVerb, 2))), wsVerb.Cells(1, i).Value) > 1 Then
            MsgBox "Duplicate..." & Chr(10) & "Header - " & wsVerb.Cells(1, i).Value & ", in cells - " & wsVerb.Cells(1, i).Address
            Exit Sub
        End If
    Next i
    
    Call FormatingData
    Set wsData = wb.Sheets("Data")
    
    With CommandBars.ActionControl
    
        If .Parameter = "All verbatim" Then
            k = 4
            For i = 2 To UBound(aVerb, 2)
                For j = 2 To UBound(aVerb, 1)
                    
                    On Error Resume Next
                    aVerb(j, i) = Application.WorksheetFunction.Clean(Trim(aVerb(j, i)))
                    
                    If Err.Number <> 0 Then
                        MsgBox "Error in cells - " & wsVerb.Cells(j, i).Address
                        Application.Goto wsVerb.Cells(j, i), True
                        Application.DisplayAlerts = False
                        Sheets("data").Delete
                        Application.DisplayAlerts = True
                        Exit Sub
                    End If
        
                    If aVerb(j, i) <> "" Then
                        wsData.Cells(k, 1) = aVerb(j, 1)
                        wsData.Cells(k, 2) = aVerb(1, i)
                        wsData.Cells(k, 3) = aVerb(j, i)
                        wsData.Cells(k, 3) = aVerb(j, i)
                        'wsData.Cells(k, 17) = aVerb(j, 1) & aVerb(1, i)
                        k = k + 1
                    End If
                
                Next j
            Next i
            wsData.Columns(3).InsertIndent 1
        End If
    
    End With
    'wsData.Range(Cells(4, 15), Cells(k, 15)).Value = wsData.Range(Cells(4, 3), Cells(k, 3)).Value
    wsData.Cells(4, 16).Value = 1
    wsData.Cells(5, 16).Value = 2
    Dim selection1 As Range, selection2 As Range
    Set selection1 = wsData.Range(Cells(4, 16), Cells(5, 16))
    Set selection2 = wsData.Range(Cells(4, 16), Cells(k - 1, 16))
    selection1.AutoFill Destination:=selection2
    Columns("B:B").WrapText = False
    Columns("O:O").WrapText = False
    Columns("Q:Q").WrapText = False
    
    Columns("E:I").EntireColumn.Hidden = True
    Columns("K:N").EntireColumn.Hidden = True
    Columns("R:T").EntireColumn.Hidden = True
    wsData.Range("D4").Select
    
        'frezee
    'Application.Goto Range("C4"), True
    With ActiveWindow
     .SplitColumn = 3
     .SplitRow = 3
     .FreezePanes = True
    End With


    'wsInfo.Activate
    'Cells(5, 8).Select
    'Cells(5, 8).Value = "=IFERROR(COUNTIFS(Data!$B$4:$B$" & UBound(aData, 1) + 2 & "," & Cells(5, 2) & ",Data!$D$4:$D$" & UBound(aData, 1) + 2 _
    '& ",""<>"")/" & Cells(5, 4) & ",""-"")"
    If WorksheetFunction.CountA(wsInfo.Range("G:G")) - 1 <> 0 Then
        Call coderName
    End If
'        "=COUNTIFS(Data!R4C2:R1953C2,Info!RC[-6],Data!R[-2]C[-4]:R[1947]C[-4],""<>"")"
'    wsInfo.Activate
    
'    Cells(5, 7).Select
'    Cells(5, 7).Value = "=IF(ISNUMBER(SEARCH(Cell(""contents"")," & nameFrame & "!C5)),MAX($H$4:H4)+1,0)"
    wsData.Activate
    
Call TurnOnStuf

End Sub


Sub templateframe()
    Dim nameFrame As String
    nameFrame = ActiveCell
    
    Sheets.Add After:=Sheets(Sheets.count)
    Sheets(ActiveSheet.Name).Name = nameFrame
    
    'Sheets.Add after:=Sheets(Sheets.Count)
    'Sheets(ActiveSheet.Name).Name = "Frame"
    Set wb = ActiveWorkbook
    Set wsInfo = wb.Sheets("Info")
    aInfo = wsInfo.Range("A4").CurrentRegion
        
    
    Dim nmQuest As String
    For i = 3 To UBound(aInfo, 1)
        If aInfo(i, 8) = nameFrame Then
            nmQuest = nmQuest & ", " & aInfo(i, 2)
        End If
    Next i

    Cells(3, 1).Value = "Quest: " & Right(nmQuest, Len(nmQuest) - 2)
      
    Cells(4, 1).Value = "CoderID"
    Cells(4, 2).Value = "ClientID"
    Cells(4, 3).Value = "Statement (Bahasa)"
    Cells(4, 4).Value = "Statement (English)"
    Cells(4, 5).Value = "Note"
    Cells(4, 6).Value = "Information"
    Cells(4, 7).Value = "Flag"
    Cells(4, 8).Value = "Index"
    Cells(4, 9).Value = "Count"
    
    Range("A4:I4").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With

    Columns("A:A").ColumnWidth = 7
    Columns("A:A").NumberFormat = "@"
    Columns("B:B").ColumnWidth = 7
    Columns("B:B").NumberFormat = "@"
    Columns("C:C").ColumnWidth = 60
    Columns("C:C").WrapText = True
    Columns("D:D").ColumnWidth = 18
    Columns("D:D").WrapText = True
    Columns("E:E").ColumnWidth = 10
    Columns("F:F").ColumnWidth = 25
    Columns("G:G").ColumnWidth = 2
    Columns("H:H").ColumnWidth = 5
    Columns("I:I").ColumnWidth = 6
    'Cells(5, 8).AddComment ("Copy Rumus di cell ini sampai paling bawah")
    'Cells(5, 8).Comment.Visible = True

    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 6.75, 5.25, 40.25, 18).Select
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset24
    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Back"
    Selection.Placement = xlFreeFloating
    Selection.OnAction = "gotoDATA"
    
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 58.75, 5.25, 60.25, 18).Select
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset24
    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Menu..."
    Selection.Placement = xlFreeFloating
    Selection.OnAction = "dataMenuFrame"
    

    'frezee
    'Application.Goto Range("C4"), True
    Cells(5, 4).Select
    With ActiveWindow
     .SplitColumn = 2
     .SplitRow = 4
     .FreezePanes = True
    End With

End Sub



Sub FormatingData()
    
    Sheets.Add After:=Sheets(Sheets.count)
    Sheets(ActiveSheet.Name).Name = "Data"
        
    Cells(3, 1).Value = "Serial"
    Cells(3, 2).Value = "Quest"
    Cells(3, 3).Value = "Verbatim"
    Cells(3, 4).Value = "Coding"
    Cells(3, 5).Value = "Search"
    Cells(3, 6).Value = "Code"
    Cells(3, 7).Value = "CodeList"
    Cells(3, 8).Value = "Transfer Code"
    Cells(3, 9).Value = "Verification"
    Cells(3, 10).Value = "Coder"
    Cells(3, 11).Value = "Verificator"
    Cells(3, 12).Value = "Information"
    Cells(3, 13).Value = "ID INTV"
    Cells(3, 14).Value = "City"
    Cells(3, 15).Value = "Note"
    Cells(3, 16).Value = "Index"
    Cells(3, 17).Value = "Indentity"
    Cells(3, 19).Value = "Last Use Code"
    Cells(3, 20).Value = "Flaging"
    
    Range("A3:T3").Select
    Call FormatHeader
    
    
    Columns("A:A").ColumnWidth = 4
    Columns("A:A").NumberFormat = "@"
    Columns("B:B").ColumnWidth = 4
    Columns("B:B").NumberFormat = "@"
    Columns("C:C").ColumnWidth = 60
    Columns("C:C").WrapText = True
    Columns("D:D").ColumnWidth = 20
    Columns("D:D").WrapText = True
    Columns("D:D").NumberFormat = "@"
    Columns("E:E").ColumnWidth = 14
    Columns("F:F").ColumnWidth = 5
    Columns("G:G").ColumnWidth = 50
    Columns("G:G").WrapText = True
    Columns("H:H").ColumnWidth = 17.57
    Columns("H:H").NumberFormat = "@"
    Columns("I:I").ColumnWidth = 60
    Columns("J:J").ColumnWidth = 6
    Columns("K:K").ColumnWidth = 8
    Columns("L:L").ColumnWidth = 30
    Columns("O:O").ColumnWidth = 20
    Columns("P:P").ColumnWidth = 6
    Columns("Q:Q").ColumnWidth = 14
    'Makesure Lascode(RepeatLastAction) on Text Format
    Cells(4, 19).NumberFormat = "@"
         
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 6.75, 5.25, 58.25, 18).Select
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset24
    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "To Frame"
    Selection.Placement = xlFreeFloating
    Selection.OnAction = "gotoInfo"
    
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 70.75, 5.25, 60.25, 18).Select
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset24
    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Menu..."
    'CommandButton1.Accelerator = "A"
    Selection.Placement = xlFreeFloating
    'Selection.OnAction = "dataMenuFrame"
    Selection.OnAction = "dataMenu"
    'Application.MacroOptions Macro:="dataMenu", Description:="", ShortcutKey:="r"
    'Selection.ShapeRange.ShapeStyle = msoShapeStylePreset64
    'Selection.ShapeRange.BackgroundStyle = msoBackgroundStylePreset2
    'With Selection.ShapeRange.Line
    '    .Visible = msoTrue
    '    .ForeColor.RGB = RGB(0, 0, 0)
    '    .Transparency = 0
    'End With
    'With Selection.ShapeRange.TextFrame2.TextRange.Font.Fill
    '    .Visible = msoTrue
    '    .ForeColor.RGB = RGB(0, 0, 0)
    '    .Transparency = 0
    '    .Solid
    'End With
             
             

End Sub
Sub gotoDATA()
    'Error handle
    If WorksheetExists2("Data") = False Then
        MsgBox "Sheet Data doesn't exist"
        Exit Sub
    Else
        Worksheets("Data").Activate
        Cells(4, 3).Select
    End If

End Sub
Sub gotoInfo()
'    If WorksheetExists2("Data") = False Then
'        MsgBox "Sheet Data doesn't exist"
'        Exit Sub
'    Else
'        Worksheets("Info").Activate
'        Cells(5, 3).Select
'    End If
    Set wb = ActiveWorkbook
    Set wsInfo = wb.Sheets("Info")

    Dim nameFrame As String, findFrame As String
    Dim rFound As Range
    
    findFrame = ActiveCell.Offset(0, -2)
    
    On Error Resume Next
    Set rFound = wsInfo.Range("B:B").Cells.Find(What:=findFrame, LookAt:=xlPart, LookIn:=xlValues, _
                SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    On Error GoTo 0

    nameFrame = rFound.Offset(0, 6)
    If WorksheetExists2(nameFrame) Then
        Sheets(nameFrame).Activate
    Else
        MsgBox "Frame " & nameFrame & " doesn't Exist/Blank, Please Create Frame..."
    End If
    




End Sub

