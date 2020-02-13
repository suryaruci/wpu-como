Attribute VB_Name = "Info"

Sub dataMenuInfo()
    Set The_Menu_info = CreateSubMenuInfo
    The_Menu_info.ShowPopup
End Sub
Public Function CreateSubMenuInfo() As CommandBar
    Call TurnOffStuf

    Set wb = ActiveWorkbook
    Set wsInfo = wb.Sheets("Info")
    
    aInfo = wsInfo.Range("A4").CurrentRegion
    
    Dim allFrame As String
    Dim i, j As Integer
    
    Const pop_up_menu_name_Info = "Pop-up Menu Info"
    
    Dim the_command_bar_info As CommandBar
    Dim the_command_bar_control_info As CommandBarControl
    Dim CodingQuest_info As CommandBarPopup
    Dim nmQuest_info As String

    'Deletes any CommandBars that may be present
    For Each menu_item In CommandBars
        If menu_item.Name = pop_up_menu_name_Info Then
          CommandBars(pop_up_menu_name_Info).Delete
        End If
    Next

    ''Add our popup menu to the CommandBars collection
    Set the_command_bar_info = CommandBars.Add(Name:=pop_up_menu_name_Info, Position:=msoBarPopup, MenuBar:=False, Temporary:=False)
        
        '*****Menu Options*****
    Set the_command_bar_control_info = the_command_bar_info.Controls.Add
        the_command_bar_control_info.Caption = "Import verbatim"
        the_command_bar_control_info.OnAction = "testmacro"
    
    Set the_command_bar_control_info = the_command_bar_info.Controls.Add
        the_command_bar_control_info.Caption = "Questions info"
        the_command_bar_control_info.OnAction = "QuestInfo"
    
    Set TranspVerbatim = the_command_bar_info.Controls.Add(Type:=msoControlPopup)
        With TranspVerbatim
            .Caption = "Transpose verbatim"
            With .Controls.Add
                .Caption = "All verbatim"
                .OnAction = "Transposeverbatim"
                .Parameter = "All verbatim"
            End With
            For i = 3 To UBound(aInfo, 1)
                If aInfo(i, 7) <> "" And InStr(allFrame, " " & aInfo(i, 7) & " ") = 0 Then
                    With .Controls.Add
                        .Caption = aInfo(i, 7)
                        .OnAction = "testmacro"
                        .Parameter = aInfo(i, 7)
                    End With
                End If
                allFrame = " " & allFrame & " " & aInfo(i, 7) & " "
            Next i
        End With
        allFrame = ""
    
    'Set the_command_bar_control_info = the_command_bar_info.Controls.Add
    '    the_command_bar_control_info.Caption = "Coder name"
    '    the_command_bar_control_info.BeginGroup = True
    '    the_command_bar_control_info.OnAction = "Codername"
    
    'Set the_command_bar_control_info = the_command_bar_info.Controls.Add
    '    the_command_bar_control_info.Caption = "Create Frame"
    '    the_command_bar_control_info.OnAction = "TestMacro"
    
    Set the_command_bar_control_info = the_command_bar_info.Controls.Add
        the_command_bar_control_info.Caption = "Productivity"
        the_command_bar_control_info.BeginGroup = True
        the_command_bar_control_info.OnAction = "productivity"
    
    Set the_command_bar_control_info = the_command_bar_info.Controls.Add
        the_command_bar_control_info.Caption = "Back to Field"
        the_command_bar_control_info.OnAction = "BackToField"
        
    Set the_command_bar_control_info = the_command_bar_info.Controls.Add
        the_command_bar_control_info.Caption = "Verification summary"
        the_command_bar_control_info.OnAction = "TestMacro"
        
        'To add more items to the menum simply copy the 3 lines above and paste below
        'All you need to do is change the caption and onaction macro names.
          
    Set CreateSubMenuInfo = the_command_bar_info
    Call TurnOnStuf

End Function

Sub QuestInfo()
    Set wb = ActiveWorkbook
    Set wsInfo = wb.Sheets("Info")
    Set wsVerb = wb.Sheets("Verbatim")
    aVerb = wsVerb.UsedRange
    
    Dim i, j, k As Integer
    Dim Nperc(), TPerc, Complx As Long
    
    Call TurnOffStuf
    
        
        With Range("A4").CurrentRegion
            On Error Resume Next
            .Offset(2).ClearContents
        End With
    '    wsInfo.Rows("5:" & Rows.Count).ClearContents
            
        'Get name file
        Cells(3, 3).Value = CreateObject("Scripting.FileSystemObject").getbaseName(ThisWorkbook.Name)
        
        
        ReDim Nperc(0 To UBound(aVerb, 2))
        For i = 2 To UBound(aVerb, 2)
            Complx = 0
            k = 0
            wsInfo.Cells(i + 3, 1) = i - 1
            wsInfo.Cells(i + 3, 2) = aVerb(1, i)
            For j = 2 To UBound(aVerb, 1)
                If aVerb(j, i) <> "" Then
                    k = k + 1
                    'N-Verb
                    wsInfo.Cells(i + 3, 4) = k
                End If
                Nperc(i) = Nperc(i) + Len(Trim(aVerb(j, i)))
            Next j
            Complx = Nperc(i) / k
            If k = 0 Then
                wsInfo.Cells(i + 3, 4) = 0
            End If
            
            'Complexity
            Select Case Complx
                Case 1 To 30
                    wsInfo.Cells(i + 3, 6) = "Very Easy"
                Case 31 To 70
                    wsInfo.Cells(i + 3, 6) = "Easy"
                Case 71 To 100
                    wsInfo.Cells(i + 3, 6) = "Middle"
                Case Is > 100
                    wsInfo.Cells(i + 3, 6) = "Hight"
            End Select
            wsInfo.Range(Cells(i + 3, 6), Cells(i + 3, 6)).HorizontalAlignment = xlCenter
                    
            TPerc = TPerc + Nperc(i)
        Next i
        
        For i = 2 To UBound(aVerb, 2)
            If (Nperc(i) / TPerc) * 100 > 0 Then
                wsInfo.Cells(i + 3, 5) = (Nperc(i) / TPerc) * 100
                wsInfo.Range(Cells(i + 3, 5), Cells(i + 3, 5)).NumberFormat = "0.00"
            Else
                wsInfo.Cells(i + 3, 5) = "-"
                wsInfo.Range(Cells(i + 3, 5), Cells(i + 3, 5)).HorizontalAlignment = xlRight
            End If
        Next i
    'Boders
    wsInfo.UsedRange.Borders.LineStyle = xlNone
    wsInfo.Range("A4").CurrentRegion.Offset(1).Resize(Range("A4").CurrentRegion.Rows.count - 1).Borders.LineStyle = xlContinuous
    
    Range("H4:H" & Range("A4").CurrentRegion.Rows.count + 2).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -16777024
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -16777024
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -16777024
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -16777024
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    Call TurnOnStuf

End Sub

Sub ImportFile()
'https://www.youtube.com/watch?v=0YNhxVu2a5s&t=434s
'https://www.youtube.com/watch?v=h_sC6Uwtwxk
    Dim myFile As String
    Dim i, j As Integer
    Dim wbFile As Workbook
    Dim wsRead As Worksheet
    Dim wsWrite As Worksheet
    Dim wsTarget As Range, aHeader As Range
    
    Set wb = ActiveWorkbook
    Set wsInfo = wb.Sheets("Info")
        
    Application.ScreenUpdating = False
       
    Cells(3, 14).Value = "Source file"
    Cells(3, 14).Font.Size = 22
    Cells(4, 14).Value = "No"
    Cells(4, 15).Value = "N-Quest"
    Cells(4, 16).Value = "N-Data"
    Cells(4, 17).Value = "Source data"
    'Cells(4, 12).HorizontalAlignment = xlCenter
    
    Columns("M:M").ColumnWidth = 2
    Columns("N:N").ColumnWidth = 4
    Columns("O:O").ColumnWidth = 8
    Columns("P:P").ColumnWidth = 8
    Columns("Q:Q").ColumnWidth = 80
    
    Range("N4:Q4").Select
    Call FormatHeader
       
    'Cek exiting sheet
    If WorksheetExists2("Verbatim_") = False Then
        Sheets.Add After:=Sheets(Sheets.count)
        Sheets(ActiveSheet.Name).Name = "Verbatim_"
        Set wsWrite = wb.Sheets("Verbatim_")
        Set wsTarget = wsWrite.Range("A3")
    Else
        Set wsWrite = wb.Sheets("Verbatim_")
        aWrite = wsWrite.Range("B4").CurrentRegion
        Set wsTarget = wsWrite.Cells(UBound(aWrite, 1) + 3, 1)
    End If
    
    'Set aHeader = wsWrite.Range("A1:ZZ1")
    aInfo = wsInfo.Range("J4").CurrentRegion
    myFile = Application.GetOpenFilename(Title:="Browse Panter file..", FileFilter:="Excel Files, *.xl*")
    wsInfo.Cells(UBound(aInfo, 1) + 5, 13) = UBound(aInfo, 1)
    wsInfo.Cells(UBound(aInfo, 1) + 5, 16).Value = myFile
    
    Set sFile = Workbooks.Open(myFile)
    Set wsRead = sFile.Worksheets("Verbatim")
    
    ' Remove the any existing filters
    If wsRead.FilterMode = True Then
        wsRead.ShowAllData
    End If
    
    ' Remove the freeze panes
    ActiveWindow.FreezePanes = False
    
    ' Get the source data range
    Dim rgData As Range, rgData2 As Range
    Dim rowEnd As Long, colEnd As Long
    
    Set rgData = wsRead.Range("B3").CurrentRegion
    aData = wsRead.Range("B3").CurrentRegion
    
    wsInfo.Cells(UBound(aInfo, 1) + 4, 11) = UBound(aData, 2) - 1
    wsInfo.Cells(UBound(aInfo, 1) + 4, 12) = UBound(aData, 1) - 3
    
    ' IMPORTANT: Do not have any blank rows in the criteria range
    Dim rgCriteria As Range
    Set rgCriteria = wsRead.Range(Cells(1, UBound(aData, 2) + 2), Cells(2, UBound(aData, 2) + 2))
    rgCriteria(1, 1).Value = Trim("Status")
    'rgCriteria(1, 2).Value = wsRead.Range("B2").Value
    rgCriteria(2, 1).Value = "<>DO"
    'rgCriteria(2, 2).Value = "<>"""""
    
    rowEnd = wsRead.Range("B3").CurrentRegion.Rows.count
    colEnd = wsRead.Range("B3").CurrentRegion.Columns.count
    
    'wsWrite.Range(Cells(3, 1), Cells(3, UBound(aData, 2))).Value = rgData.Range(Cells(2, 1), Cells(2, UBound(aData, 2))).Value
    rgData.AdvancedFilter xlFilterCopy, rgCriteria, wsTarget
    'rgData.AdvancedFilter xlFilterCopy, rgCriteria, wsTarget
    sFile.Close False
    
    Set wsVerb = wb.Sheets("Verbatim_")
    Call removeBlankFromula(wsVerb)
    
    'https://www.youtube.com/watch?v=JPbwrak4hi4
    'Remove duplicate data
    Set rgData2 = wsVerb.Range("B3").CurrentRegion
    
    On Error Resume Next
    rgData2.RemoveDuplicates Columns:=Array(3, 2), Header:=xlNo
    On Error GoTo 0
    
    Application.ScreenUpdating = True

End Sub

Sub productivity()
    Set wb = ActiveWorkbook
    Set wsInfo = wb.Sheets("Info")

    
    Call TurnOffStuf
    Cells(3, 11).Value = "Productivity"
    Cells(3, 11).Font.Size = 22
    Cells(4, 11).Value = "Coder"
    Cells(4, 12).Value = "%"
    Cells(4, 12).HorizontalAlignment = xlCenter
    
    Columns("J:J").ColumnWidth = 2
    Columns("K:K").ColumnWidth = 17
    Columns("L:L").ColumnWidth = 8
    
    Range("K4:L4").Select
    Call FormatHeader

    If WorksheetExists2("Data") Then
        
        Set wsData = wb.Sheets("Data")
        aData = wsData.UsedRange
        
        With wsInfo.Range("K4").CurrentRegion
            On Error Resume Next
            .Offset(2).ClearContents
        End With
        
        Dim rngCoder As Range
        Dim tgtCoder As Range
        Set rngCoder = wsData.Range("J3:J" & UBound(aData, 1) + 2)
        Set tgtCoder = wsInfo.Range("K4")
        rngCoder.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=tgtCoder, Unique:=True
           
        If wsInfo.Range("K5") = "" Then
            wsInfo.Range("K5").Delete Shift:=xlUp
        End If
        
        aInfo = wsInfo.Range("K4").CurrentRegion
        
        Dim i As Integer, j As Long
        Dim Nperc(), TPerc As Long
    '    Nperc(0) = 0
        ReDim Nperc(0 To UBound(aInfo, 1))
        For i = 3 To UBound(aInfo, 1)
            For j = 2 To UBound(aData, 1)
                If Trim(LCase(aData(j, 10))) = Trim(LCase(aInfo(i, 1))) Then
                    Nperc(i) = Nperc(i) + Len(Trim(aData(j, 3)))
                End If
            Next j
            TPerc = TPerc + Nperc(i)
        Next i
        For i = 3 To UBound(aInfo, 1)
            wsInfo.Cells(i + 2, 12) = (Nperc(i) / TPerc) * 100
        Next i
        
        wsInfo.Range("K4:L100").Borders.LineStyle = xlNone
        wsInfo.Range("K4").CurrentRegion.Offset(0).Borders.LineStyle = xlContinuous
        wsInfo.Range("L5:L" & UBound(aInfo, 1) + 3).NumberFormat = "0.00"
    Else
        With wsInfo.Range("K4").CurrentRegion
            On Error Resume Next
            .Offset(1).ClearContents
        End With
        MsgBox "Sheet Data doesn't Exist/Blank, Please Transpose verbatim..."

    End If
    
    
    Call TurnOnStuf
    
End Sub
Sub GoToCreate()
'https://www.youtube.com/watch?v=Eh0tlrNUHaQ
    Dim vArr As Variant, i As Integer
    Dim oMenu As CommandBar, oItem As CommandBarControl
    Set oMenu = CommandBars.Add("", msoBarPopup, , True)
    
    vArrCap = Array("Go To Question", "Create Frame", "Go To Frame", "Refresh Coder Name")
    vArrAct = Array("GoToQuest", "CreateFrame", "GoToFrame", "RefreshData")
    For i = 0 To UBound(vArrCap)
        If i = 2 Then
            oItem.BeginGroup = True
        End If
        Set oItem = oMenu.Controls.Add
        oItem.Caption = vArrCap(i)
        oItem.OnAction = vArrAct(i)
    Next i
    oMenu.ShowPopup
End Sub
Sub refreshdata()
    Dim nameFrame As String
'    nameFrame = ActiveCell
    If WorksheetExists2("Data") = False Then
        MsgBox "Sheet Data doesn't create, Please transpose data..."
    Else
        Call coderName
    End If
End Sub
Sub coderName()
    Call TurnOffStuf
    
    
    Set wsInfo = wb.Sheets("Info")
    
    aInfo = wsInfo.Range("A3").CurrentRegion
    On Error Resume Next
    wsInfo.Range("A4:H" & UBound(aInfo, 1) + 2).Sort key1:=[A5], order1:=xlAscending, Header:=xlYes
    
    wsData.Activate
    Set wsData = wb.Sheets("Data")
    aData = wsData.UsedRange
    
    Call NormalSetting
    
    Dim nQuest() As Long
    ReDim nQuest(UBound(aInfo))
    
    k = 4
    For i = 0 To UBound(aInfo, 1)
        'On Error Resume Next
        nQuest(i) = WorksheetFunction.CountIf(wsData.Range(Cells(4, 2), Cells(UBound(aData, 1) + 2, 2)), aInfo(i + 3, 2))
        If nQuest(i) <> 0 Then
            wsData.Range(Cells(k, 10), Cells(nQuest(i) + k - 1, 10)).Value = aInfo(i + 3, 7)
            k = k + nQuest(i)
        End If
    Next i
    
    Call TurnOnStuf

End Sub
Sub createframe()
    Dim nameFrame As String
    nameFrame = ActiveCell
    If nameFrame = "" Then
        MsgBox "Empty frame..."
        Exit Sub
    End If
    If WorksheetExists2(nameFrame) Then
        MsgBox "Frame " & nameFrame & " does Exist, Please Go To Frame..."
    Else
        Call templateframe
    End If
End Sub
Sub gotoframe()
    Dim nameFrame As String
    nameFrame = ActiveCell
    If WorksheetExists2(nameFrame) Then
        Sheets(nameFrame).Activate
    Else
        MsgBox "Frame " & nameFrame & " doesn't Exist/Blank, Please Create Frame..."
    End If
    
End Sub
Sub gotoquest()
    Dim nameQuest As String
    Dim found As Range
    ActiveCell.Offset(0, -6).Activate
    nameQuest = ActiveCell
    With Sheets("DATA").Range("B:B")
        Set found = .Find(nameQuest, LookIn:=xlValues)
        If found Is Nothing Then
            MsgBox "Not found Question/Verbatim for " & nameQuest
            Application.DisplayAlerts = False
            Sheets("info").Activate
            Application.DisplayAlerts = True
            Exit Sub
        Else
            Sheets("Data").Activate
            found.Select
        End If
    End With
    
End Sub
Function TotalCoded()
    
    Set wb = ActiveWorkbook
    Set wsInfo = wb.Sheets("Info")
    
    aInfo = wsInfo.Range("A4").CurrentRegion
    If WorksheetExists2("Data") Then
        Set wsData = wb.Sheets("Data")
        aData = wsData.UsedRange
        For i = 5 To UBound(aInfo, 1) + 2
            If Cells(i, 4) <> 0 Then
                Cells(i, 9).Value = (WorksheetFunction.CountIfs(wsData.Range("B4:B" & UBound(aData, 1) + 2), Cells(i, 2).Value, _
                wsData.Range("D4:D" & UBound(aData, 1) + 2), "<>") / Cells(i, 4).Value) * 100
                Range(Cells(i, 9), Cells(i, 9)).NumberFormat = "00.00"
            Else
                Cells(i, 9).Value = "-"
                Range(Cells(i, 9), Cells(i, 9)).HorizontalAlignment = xlRight
            End If
        Next i
    Else
        wsInfo.Range("I5:I" & UBound(aInfo, 1) + 2).Value = ""
        'Exit Function
        'MsgBox "Sheet Data doesn't Exist/Blank, Please Transpose verbatim..."
    End If
        
End Function

Sub BackToField()
    
    Dim wsBF As Worksheet
    
    Set wsInfo = wb.Sheets("Info")
    Set wsData = wb.Sheets("Data")
    
    Call TurnOffStuf
    
'    'Error handle
'    If WorksheetExists2("UpdateCSV") = False Then
'        Sheets.Add After:=Sheets(Sheets.count)
'        Sheets(ActiveSheet.Name).Name = "UpdateCSV"
'        Set wsCSV = wb.Sheets("UpdateCSV")
'        wsVerb.Range("A1").CurrentRegion.Copy wsCSV.Range("A1")
'        wsCSV.UsedRange.Select
'        Selection.NumberFormat = "@"
'    Else
'        Set wsCSV = wb.Sheets("UpdateCSV")
'    End If'

'    acsv = wsCSV.Range("A1").CurrentRegion
    
    'Error handle
    If WorksheetExists2("BackToField") = False Then
    
        Sheets.Add After:=Sheets(Sheets.count)
        Sheets(ActiveSheet.Name).Name = "BackToField"
        
        Cells(2, 4).Value = "Project : "
        Cells(2, 5).Value = wsInfo.Cells(3, 3).Value
          
        Cells(3, 1).Value = "No"
        Cells(3, 2).Value = "ID INTV"
        Cells(3, 3).Value = "Serial"
        Cells(3, 4).Value = "Quest"
        Cells(3, 5).Value = "Verbatim"
        Cells(3, 6).Value = "Concern"
        Cells(3, 7).Value = "Note"
        Cells(3, 8).Value = "Confirm from Field"
        Cells(3, 9).Value = "Code"
        
        Range("A3:I3").Select
        Call FormatHeader
    
        Columns("A:A").ColumnWidth = 4
        Columns("B:B").ColumnWidth = 7
        Columns("C:C").ColumnWidth = 7
        Columns("D:D").ColumnWidth = 7
        Columns("E:E").ColumnWidth = 60
        Columns("E:E").WrapText = True
        Columns("F:F").ColumnWidth = 20
        Columns("F:F").WrapText = True
        Columns("G:G").ColumnWidth = 20
        Columns("H:H").ColumnWidth = 40
        Columns("I:I").ColumnWidth = 7
    
        ActiveSheet.Shapes.AddShape(msoShapeRectangle, 6.75, 5.25, 75.25, 18).Select
        Selection.ShapeRange.ShapeStyle = msoShapeStylePreset24
        Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Adding Code"
        Selection.Placement = xlFreeFloating
        Selection.OnAction = "testMacro"
        
        Set wsBF = wb.Sheets("BackToField")
'        aBF = wsBF.Range("A3").CurrentRegion
    
    Else
        Set wsBF = wb.Sheets("BackToField")
'        aBF = wsBF.Range("A3").CurrentRegion
        wsBF.Activate
    End If
    
    Call GetDataCallback
    
    Dim i As Long, j As Long
    Dim Containt As Boolean
    aBF = wsBF.Range("A3").CurrentRegion
    
    'Pick red word in sentence
    'https://www.mrexcel.com/board/threads/copy-only-the-text-of-a-certain-color-from-a-cell.753352/
    For i = 4 To UBound(aBF, 1) + 1
        Cells(i, 6).Value = ""
        For j = 1 To Len(Cells(i, 5).Value)
            If Cells(i, 5).Characters(Start:=j, Length:=1).Font.Color = vbRed Then
                Cells(i, 6).Value = Cells(i, 6).Value & Mid(Cells(i, 5), j, 1)
                Cells(i, 6).Font.Color = vbRed
                Containt = True
            End If
        Next j
    Next i
    If Containt = False Then
        MsgBox "No verbatim back to field"
        Sheets("BackToField").Delete
        wsInfo.Activate
        Exit Sub
    End If
    
    Range("E4").Select

    
    'Save file
    Dim nmFile As String, XnmFile As String
    Dim index As Integer
    nmFile = ThisWorkbook.path & "\BackToField " & _
    Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1)) & ".xlsx"
    ActiveSheet.Copy
    
    'Indexing file if exist file
    XnmFile = GetNextAvailableName(nmFile)
    ActiveWorkbook.SaveAs XnmFile, FileFormat:=51
    ActiveWorkbook.Close SaveChanges:=False
    
    Call TurnOnStuf
    
    MsgBox "Back to field File created :" & vbNewLine & XnmFile
   
End Sub

Sub GetDataCallback()
    
    Set wb = ActiveWorkbook
    
    Set wsData = wb.Sheets("Data")
    Set wsBF = wb.Sheets("BackToField")
    aBF = wsBF.Range("A3:I3")
    
    Dim rgData As Range, rgBF As Range, rgCriteria As Range
    ' Remove the any existing filters
    If wsData.FilterMode = True Then
        wsData.ShowAllData
    End If
    
    
    Set rgData = wsData.Range("A3").CurrentRegion
    Set rgCriteria = wsBF.Range(Cells(3, UBound(aBF, 2) + 2), Cells(4, UBound(aBF, 2) + 2))
    rgCriteria(1, 1).Value = Trim("Note")
    rgCriteria(2, 1).Value = "<>"
    'rgCriteria(2, 1).Value = "ada"
    Set rgBF = wsBF.Range("B3:E3")
    Set rgBF2 = wsBF.Range("G3")
    
    rgData.AdvancedFilter xlFilterCopy, rgCriteria, rgBF
    rgData.AdvancedFilter xlFilterCopy, rgCriteria, rgBF2
    
    rgCriteria(1, 1).Value = ""
    rgCriteria(2, 1).Value = ""
    
End Sub
Sub summary()

End Sub
