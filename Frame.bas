Attribute VB_Name = "Frame"
Sub dataMenuFrame()
    Set The_Menu_Frame = CreateSubMenuFrame
    The_Menu_Frame.ShowPopup
End Sub
Public Function CreateSubMenuFrame() As CommandBar
    Call TurnOffStuf

    Const pop_up_menu_name_Frame = "Pop-up Menu Frame"
    
    Dim the_command_bar_Frame As CommandBar
    Dim the_command_bar_control_Frame As CommandBarControl
    Dim CodingQuest_Frame As CommandBarPopup
    Dim nmQuest_Frame As String

    'Deletes any CommandBars that may be present
    For Each menu_item In CommandBars
        If menu_item.Name = pop_up_menu_name_Frame Then
          CommandBars(pop_up_menu_name_Frame).Delete
        End If
    Next

    ''Add our popup menu to the CommandBars collection
    Set the_command_bar_Frame = CommandBars.Add(Name:=pop_up_menu_name_Frame, Position:=msoBarPopup, MenuBar:=False, Temporary:=False)
        
        '*****Menu Options*****
    Set the_command_bar_control_Frame = the_command_bar_Frame.Controls.Add
        the_command_bar_control_Frame.Caption = "Frequency"
        the_command_bar_control_Frame.OnAction = "freq"
    
    Set the_command_bar_control_Frame = the_command_bar_Frame.Controls.Add
        the_command_bar_control_Frame.Caption = "Run report..."
        the_command_bar_control_Frame.OnAction = "errorFrame"
    
    Set the_command_bar_control_Frame = the_command_bar_Frame.Controls.Add
        the_command_bar_control_Frame.Caption = "Create Query"
        the_command_bar_control_Frame.OnAction = "testmacro"
        
    'Set the_command_bar_control_Frame = the_command_bar_Frame.Controls.Add
    '    the_command_bar_control_Frame.Caption = "Coder name"
    '    the_command_bar_control_Frame.BeginGroup = True
    '    the_command_bar_control_Frame.OnAction = "Codername"
    
    Set CreateSubMenuFrame = the_command_bar_Frame
    Call TurnOnStuf

End Function
Sub freq()
    Set wb = ActiveWorkbook
    Set wsData = wb.Sheets("Data")
    Set wsFrame = wb.ActiveSheet
    Dim nQuest As String
    Dim i As Integer
    
    Call TurnOffStuf
    
    nQuest = Mid(Cells(3, 1).Value, 8, Len(Cells(3, 1).Value))
    Cells(4, 10) = Trim(nQuest)
    Range("J4").TextToColumns DataType:=xlDelimited, ConsecutiveDelimiter:=True, comma:=True
    Range("J4").TextToColumns DataType:=xlDelimited, ConsecutiveDelimiter:=False, comma:=False
    
    Columns("A:B").NumberFormat = "@"
    
    aFrame = wsFrame.UsedRange
    aData = wsData.UsedRange
    
    Range(Cells(4, 10), Cells(4, UBound(aFrame, 2))).Select
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
    Selection.ColumnWidth = 6

    'Columns("A:A").ColumnWidth = 7

    
    For i = 10 To UBound(aFrame, 2)
        For j = 3 To UBound(aFrame, 1)
            If (Trim(aFrame(j, 1)) = "" And Trim(aFrame(j, 2)) = "") Or (IsNumeric(aFrame(j, 1)) = False Or IsNumeric(aFrame(j, 2)) = False) Then
                aFrame(j, i) = ""
            ElseIf WorksheetFunction.CountIfs(wsData.Range("B4:B" & UBound(aData, 1) + 2), Trim(aFrame(2, i)), wsData.Range("D4:D" & UBound(aData, 1) + 2), "*" & aFrame(j, 1) & aFrame(j, 2) & "*") = 0 Then
                    aFrame(j, i) = "-"
                Else
                    aFrame(j, i) = WorksheetFunction.CountIfs(wsData.Range("B4:B" & UBound(aData, 1) + 2), Trim(aFrame(2, i)), wsData.Range("D4:D" & UBound(aData, 1) + 2), "*" & aFrame(j, 1) & aFrame(j, 2) & "*")
            End If
        Next j
        'aFrame(j, i)
    Next i
    wsFrame.UsedRange = aFrame
    For i = 3 To UBound(aFrame, 1)
        If aFrame(i, 10) = "" Then
            aFrame(i, 9) = ""
        Else
            aFrame(i, 9) = WorksheetFunction.Sum(Range(Cells(i + 2, 10), Cells(i + 2, UBound(aFrame, 2))))
        End If
    Next i
    wsFrame.UsedRange = aFrame
    Cells(5, 4).Select
    Call TurnOnStuf

End Sub
Sub errorFrame()
    Set wb = ActiveWorkbook
    Set wsData = wb.Sheets("Data")
    Set wsFrame = wb.ActiveSheet
    aFrame = wsFrame.UsedRange
    
    Dim i, j As Integer
    
    'Cek Duplicate Statement or Code
    For i = 5 To UBound(aFrame, 1) + 2
        On Error Resume Next
        If WorksheetFunction.CountIf(wsFrame.Range(Cells(5, 3), Cells(UBound(aFrame, 1) + 2, 3)), wsFrame.Cells(i, 3).Value) > 1 Then
            MsgBox "Duplicate statement..." & Chr(10) & wsFrame.Cells(i, 3).Value & Chr(10) & "Position - " & wsFrame.Cells(i, 3).Address
            Exit Sub
        End If
        If WorksheetFunction.CountIf(wsFrame.Range(Cells(5, 1), Cells(UBound(aFrame, 1) + 2, 2)), wsFrame.Cells(i, 1).Value & wsFrame.Cells(i, 2).Value) > 1 Then
            MsgBox "Duplicate Code..." & Chr(10) & wsFrame.Cells(i, 2).Value & Chr(10) & "Position - " & wsFrame.Cells(i, 2).Address
            Exit Sub
        End If
    Next i
    
    If WorksheetFunction.Sum(Range(Cells(5, 9), Cells(UBound(aFrame, 1), 9))) = 0 Then
        MsgBox "Please run frequency!!!"
        Exit Sub
    End If
    
    j = 0
    For i = 3 To UBound(aFrame, 1)
        If aFrame(i, 9) = 0 And aFrame(i, 1) <> "" Then
            aFrame(i, 6) = ">>> Check code " & aFrame(i, 1) & ""
            j = j + 1
        End If
    Next i
    wsFrame.UsedRange = aFrame
    If j = 0 Then
        MsgBox "Clean..."
    Else
        MsgBox j & " - Error"
    End If

End Sub


Sub CreateQuery()
    
    Dim wsBF As Worksheet
    
    Set wsInfo = wb.Sheets("Info")
    Set wsData = wb.Sheets("Data")
    
    Call TurnOffStuf
    
    
    'Error handle
    If WorksheetExists2("Query") = False Then
    
        Sheets.Add After:=Sheets(Sheets.count)
        Sheets(ActiveSheet.Name).Name = "Query"
        
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
        
        Set wsBF = wb.Sheets("Query")
'        aBF = wsBF.Range("A3").CurrentRegion
    
    Else
        Set wsBF = wb.Sheets("Query")
'        aBF = wsBF.Range("A3").CurrentRegion
        wsBF.Activate
    End If
    
    Call GetQuery
    
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
        Sheets("Query").Delete
        Exit Sub
    End If
    
    Range("E4").Select

    
    'Save file
    Dim nmFile As String, XnmFile As String
    Dim index As Integer
    nmFile = ThisWorkbook.path & "\Query " & _
    Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1)) & ".xlsx"
    ActiveSheet.Copy
    
    'Indexing file if exist file
    XnmFile = GetNextAvailableName(nmFile)
    ActiveWorkbook.SaveAs XnmFile, FileFormat:=51
    ActiveWorkbook.Close SaveChanges:=False
    
    Call TurnOnStuf
    
    MsgBox "Back to field File created :" & vbNewLine & XnmFile
   
End Sub

Sub GetQuery()
    
    Set wb = ActiveWorkbook
    
    Set wsData = wb.Sheets("Data")
    Set wsBF = wb.Sheets("Query")
    aBF = wsBF.Range("A3:I3")
    
    Dim rgData As Range, rgBF As Range, rgCriteria As Range
    ' Remove the any existing filters
    If wsData.FilterMode = True Then
        wsData.ShowAllData
    End If
    
    
    Set rgData = wsData.Range("A3").CurrentRegion
    Set rgCriteria = wsBF.Range(Cells(3, UBound(aBF, 2) + 2), Cells(4, UBound(aBF, 2) + 2))
    rgCriteria(1, 1).Value = Trim("Note")
    rgCriteria(2, 1).Value = "Query"
    'rgCriteria(2, 1).Value = "ada"
    Set rgBF = wsBF.Range("B3:E3")
    Set rgBF2 = wsBF.Range("G3")
    
    rgData.AdvancedFilter xlFilterCopy, rgCriteria, rgBF
    rgData.AdvancedFilter xlFilterCopy, rgCriteria, rgBF2
    
    rgCriteria(1, 1).Value = ""
    rgCriteria(2, 1).Value = ""
    
End Sub



