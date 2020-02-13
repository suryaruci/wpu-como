Attribute VB_Name = "Data"
Public The_Menu As CommandBar

Sub dataMenu()
    Set The_Menu = CreateSubMenu
    The_Menu.ShowPopup
End Sub
'https://www.youtube.com/watch?v=z7rfqnueCeE
Public Function CreateSubMenu() As CommandBar
    Call TurnOffStuf

    Set wb = ActiveWorkbook
    Set wsInfo = wb.Sheets("Info")
    
    aInfo = wsInfo.Range("A4").CurrentRegion
    
    Dim allFrame As String
    Dim i, j As Integer
    
    Const pop_up_menu_name = "Pop-up Menu"
    
    Dim the_command_bar As CommandBar
    Dim the_command_bar_control As CommandBarControl
    Dim CodingQuest As CommandBarPopup
    Dim nmQuest As String

    'Deletes any CommandBars that may be present
    For Each menu_item In CommandBars
        If menu_item.Name = pop_up_menu_name Then
          CommandBars(pop_up_menu_name).Delete
        End If
    Next

    ''Add our popup menu to the CommandBars collection
    Set the_command_bar = CommandBars.Add(Name:=pop_up_menu_name, Position:=msoBarPopup, MenuBar:=False, Temporary:=False)
        
        '*****Menu Options*****
    vArrSet = Array("ECodingQuest", "SCodingQuest", "ACodingQuest", "checkCode", "VerifyQuest")
    vArrCap = Array("Exact coding", "Similarity coding", "Search...", "Run report...", "Verification", "Create Identity", "Export to CSV file", "Transfer and Export to CSV file")
    vArrAct = Array("exactcod", "SimilarityCod", "searchable", "runReport", "verification", "CreateIdentity", "UpdateToOpen", "transfer")
    For j = 0 To UBound(vArrCap) - 1
        If j < 5 Then
            Set vArrSet(j) = the_command_bar.Controls.Add(Type:=msoControlPopup)
                With vArrSet(j)
                    .Caption = vArrCap(j)
                    For i = 3 To UBound(aInfo, 1)
                        If aInfo(i, 8) <> "" And InStr(allFrame, " " & aInfo(i, 8) & " ") = 0 Then
                            With .Controls.Add
                                .Caption = aInfo(i, 8)
                                .OnAction = vArrAct(j)
                                '.OnAction = "ExactCod"
                                .Parameter = aInfo(i, 8)
                            End With
                        End If
                        allFrame = " " & allFrame & " " & aInfo(i, 8) & " "
                    Next i
                    If j = 2 Then
                        .BeginGroup = True
                    End If
                End With
                allFrame = ""
        Else
            Set the_command_bar_control = the_command_bar.Controls.Add
                the_command_bar_control.Caption = vArrCap(j)
                If j = 5 Or j = 6 Then
                    the_command_bar_control.BeginGroup = True
                End If
                the_command_bar_control.OnAction = vArrAct(j)
                'the_command_bar_control.BeginGroup = True
                
        End If
    Next j
        
    Set TransferAndExport = the_command_bar.Controls.Add(Type:=msoControlPopup)
        With TransferAndExport
            .Caption = "Transfer and Export to CSV file"
            With .Controls.Add
                .Caption = "All"
                .OnAction = "TransferAll"
                '.Parameter = "All"
            End With
            For i = 3 To UBound(aInfo, 1)
                If aInfo(i, 8) <> "" And InStr(allFrame, " " & aInfo(i, 8) & " ") = 0 Then
                    With .Controls.Add
                        .Caption = aInfo(i, 8)
                        .OnAction = "transferPerFrame"
                        '.OnAction = "ExactCod"
                        .Parameter = aInfo(i, 8)
                        If i = 3 Then
                            .BeginGroup = True
                        End If
                    End With
                End If
                allFrame = " " & allFrame & " " & aInfo(i, 8) & " "
            Next i
        End With
        allFrame = ""
        
        'To add more items to the menum simply copy the 3 lines above and paste below
        'All you need to do is change the caption and onaction macro names.
          
    Set CreateSubMenu = the_command_bar
    Call TurnOnStuf

End Function

Function listFrame(k, aStart As Variant, aEnd As Variant, nmFrame)
'https://www.tek-tips.com/viewthread.cfm?qid=1739131
    Call TurnOffStuf
    With CommandBars.ActionControl

        Dim nCoding As Long
        Dim sCoding As String, bCoding As String
        Dim nFrame As Long
        Dim sFrame As String
        Dim bFrame() As String
        
        Dim aFrm() As String
        Dim aCod() As String
        Dim nFrm, nCod, i As Integer
        Dim bFrm As Boolean, bCod As Boolean
        
        Set wb = ActiveWorkbook
        Set wsInfo = wb.Sheets("Info")
        Set wsData = wb.Sheets("Data")
        
        'check Available Frame
        If WorksheetExists2(.Parameter) = False Then
            MsgBox "Frame " & .Parameter & " doesn't Exist/Blank, Please Create Frame..."
            Exit Function
        End If

        Set wsFrame = wb.Sheets(.Parameter)
        
        aData = wsData.UsedRange
        aInfo = wsInfo.Range("A4").CurrentRegion
        
        nmFrame = .Parameter
        
        j = -1
        Dim nmQuest() As String
        For i = 2 To UBound(aInfo, 1)
            If aInfo(i, 8) = .Parameter Then
                j = j + 1
                ReDim Preserve nmQuest(j)
                nmQuest(j) = aInfo(i, 2)
            End If
        Next i
        
        k = UBound(nmQuest, 1)
        j = 2
        'Dim verbStart() As String, verbEnd() As String, codStart() As String, codEnd() As String
        ReDim aStart(k), aEnd(k)
        ', codStart(k), codEnd(k), frmStart(k), frmEnd(k), sourceVerb(k), sourceCod(k)
        
        
        For i = 0 To k
            For j = j To UBound(aData, 1)
                If aData(j, 2) = nmQuest(i) Then
                    'verbStart(i) = wsData.Cells(j, 3).Address(RowAbsolute:=False)
                    'codStart(i) = wsData.Cells(j, 4).Address(RowAbsolute:=False)
                    aStart(i) = j
                    aEnd(i) = WorksheetFunction.CountIf(wsData.Range("B3:B" & UBound(aData, 1)), nmQuest(i))
                    Exit For
                End If
            Next j
            j = aEnd(i) + 1
            'verbEnd(i) = wsData.Cells(frmStart(i) + frmEnd(i) - 1, 3).Address(RowAbsolute:=False)
            'codEnd(i) = wsData.Cells(frmStart(i) + frmEnd(i) - 1, 4).Address(RowAbsolute:=False)
            'sourceVerb(i) = wsData.Range(verbStart(i), verbEnd(i))
            'sourceCod(i) = wsData.Range(codStart(i), codEnd(i))
            aFrame = wsFrame.UsedRange
        Next i
    Call TurnOnStuf
    End With
    
End Function
Sub ExactCod()
    Dim bFrame() As String
    k = 0
    Call listFrame(k, aStart, aEnd, nmFrame)
    
    Call TurnOffStuf
    
    Call NormalSetting
    
    m = 0
    For i = 0 To k
        sCoding = wsData.UsedRange
        For nCoding = aStart(i) To aStart(i) + aEnd(i) - 1
            'menghilangkan char yang tidak diinginkan
            sCoding(nCoding, 3) = Trim(LCase(sCoding(nCoding, 3)))
            'sCoding = Trim(LCase(sourceVerb(i)(nCoding, 1)))
            aCod = Split(sCoding(nCoding, 3), ",")
            'sourceVerb(i)(nCoding, 1) = ""
            For nCod = LBound(aCod) To UBound(aCod)
                'trmCoding = RemovePunctuation(aCod(nCod))
                bCoding = ""
                For nFrame = 2 To UBound(aFrame, 1)
                    sFrame = Trim(LCase(aFrame(nFrame, 3)))
                    aFrm = Split(sFrame, "/")
                    bFrm = False
                    For nFrm = LBound(aFrm) To UBound(aFrm)
                        'If Similarity(RemovePunctuation(Trim(LCase(aCod(nCod)))), Trim(aFrm(nFrm)), , 1) >= 0.9 Then
                        If Trim(LCase(aCod(nCod))) = Trim(aFrm(nFrm)) Then
                            'wsData.Cells(nCoding, 3) = aFrame(nFrame, 2)
                            bCoding = aFrame(nFrame, 2)
                            bFrm = True
                            Exit For
                        End If
                    Next nFrm
                    If bFrm Then Exit For
                Next nFrame
                If sCoding(nCoding, 4) = "" Then
                    sCoding(nCoding, 4) = bCoding
                Else
                    sCoding(nCoding, 4) = Trim(sCoding(nCoding, 4) & " " & bCoding)
                End If
                If bCoding = "" And aCod(nCod) <> "" Then
                    m = m + 1
                    ReDim Preserve bFrame(m)
                    bFrame(m) = aCod(nCod)
                End If
            Next nCod
        Next nCoding
        
        wsData.Range("A4").CurrentRegion = sCoding
        'wsData.Cells(3, 1).Select
        'Range(Cells(frmStart(i), 4), Cells(frmStart(i) + frmEnd(i) - 1, 4)).Resize(frmEnd(i), 1) = sourceVerb(i)
    Next i
            
    If MsgBox("Do you want to create query?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    If m > 0 Then
        Call lstQuery(bFrame)
    End If
    Call TurnOnStuf

End Sub
Sub lstQuery(bFrame)
            
    Call TurnOffStuf
    
    wsFrame.Activate
    aFrame = wsFrame.UsedRange
    For m = 1 To UBound(bFrame, 1)
        wsFrame.Cells(UBound(aFrame, 1) + 5 + m, 3).Value = bFrame(m)
    Next
    On Error Resume Next
    Dim Querydata As Range
    Set Querydata = wsFrame.Range(Cells(UBound(aFrame, 1) + 5, 3), Cells(UBound(aFrame, 1) + 5 + m, 3))
    Querydata.RemoveDuplicates Columns:=Array(1), Header:=xlNo
    Querydata.SpecialCells(xlCellTypeBlanks).Delete xlShiftUp
    On Error GoTo 0
    
    Call TurnOnStuf
       
    'Call removeBlankFromula(wsFrame)
End Sub
Function SimilarityCod()
'https://www.tek-tips.com/viewthread.cfm?qid=1739131
    Dim bFrame() As String
    
    Call TurnOffStuf
    k = 0
    Call listFrame(k, sourceVerb, sourceCod, frmStart, frmEnd)
    m = 0
    
    Call NormalSetting
    
    For i = 1 To k
        For nCoding = 1 To UBound(sourceVerb(i), 1)
            sCoding = Trim(LCase(sourceVerb(i)(nCoding, 1)))
            aCod = Split(sCoding, ",")
            'bCod = False
            sourceVerb(i)(nCoding, 1) = ""
            For nCod = LBound(aCod) To UBound(aCod)
                'trmCoding = RemovePunctuation(aCod(nCod))
                bCoding = ""
                For nFrame = 2 To UBound(aFrame, 1)
                    sFrame = Trim(LCase(aFrame(nFrame, 3)))
                    aFrm = Split(sFrame, "/")
                    bFrm = False
                    For nFrm = LBound(aFrm) To UBound(aFrm)
                        If FuzzyMatchByWord(RemovePunctuation(Trim(LCase(aCod(nCod)))), Trim(aFrm(nFrm)), False) >= 50 Then
                        'If Trim(LCase(aCod(nCod))) = Trim(aFrm(nFrm)) Then
                            'wsData.Cells(nCoding, 3) = aFrame(nFrame, 2)
                            bCoding = aFrame(nFrame, 2)
                            bFrm = True
                            'Exit For
                        End If
                    Next nFrm
                    'If bFrm Then Exit For
                Next nFrame
                If sourceVerb(i)(nCoding, 1) = "" Then
                    sourceVerb(i)(nCoding, 1) = bCoding
                Else
                    sourceVerb(i)(nCoding, 1) = Trim(sourceVerb(i)(nCoding, 1) & " " & bCoding)
                End If
                If bCoding = "" And aCod(nCod) <> "" Then
                    m = m + 1
                    ReDim Preserve bFrame(m)
                    bFrame(m) = aCod(nCod)
                End If
            Next nCod
        Next nCoding
        
        wsData.Cells(3, 1).Select
        Range(Cells(frmStart(i), 4), Cells(frmStart(i) + frmEnd(i) - 1, 4)).Resize(frmEnd(i), 1) = sourceVerb(i)
    Next i
            
    If m > 0 Then
        Call lstQuery(bFrame)
    End If

    Call TurnOnStuf

End Function
Sub Searchable()
    
    
    Call listFrame(k, aStart, aEnd, nmFrame)

    'check Available Frame
    If nmFrame = "" Then
        Exit Sub
    End If
    
    Call TurnOffStuf

    Set wb = ActiveWorkbook
    Set wsData = wb.Sheets("Data")
    Set wsInfo = wb.Sheets("Info")
    aData = wsData.UsedRange
    aInfo = wsInfo.Range("A4").CurrentRegion
    
    'reset Last used code
    wsData.Range("S4").Value = ""
    
    Dim Ncodelist As Integer
    
    'Flaging
    On Error Resume Next
    Sheets(nmFrame).Activate
    On Error GoTo 0
    
    'Cells(5, 7).Select
    Cells(5, 7).Value = "=IF(ISNUMBER(SEARCH(Cell(""contents"")," & nmFrame & "!C5)),MAX($G$4:G4)+1,0)"
    Cells(5, 7).Select
    Selection.Copy
    Range(Cells(6, 7), Cells(UBound(aFrame, 1) + 4, 7)).Select
    ActiveSheet.Paste
    Columns("G:H").EntireColumn.Hidden = True
    
    If WorksheetFunction.CountA(Range("C:C")) < 3 Then
        MsgBox "Frame is Blank"
        Exit Sub
    End If
    
    'Indexing
    Cells(5, 8).Value = 1
    Cells(6, 8).Value = 2
    Dim selection1 As Range, selection2 As Range
    Set selection1 = Range(Cells(5, 8), Cells(6, 8))
    Set selection2 = Range(Cells(5, 8), Cells(UBound(aFrame, 1) + 4, 8))
    selection1.AutoFill Destination:=selection2
    
    Sheets("Data").Activate
    
    Call NormalSetting
    
    Columns("E:G").EntireColumn.Hidden = False
    
    Cells(4, 20).Select
    
    'https://superuser.com/questions/986501/numbering-of-rows-in-a-filter
    'Automatic numbering available cell, non numbering if hide row
    Cells(4, 20).Value = "=IF(E4>0,1,IF(AGGREGATE(2,5,T$3:T3)>0,AGGREGATE(2,5,T$3:T3)+1,""""))"
    'Cells(4, 20).Value = "=IF(E4>0,1,IF(T3>0,T3+1,""""))"
    Selection.Copy
    Range(Cells(4, 20), Cells(UBound(aData, 1) + 15, 20)).Select
    ActiveSheet.Paste
        
    
    For i = 0 To k
        If aStart(i) <> "" Then
        
            Cells(aStart(i) + 2, 7).Offset(0, -1).Select
            Cells(aStart(i) + 2, 7).Offset(0, -1).Value = "=IFERROR(if(INDEX(" & nmFrame & "!A:A ,MATCH(" _
            & Cells(aStart(i) + 2, 7).Address(RowAbsolute:=False) & "," & nmFrame & "!C:C,0))<>0,INDEX(" & nmFrame & "!A:A ,MATCH(" _
            & Cells(aStart(i) + 2, 7).Address(RowAbsolute:=False) & "," & nmFrame & "!C:C,0)),INDEX(" & nmFrame & "!B:B ,MATCH(" _
            & Cells(aStart(i) + 2, 7).Address(RowAbsolute:=False) & "," & nmFrame & "!C:C,0))),"""")"
            Selection.Copy
            Range(Cells(aStart(i) + 2, 7).Offset(0, -1), Cells(aStart(i) + aEnd(i) + 1, 7).Offset(0, -1)).Select
            ActiveSheet.Paste
            
            Cells(aStart(i) + 2, 7).Select
            Cells(aStart(i) + 2, 7).Value = "=IFERROR(INDEX(" & nmFrame & "!C:C,MATCH(" & Cells(aStart(i) + 2, 7).Offset(0, 13).Address(RowAbsolute:=False) & "," & nmFrame & "!G:G,0)),"""")"
            Selection.Copy
            Range(Cells(aStart(i) + 2, 7), Cells(aStart(i) + aEnd(i) + 1, 7)).Select
            ActiveSheet.Paste
        End If
    Next i
        
    'ActiveWindow.NewWindow
    'Windows("Book1:1").Activate
    'Windows("Book1:2").Activate
    'Application.Left = 772
    'Application.Top = 1
    'Application.Width = 770.25
    'Application.Height = 546
    'Windows("Book1:1").Activate
    
    'Cek exiting Workbook
    On Error Resume Next
    Windows(ThisWorkbook.Name & ":2").Close

    'Display Side by side sheet
    ActiveWindow.NewWindow
    Windows(ActiveWorkbook.Name & ":2").Activate
    With Application
        .WindowState = xlNormal
        .Left = dllGetHorizontalResolution * 0.5
        .Top = 0
        .Width = dllGetHorizontalResolution * 0.25
        .Height = dllGetVerticalResolution
    End With
    Sheets(nmFrame).Select
    ActiveWindow.Zoom = 90
    Windows(ActiveWorkbook.Name & ":1").Activate
    With Application
        .WindowState = xlNormal
        .Left = 0
        .Top = 0
        .Width = dllGetHorizontalResolution * 0.5
        .Height = dllGetVerticalResolution
    End With
    Sheets("Data").Select
    ActiveWindow.Zoom = 90
    'Windows.Arrange ArrangeStyle:=xlVertical
    
    'Message to use shortcut add Code
    MsgBox "Type keyword in Search-Col E" & vbNewLine & vbNewLine & _
    "1. Adding code (Code-Col F to Coding-Col D) : Select code in column Code then Ctrl + M" & vbNewLine & vbNewLine & _
    "2. Repeat Last code : select cell/range (Coding-Col D) then Pres Ctrl + L" & vbNewLine & vbNewLine & _
    "3. Delete/undo Last code : select cell/range then Pres Ctrl + Shift + L"
    
    'Go to Question
    With Sheets("DATA").Range("B:B")
        nameQuest = Cells(aStart(0) + 2, 2).Value
        Set found = .Find(nameQuest, LookIn:=xlValues)
        Application.Goto found, True
    End With

    Call TurnOnStuf

End Sub
Sub runReport()
    
    On Error Resume Next
    If wsData.Range("L:L").EntireColumn.Hidden = True Then
        wsData.Range("L:L").EntireColumn.Hidden = False
    End If
    Call TurnOffStuf
    Call listFrame(k, aStart, aEnd, nmFrame)
    
    Call NormalSetting

    sCoding = wsData.UsedRange
    Dim Nerr As Long
    Nerr = 0
    For i = 0 To k
        For nCoding = aStart(i) To aStart(i) + aEnd(i) - 1
            'menghilangkan char yang tidak diinginkan
            sCoding(nCoding, 4) = Trim(sCoding(nCoding, 4))
            sCoding(nCoding, 4) = Replace(sCoding(nCoding, 4), ",", " ")
            sCoding(nCoding, 4) = Replace(sCoding(nCoding, 4), ".", " ")
            sCoding(nCoding, 4) = Replace(sCoding(nCoding, 4), ";", " ")
            Do While InStr(1, sCoding(nCoding, 4), "  ")
                sCoding(nCoding, 4) = Replace(sCoding(nCoding, 4), "  ", " ")
            Loop
            sCoding(nCoding, 12) = ""
            If sCoding(nCoding, 4) = "" Then
                'cek jika belum di coding, disimpan di column information
                sCoding(nCoding, 12) = ">>>Empty..."
                Nerr = Nerr + 1
                GoTo 10
            End If
            aCod = Split(sCoding(nCoding, 4), " ")
            'If UBound(aCod) > 0 And (UBound(aCod) * 10 < UBound(Split(sCoding(nCoding, 3), " "))) Then
            '    sCoding(nCoding, 12) = sCoding(nCoding, 12) & Chr(10) & ">>>Incomplite coding (less code)"
            'End If
            'If UBound(aCod) > 0 And (UBound(aCod) * 2 > UBound(Split(sCoding(nCoding, 3), " "))) Then
            '    sCoding(nCoding, 12) = sCoding(nCoding, 12) & Chr(10) & ">>>Incomplite coding (to much code)"
            'End If
            For nCod = LBound(aCod) To UBound(aCod)
                bFrm = False
                For nFrame = 2 To UBound(aFrame, 1)
                    If Trim(LCase(aFrame(nFrame, 1))) <> "" Then
                        sFrame = Trim(LCase(aFrame(nFrame, 1)))
                    Else
                        sFrame = Trim(LCase(aFrame(nFrame, 2)))
                    End If
                    If Trim(LCase(aCod(nCod))) = Trim(sFrame) Then
                        If sCoding(nCoding, 12) = "" Then
                            bFrm = True
                            Exit For
                        End If
                    End If
                Next nFrame
                If bFrm = False Then
                    sCoding(nCoding, 12) = sCoding(nCoding, 12) & Chr(10) & ">>>not allowed response " & Trim(LCase(aCod(nCod)))
                    Nerr = Nerr + 1
                End If
            Next nCod
10:
        Next nCoding
        
        wsData.Range("D:D").NumberFormat = "@"
        wsData.Range("A4").CurrentRegion = sCoding
        'sCoding = wsData.UsedRange
        
    
    Next i
    
    'report error
    If Nerr = 0 Then
        MsgBox "Clean..."
    Else
        MsgBox Nerr & "-Error"
    End If

    'Go to Question
    With Sheets("DATA").Range("B:B")
        nameQuest = Cells(aStart(0) + 2, 2).Value
        Set found = .Find(nameQuest, LookIn:=xlValues)
        Application.Goto found, True
        'found.Select
    End With

    Call TurnOnStuf

End Sub
Sub Verification()
    
    Call TurnOffStuf
    Call listFrame(k, aStart, aEnd, nmFrame)
    sCoding = wsData.UsedRange
    
    Call NormalSetting
    
    For i = 0 To k
        For nCoding = aStart(i) To aStart(i) + aEnd(i) - 1
            'menghilangkan char yang tidak diinginkan
            sCoding(nCoding, 4) = Trim(sCoding(nCoding, 4))
            sCoding(nCoding, 4) = Replace(sCoding(nCoding, 4), ",", " ")
            sCoding(nCoding, 4) = Replace(sCoding(nCoding, 4), ".", " ")
            sCoding(nCoding, 4) = Replace(sCoding(nCoding, 4), ";", " ")
            Do While InStr(1, sCoding(nCoding, 4), "  ")
                sCoding(nCoding, 4) = Replace(sCoding(nCoding, 4), "  ", " ")
            Loop
            If sCoding(nCoding, 4) = "" Then
                MsgBox "Please run report....data empty"
                Exit Sub
            End If
            aCod = Split(sCoding(nCoding, 4), " ")
            For nCod = LBound(aCod) To UBound(aCod)
                bFrm = False
'                sCoding(nCoding, 12) = ""
                For nFrame = 2 To UBound(aFrame, 1)
                    If aFrame(nFrame, 1) <> "" Then
                        sFrame = Trim(LCase(aFrame(nFrame, 1)))
                    Else
                        sFrame = Trim(LCase(aFrame(nFrame, 2)))
                    End If
                    If Trim(LCase(aCod(nCod))) = Trim(sFrame) Then
                        If sCoding(nCoding, 9) = "" Then
                            sCoding(nCoding, 9) = Trim(aFrame(nFrame, 2) & " | " & aFrame(nFrame, 3))
                            bFrm = True
                            Exit For
                        Else
                            sCoding(nCoding, 9) = Trim(sCoding(nCoding, 9) & Chr(10) & aFrame(nFrame, 2) & " | " & aFrame(nFrame, 3))
                            bFrm = True
                            Exit For
                        End If
                    End If
                Next nFrame
                If bFrm = False Then
                    MsgBox "Please run report....data not clean"
                    Exit Sub
                End If
            Next nCod
10:
        Next nCoding
    Next i
    
    If wsData.Range("I:I").EntireColumn.Hidden = True Then
        wsData.Range("I:I").EntireColumn.Hidden = False
    End If
            
    wsData.Range("A4").CurrentRegion = sCoding
    
    'Randomize
    For i = 0 To k
        'https://www.computergaga.com/blog/pick-names-at-random-from-a-list-excel-vba1/
        Dim HowMany As Integer
        Dim NoOfNames As Long
        Dim RandomNumber As Integer
        Dim Names() As String 'Array to store randomly selected names
        Dim iRand As Long
        Dim CellsOut As Long 'Variable to be used when entering names onto worksheet
        Dim ArI As Byte 'Variable to increment through array indexes
        
        'Randomize 10%
        'NoOfNames = 0
        NoOfNames = aEnd(i)
        HowMany = Application.WorksheetFunction.RoundUp(NoOfNames * 0.1, 0)
        
        ReDim Names(1 To HowMany) 'Set the array size to how many names required
        'NoOfNames = nCoding ' Find how many names in the list
        'NoOfNames = Application.CountA(Range("A:A")) - 1 ' Find how many names in the list
        iRand = 1
        
        Do While iRand <= HowMany
RandomNo:
            RandomNumber = Application.RandBetween(aStart(i) + 2, aStart(i) + 2 + NoOfNames)
            'Check to see if the name has already been picked
            For ArI = LBound(Names) To UBound(Names)
                If Names(ArI) = Cells(RandomNumber, 1).Value Then
                    GoTo RandomNo
                End If
            Next ArI
            Names(iRand) = Cells(RandomNumber, 1).Value ' Assign random name to the array
            iRand = iRand + 1
            Cells(RandomNumber, 9).Interior.ColorIndex = 37
        Loop
    Next i
    
    'Go to Question
    With Sheets("DATA").Range("B:B")
        nameQuest = Cells(aStart(0) + 2, 2).Value
        Set found = .Find(nameQuest, LookIn:=xlValues)
        Application.Goto found, True
    End With
    
    Call TurnOnStuf

End Sub

Sub UpdatetoOpen()

    Call TurnOffStuf

    Set wb = ActiveWorkbook
    Set wsVerb = wb.Sheets("Verbatim")
    Set wsData = wb.Sheets("Data")
    aData = wsData.UsedRange
    Dim trans As Boolean
    
    Call NormalSetting

    'Error handle
    If WorksheetExists2("UpdateCSV") = False Then
        Sheets.Add After:=Sheets(Sheets.count)
        Sheets(ActiveSheet.Name).Name = "UpdateCSV"
        Set wsCSV = wb.Sheets("UpdateCSV")
        wsVerb.Range("A1").CurrentRegion.Copy wsCSV.Range("A1")
        wsCSV.UsedRange.Select
        Selection.NumberFormat = "@"
        Selection.WrapText = False
    Else
        Set wsCSV = wb.Sheets("UpdateCSV")
    End If

    acsv = wsCSV.Range("A1").CurrentRegion

    sCoding = wsData.UsedRange
    
    k = 2
    
    Dim targetverb As String
    For i = 2 To UBound(acsv, 2)
        For j = 2 To UBound(acsv, 1)
            For k = k To UBound(aData, 1)
                If Application.WorksheetFunction.Clean(Trim(acsv(j, i))) <> "" Then
                    
                    If sCoding(k, 4) = "" Then
                        MsgBox "There is empty data..." & wsData.Cells(k + 2, 4).Address
                        Application.Goto wsData.Cells(k + 2, 4), True
                        Application.DisplayAlerts = False
                        Sheets("updateCSV").Delete
                        Application.DisplayAlerts = True
                        Exit Sub
                    End If
                    
                    If Trim(acsv(j, 1)) = Trim(sCoding(k, 1)) And Trim(acsv(1, i)) = Trim(sCoding(k, 2)) And Trim(sCoding(k, 8)) <> "" Then
                        acsv(j, i) = Trim(sCoding(k, 8))
                        Exit For
                    ElseIf Trim(acsv(j, 1)) = Trim(sCoding(k, 1)) And Trim(acsv(1, i)) = Trim(sCoding(k, 2)) And Trim(sCoding(k, 4)) <> "" Then
                    'If acsv(j, 1) = sCoding(k, 1) And acsv(1, i) = sCoding(k, 2) And Trim(sCoding(k, 4)) <> "" Then
                        acsv(j, i) = Trim(sCoding(k, 4))
                        Exit For
                    End If
                    
                Else
                    acsv(j, i) = ""
                    k = k - 1
                    Exit For
                End If
                
            Next k
            k = k + 1
        Next j
    Next i
    
    'If wsData.Range("I:I").EntireColumn.Hidden = True Then
    wsData.Range("E:G").EntireColumn.Hidden = True
    wsData.Range("I:L").EntireColumn.Hidden = True
    'End If
    
    wsCSV.Range("A4").CurrentRegion = acsv
    wsCSV.Activate
    Selection.WrapText = False
    
    'Save file
    Dim nmFile As String, XnmFile As String
    Dim index As Integer
    nmFile = ThisWorkbook.path & "\updateCSV " & _
    Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1)) & ".csv"
    ActiveSheet.Copy
    
    'Indexing file if exist file
    XnmFile = GetNextAvailableName(nmFile)
    ActiveWorkbook.SaveAs XnmFile, FileFormat:=xlCSV
    ActiveWorkbook.Close SaveChanges:=False
    
    Call TurnOnStuf
    
    MsgBox "CSV File Created :" & vbNewLine & XnmFile
        

End Sub
Sub TransferPerFrame()
    
    Call TurnOffStuf
    Call listFrame(k, aStart, aEnd, nmFrame)
    
    If wsData.Range("H:H").EntireColumn.Hidden = True Then
        wsData.Range("H:H").EntireColumn.Hidden = False
    End If
        
    Call NormalSetting
        
    sCoding = wsData.UsedRange
    
    For i = 0 To k
        For nCoding = aStart(i) To aStart(i) + aEnd(i) - 1
            sCoding(nCoding, 4) = Trim(sCoding(nCoding, 4))
            sCoding(nCoding, 4) = Replace(sCoding(nCoding, 4), ",", " ")
            sCoding(nCoding, 4) = Replace(sCoding(nCoding, 4), ".", " ")
            sCoding(nCoding, 4) = Replace(sCoding(nCoding, 4), ";", " ")
            Do While InStr(1, sCoding(nCoding, 4), "  ")
                sCoding(nCoding, 4) = Replace(sCoding(nCoding, 4), "  ", " ")
            Loop
            aCod = Split(sCoding(nCoding, 4), " ")
            sCoding(nCoding, 8) = ""
            For nCod = LBound(aCod) To UBound(aCod)
                For nFrame = LBound(aFrame, 1) + 1 To UBound(aFrame, 1)
                    If aFrame(nFrame, 1) <> "" Then
                        sFrame = Trim(LCase(aFrame(nFrame, 1)))
                    Else
                        sFrame = Trim(LCase(aFrame(nFrame, 2)))
                    End If
                    If Trim(LCase(aCod(nCod))) = Trim(sFrame) Then
                        sCoding(nCoding, 8) = Trim(sCoding(nCoding, 8) & " " & Trim(aFrame(nFrame, 2)))
                        Exit For
                    End If
                Next nFrame
            Next nCod
        Next nCoding
    Next i
    
    wsData.Range("A4").CurrentRegion = sCoding
    
    Call UpdatetoOpen
    
    Call TurnOnStuf
    
    

End Sub
Sub TransferAll()
       
    Dim nCoding As Long
    'Dim sCoding As String, bCoding As String
    Dim nFrame As Long
    Dim sFrame As String
    Dim bFrame() As String
    Dim nmFrame() As String
    Dim allFrame As String
    
    Dim aFrm() As String
    Dim aCod() As String
    Dim nFrm, nCod, i, j, k, m As Integer
    Dim n As Long
    Dim bFrm, bCod As Boolean
    
    Dim aStart As Variant, aEnd As Variant
      
    Set wb = ActiveWorkbook
    Set wsInfo = wb.Sheets("Info")
    Set wsData = wb.Sheets("Data")
    aData = wsData.UsedRange
        
    aInfo = wsInfo.Range("A4").CurrentRegion
            
    If wsData.Range("H:H").EntireColumn.Hidden = True Then
        wsData.Range("H:H").EntireColumn.Hidden = False
    End If
        
    Call NormalSetting
            
    Call TurnOffStuf
    
    sCoding = wsData.UsedRange
    
    ReDim nmFrame(0 To UBound(aInfo, 1))
    j = 0
    For i = 3 To UBound(aInfo, 1)
        If aInfo(i, 8) <> "" And InStr(allFrame, aInfo(i, 8)) = 0 Then
            ReDim Preserve nmFrame(j)
            nmFrame(j) = aInfo(i, 8)
            j = j + 1
        End If
        allFrame = allFrame & aInfo(i, 8)
    Next i
    n = 2
    
    For m = 0 To UBound(nmFrame, 1)
        
        Set wsFrame = wb.Sheets(nmFrame(m))
        aFrame = wsFrame.UsedRange
        
        j = -1
        Dim nmQuest() As String
        For i = 3 To UBound(aInfo, 1)
            If aInfo(i, 8) = nmFrame(m) Then
                j = j + 1
                ReDim Preserve nmQuest(j)
                nmQuest(j) = aInfo(i, 2)
            End If
        Next i
        
        k = UBound(nmQuest, 1)
        ReDim aStart(k), aEnd(k)
        
        For i = 0 To k
            For n = n To UBound(aData, 1)
                If aData(n, 2) = nmQuest(i) Then
                    aStart(i) = n
                    aEnd(i) = WorksheetFunction.CountIf(wsData.Range("B3:B" & UBound(aData, 1)), nmQuest(i))
                    Exit For
                End If
            Next n
            n = aEnd(i) + 1
        Next i
        
        For i = 0 To k
            For nCoding = aStart(i) To aStart(i) + aEnd(i) - 1
                sCoding(nCoding, 4) = Trim(sCoding(nCoding, 4))
                sCoding(nCoding, 4) = Replace(sCoding(nCoding, 4), ",", " ")
                sCoding(nCoding, 4) = Replace(sCoding(nCoding, 4), ".", " ")
                sCoding(nCoding, 4) = Replace(sCoding(nCoding, 4), ";", " ")
                Do While InStr(1, sCoding(nCoding, 4), "  ")
                    sCoding(nCoding, 4) = Replace(sCoding(nCoding, 4), "  ", " ")
                Loop
                aCod = Split(sCoding(nCoding, 4), " ")
                sCoding(nCoding, 8) = ""
                For nCod = LBound(aCod) To UBound(aCod)
                    For nFrame = LBound(aFrame, 1) + 2 To UBound(aFrame, 1)
                        If aFrame(nFrame, 1) <> "" Then
                            sFrame = Trim(LCase(aFrame(nFrame, 1)))
                        Else
                            sFrame = Trim(LCase(aFrame(nFrame, 2)))
                        End If
                        If Trim(LCase(aCod(nCod))) = Trim(sFrame) Then
                            sCoding(nCoding, 8) = Trim(sCoding(nCoding, 8) & " " & Trim(aFrame(nFrame, 2)))
                            Exit For
                        End If
                    Next nFrame
                Next nCod
            Next nCoding
        Next i
    Next m
    
    wsData.Range("A4").CurrentRegion = sCoding
    
    Call UpdatetoOpen
    
    Call TurnOnStuf

End Sub
    

Sub infoName()
    
    Set wb = ActiveWorkbook
    Set wsInfo = wb.Sheets("Info")
  
    aInfo = wsInfo.Range("A4").CurrentRegion
    
    Call TurnOffStuf
    Dim allFrame As String, allCoder As String
    Dim i, j, k As Integer
    Dim nFrame(), nQuest(), nCoder()
    ReDim nFrame(0 To UBound(aInfo, 1))
    ReDim nQuest(0 To UBound(aInfo, 1))
    ReDim nCoder(0 To UBound(aInfo, 1))
    j = 0
    k = 0
    For i = 2 To UBound(aInfo, 1)
        nQuest(i) = aInfo(i, 2)
        If aInfo(i, 8) <> "" And InStr(allFrame, aInfo(i, 8)) = 0 Then
            ReDim Preserve nFrame(j)
            nFrame(j) = aInfo(i, 8)
            j = j + 1
        End If
        allFrame = allFrame & aInfo(i, 8)
        If aInfo(i, 7) <> "" And InStr(allCoder, aInfo(i, 7)) = 0 Then
            ReDim Preserve nCoder(k)
            nCoder(k) = aInfo(i, 7)
            k = k + 1
        End If
        allCoder = allCoder & aInfo(i, 7)
    Next i
    
    Call TurnOnStuf

    
End Sub
Sub CreateIdentity()
        
    Call TurnOffStuf
    
    Dim i As Long
    Set wsData = wb.Sheets("Data")
    aData = wsData.UsedRange
    For i = 2 To UBound(aData, 1)
        wsData.Cells(i + 2, 17) = aData(i, 1) & aData(i, 2)
    Next i
    
    Call TurnOnStuf

End Sub


Sub AddCode()
Attribute AddCode.VB_ProcData.VB_Invoke_Func = "m\n14"
'https://www.excelcampus.com/library/find-the-first-used-cell-vba-macro/
    
'Ctrl + M --> AddCode in Sheet DATA column D with Code in Column F
'Setting Shortcut-->View-->Macros-->Select Macros Name-->Option
    Dim rFound As Range
    Dim RngTarget As Range, Target As Range
    Dim aData As Variant
    
    Set wb = ActiveWorkbook
    Set wsData = wb.Sheets("Data")
    aData = wsData.UsedRange
    
    
    
    On Error Resume Next
    Set RngTarget = Application.Intersect(ActiveCell, wsData.Range(Cells(4, 6), Cells(UBound(aData, 1) + 2, 6)))
    
    'Check double code in Column Coding
    If Not RngTarget Is Nothing Then
    
        'check duplicate keyword
        If wsData.Range(Cells(4, 5), Cells(UBound(aData, 1) + 2, 5)).Cells.SpecialCells(xlCellTypeConstants).count > 1 Then
            MsgBox "Double keyword to search, please choose the one"
            Exit Sub
        End If
    
        On Error Resume Next
        Set rFound = wsData.Range("E:E").Cells.Find(What:="?*", After:=wsData.Cells(3, 5), LookAt:=xlPart, LookIn:=xlValues, _
                        SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
        On Error GoTo 0
        
        If ActiveCell.Value <> "" And rFound.Offset(, -1) <> "" Then
            If InStr(rFound.Offset(, -1), ActiveCell) > 0 Then
                MsgBox "Code already in Coding"
                Exit Sub
            End If
            rFound.Offset(, -1) = rFound.Offset(, -1) & " " & ActiveCell.Value
            'Last Action for repeat coding
            Cells(4, 19).Value = ActiveCell.Value
            rFound.Select
        ElseIf ActiveCell.Value <> "" And rFound.Offset(, -1) = "" Then
            rFound.Offset(, -1) = ActiveCell.Value
            'Last Action for repeat coding
            Cells(4, 19).Value = ActiveCell.Value
            rFound.Select
        End If
        
    End If
    
    On Error GoTo 0
End Sub
Sub RepeatLastAction()
Attribute RepeatLastAction.VB_ProcData.VB_Invoke_Func = "l\n14"
'Ctrl + L --> Repeat last Action/AddCode in Sheet DATA column D
'Setting Shortcut-->View-->Macros-->Select Macros Name-->Option
    Set wb = ActiveWorkbook
    Set wsData = wb.Sheets("Data")
    
    Dim rgCode As Range
    'select range without select hide cells
    If Selection.Cells.count = 1 Then
        Set rgCode = ActiveCell
    Else
        Set rgCode = Selection.SpecialCells(xlCellTypeVisible)
    End If
    If Intersect(rgCode, wsData.Range("D:D")) Is Nothing Then
        MsgBox "Please select cell/range to add Code in Sheet:Data, column:Coding !!!"
        Exit Sub
    Else
        For Each cell In rgCode
            If cell.Value <> "" Then
                If InStr(cell, Cells(4, 19)) > 0 Then
                    MsgBox "Code already in Coding"
                    Exit Sub
                End If
                cell.Value = cell & " " & Cells(4, 19).Value
            Else
                cell.Value = Cells(4, 19).Value
            End If
        Next cell
    End If
End Sub
Sub UndoLastAction()
Attribute UndoLastAction.VB_ProcData.VB_Invoke_Func = "L\n14"
'Ctrl + Shift + L --> Undo last Action/ delete AddCode in Sheet DATA column D
'Setting Shortcut-->View-->Macros-->Select Macros Name-->Option
    Set wb = ActiveWorkbook
    Set wsData = wb.Sheets("Data")
    
    Dim rgCode As Range
    'select range without select hide cells
    'select range without select hide cells
    If Selection.Cells.count = 1 Then
        Set rgCode = ActiveCell
    Else
        Set rgCode = Selection.SpecialCells(xlCellTypeVisible)
    End If
    
    If Intersect(rgCode, wsData.Range("D:D")) Is Nothing Then
        MsgBox "Please select cell/range to add Code in Sheet:Data, column:Coding !!!"
        Exit Sub
    Else
        For Each cell In rgCode
'            If cell.Value <> "" Then
            If cell.Value = Cells(4, 19).Value Then
                cell.Value = ""
            'End If
                'cell.Value = cell & " " & Cells(4, 19).Value
            ElseIf InStr(cell, Cells(4, 19)) > 0 And Len(cell) > Len(Cells(4, 19)) Then
                cell.Value = Left(cell, (Len(cell) - (Len(Cells(4, 19)) + 1)))
            End If
        Next cell
    End If
    
'    With Application
'        .EnableEvents = False
'        .Undo
'        .EnableEvents = True
'    End With

End Sub
