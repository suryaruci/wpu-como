Attribute VB_Name = "Library"
Sub TestMacro()
    MsgBox "On Progress...."
End Sub

Sub StrPercent()
'https://www.mrexcel.com/forum/excel-questions/702473-vba-compare-strings-percentage.html
    Call TurnOffStuf
'    Application.ScreenUpdating = False
    LR = Sheets("Sh1").Cells(Rows.count, 1).End(xlUp).row
    With Sheets("Sh3")
        .Range("A1:A" & LR).Value = Sheets("Sh1").Range("A1:A" & LR).Value
        .Range("B1:B" & LR).Value = Sheets("Sh2").Range("A1:A" & LR).Value
        For r = 2 To LR
            If Trim(.Cells(r, "A").Value) = Trim(.Cells(r, "B").Value) Then
                .Cells(r, "C").Value = 1
                GoTo Done
            End If
            str1 = Trim(.Cells(r, "A").Value)
            str2 = Trim(.Cells(r, "B").Value)
            Len1 = Len(str1)
            Len2 = Len(str2)
            same = 0
            For c = 1 To Len2
                If InStr(1, str1, Mid(str2, c, 1), 1) Then
                    same = same + 1
                    str1 = Replace(str1, Mid(str2, c, 1), "*", 1, 1)
                End If
            Next c
            .Cells(r, "C").Value = same / Len1
Done:
        Next r
    End With
    Call TurnOnStuf
'    Application.ScreenUpdating = True
End Sub
Public Sub SaveCSV(acsv() As String, path As String, Optional Delim As String = ",", Optional quote As String = "")
Dim opf As Long
Dim row As Long
Dim column As Long
 
    If Dir(path) <> "" Then Kill path
 
    opf = FreeFile
    Open path For Binary As #opf
 
    For row = LBound(acsv, 1) To UBound(acsv, 1)
        For column = LBound(acsv, 2) To UBound(acsv, 2)
            Put #opf, , quote & Array(row, column) & quote
            If column < UBound(acsv, 2) Then Put #opf, , Delim
        Next column
        If row < UBound(acsv, 1) Then Put #opf, , vbCrLf
    Next row
    
    Close #opf
 
End Sub

Function Test() As String
    Test = ActiveCell
    'test = x
    Dim nameQuest As String
    Dim found As Range
    'ActiveCell.Offset(0, -6).Activate
    nameQuest = ActiveCell
    With Sheets("DATA").Range("B:B")
        Set found = .Find(nameQuest, LookIn:=xlValues)
        If found Is Nothing Then
            MsgBox "Not found Question/Verbatim for " & nameQuest
            Application.DisplayAlerts = False
            Sheets("info").Activate
            Application.DisplayAlerts = True
            Exit Function
        Else
            Sheets("Data").Activate
            found.Select
        End If
    End With

End Function
Sub frm()
    frm1.Show
End Sub


' ExcelMacroMastery.com
'
' Description: Read data from another workbook into the current one using Insert.
' This will still work if the workbook is closed.
' Source worksheet: Sales
' Destination worksheet: CompanyOut
' Author: Paul Kelly
Private Sub ReadFromDifferentWorkbook()
    
    Dim connection As New ADODB.connection
    
    Dim sourceFile As String
    sourceFile = ThisWorkbook.path & Application.PathSeparator & "Source.xlsx"
     
    
    connection.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & _
              ";Extended Properties=""Excel 8.0;HDR=NO;"";"

    Dim sourceSheet As String
    sourceSheet = "[Excel 12.0;HDR=NO;DATABASE=" & sourceFile & "]"
    
    Dim query As String
    query = "Insert Into [Sheet1$] Select STATUS From " & sourceSheet & ".[Verbatim$]"
    'query = "Insert Into [CompanyOut$] Select Company,Sales From " & sourceSheet & ".[Sales$] "
    'query = "Insert into [Out$] SELECT Fruit,Sales  FROM [Simple$] Where Fruit='Orange'"
    'query = "Insert Into [Out$] From " & sourceSheet
    connection.Execute query

    connection.Close
End Sub

Sub Get_Data_From_File()
'https://www.youtube.com/watch?v=h_sC6Uwtwxk
    Dim FileToOpen As Variant
    Dim OpenBook As Workbook
    Application.ScreenUpdating = False
    FileToOpen = Application.GetOpenFilename(Title:="Browse for your File & Import Range", FileFilter:="Excel Files (*.xls*),*xls*")
    If FileToOpen <> False Then
        Set OpenBook = Application.Workbooks.Open(FileToOpen)
        OpenBook.Sheets("Verbatim").Range("A1:E20").Copy
        ThisWorkbook.Worksheets("SelectFile").Range("A3").PasteSpecial xlPasteValues
        OpenBook.Close False
    End If
    Application.ScreenUpdating = True
End Sub
Sub Coding()

    Dim nCoding As Long
    Dim sCoding As String
    Dim nFrame As Long
    Dim sFrame As String
    
    Dim aFrm
    Dim nFrm, LcRow As Integer
    Dim bFrm As Boolean
    
    
    Set wb = ActiveWorkbook
    Set wsData = wb.Sheets("Data")
    Set wsFrame = wb.Sheets("Q28a")
    
    Call TurnOffStuf
    aCoding = wsData.UsedRange
    aFrame = wsFrame.UsedRange
    LcRow = Cells(Rows.count, 3).End(xlUp).row
    'if range(cells(4,4),cells(lcrow,4)).Application.
    For nCoding = 2 To UBound(aCoding, 1)
'        Cells(nCoding, 1).Value = nCoding - 1
        sCoding = Trim(LCase(aCoding(nCoding, 3)))
        For nFrame = 2 To UBound(aFrame, 1)
            sFrame = Trim(LCase(aFrame(nFrame, 3)))
            aFrm = Split(sFrame, "/")
            bFrm = False
            For nFrm = LBound(aFrm) To UBound(aFrm)
                If Trim(LCase(sCoding)) = Trim(aFrm(nFrm)) Then
                    'wsData.Cells(nCoding, 3) = aFrame(nFrame, 2)
                    aCoding(nCoding, 4) = aFrame(nFrame, 2)
                    bFrm = True
                    Exit For
                End If
            Next nFrm
            If bFrm Then Exit For
        Next nFrame
        
    Next nCoding
    wsData.Cells(3, 1).Select
    Range("A3").Resize(LcRow - 2, 4) = aCoding
    
    Call TurnOnStuf
End Sub

Function IsInArray(ByVal stringToBeFound As String, ByVal arr As Variant) As Boolean
Dim element
For Each element In arr
    If element = stringToBeFound Then
        IsInArray = True
        Exit Function
    End If
Next element
End Function

Sub Verificationgakkepake()

    Dim CurCell As Object
    
    Call TurnOffStuf
    
    Set wb = ActiveWorkbook
    Set wsData = wb.Sheets("Data")
    Set wsFrame = wb.Sheets("Story")
    aData = wsData.UsedRange
    aFrame = wsFrame.UsedRange
    
    'Please save Keywords in column 1 of Sheet2 and corresponding codes in column 2 in the same sheet
    'Save Text in column 1 of Sheet1 after running this macro it will generate codes in column 2 in this sheet
    'Dim txt As Variant
    Dim x, i, key_w, y, z As Integer
    Dim searchrange, foundcell As Range
  
    'Range(Cells(4, 8), Cells(n, 8)).Value = Range(Cells(4, 4), Cells(n, 4)).Value
    x = 0
    y = 1
    For i = 4 To UBound(aData, 1)
        wsData.Cells(i, 9) = ""
        Set searchrange = wsData.Cells(i, 4)
        y = Len(Replace(Replace(Replace(searchrange, " ", ""), ";", ""), ",", ""))
        'Call JaroWink(wsFrame.Cells(key_w, 1), searchrange)
        wsData.Cells(i, 4).Font.Color = vbRed
        For key_w = 5 To UBound(aFrame, 1)
            'Set foundcell = JaroWink(wsFrame.Cells(key_w, 1), searchrange)
            Set foundcell = searchrange.Find(What:=wsFrame.Cells(key_w, 2))
            'If foundcell > 0.6 Then
            If Not foundcell Is Nothing Then
                'y = InStr(1, searchrange, wsFrame.Cells(key_w, 2), vbTextCompare)
                'searchrange.Characters(y, Len(wsFrame.Cells(key_w, 2))).Font.Color = vbRed
                If (Len(wsData.Cells(i, 4)) > 0) Then
                    x = InStr(1, searchrange.Text, wsFrame.Cells(key_w, 2), vbTextCompare)
                    If x > 0 Then searchrange.Characters(x, Len(wsFrame.Cells(key_w, 2))).Font.Color = vbBlack
                    'If CurCell.Value = wsFrame.Cells(key_w, 2) Then CurCell.Interior.Color = RGB(0, 204, 0)
     '               wsData.Cells(i, 9) = wsData.Cells(i, 9) & Chr(10) & wsFrame.Cells(key_w, 2) & " | " & wsFrame.Cells(key_w, 3)
                Else
                    x = InStr(1, searchrange.Text, wsFrame.Cells(key_w, 2), vbTextCompare)
                    If x > 0 Then searchrange.Characters(x, Len(wsFrame.Cells(key_w, 2))).Font.Color = vbBlack
     '               wsData.Cells(i, 9) = wsFrame.Cells(key_w, 2) & " | " & wsFrame.Cells(key_w, 3)
                End If
                z = Len(wsFrame.Cells(key_w, 2)) + z
                If y = z Then
                    Exit For
                End If
            End If
        Next
        x = 0
        z = 0
    Next
    
    Call TurnOnStuf

End Sub
Public Function Similarity(ByVal string1 As String, _
    ByVal string2 As String, _
    Optional ByRef RetMatch As String, _
    Optional min_match = 1) As Single
Dim b1() As Byte, b2() As Byte
Dim lngLen1 As Long, lngLen2 As Long
Dim lngResult As Long

If UCase(string1) = UCase(string2) Then
    Similarity = 1
Else:
    lngLen1 = Len(string1)
    lngLen2 = Len(string2)
    If (lngLen1 = 0) Or (lngLen2 = 0) Then
        Similarity = 0
    Else:
        b1() = StrConv(UCase(string1), vbFromUnicode)
        b2() = StrConv(UCase(string2), vbFromUnicode)
        lngResult = Similarity_sub(0, lngLen1 - 1, _
        0, lngLen2 - 1, _
        b1, b2, _
        string1, _
        RetMatch, _
        min_match)
        Erase b1
        Erase b2
        If lngLen1 >= lngLen2 Then
            Similarity = lngResult / lngLen1
        Else
            Similarity = lngResult / lngLen2
        End If
    End If
End If

End Function

Private Function Similarity_sub(ByVal start1 As Long, ByVal end1 As Long, _
                                ByVal start2 As Long, ByVal end2 As Long, _
                                ByRef b1() As Byte, ByRef b2() As Byte, _
                                ByVal FirstString As String, _
                                ByRef RetMatch As String, _
                                ByVal min_match As Long, _
                                Optional recur_level As Integer = 0) As Long
'* CALLED BY: Similarity *(RECURSIVE)

Dim lngCurr1 As Long, lngCurr2 As Long
Dim lngMatchAt1 As Long, lngMatchAt2 As Long
Dim i As Long
Dim lngLongestMatch As Long, lngLocalLongestMatch As Long
Dim strRetMatch1 As String, strRetMatch2 As String

If (start1 > end1) Or (start1 < 0) Or (end1 - start1 + 1 < min_match) _
Or (start2 > end2) Or (start2 < 0) Or (end2 - start2 + 1 < min_match) Then
    Exit Function '(exit if start/end is out of string, or length is too short)
End If

For lngCurr1 = start1 To end1
    For lngCurr2 = start2 To end2
        i = 0
        Do Until b1(lngCurr1 + i) <> b2(lngCurr2 + i)
            i = i + 1
            If i > lngLongestMatch Then
                lngMatchAt1 = lngCurr1
                lngMatchAt2 = lngCurr2
                lngLongestMatch = i
            End If
            If (lngCurr1 + i) > end1 Or (lngCurr2 + i) > end2 Then Exit Do
        Loop
    Next lngCurr2
Next lngCurr1

If lngLongestMatch < min_match Then Exit Function

lngLocalLongestMatch = lngLongestMatch
RetMatch = ""

lngLongestMatch = lngLongestMatch _
+ Similarity_sub(start1, lngMatchAt1 - 1, _
start2, lngMatchAt2 - 1, _
b1, b2, _
FirstString, _
strRetMatch1, _
min_match, _
recur_level + 1)
If strRetMatch1 <> "" Then
    RetMatch = RetMatch & strRetMatch1 & "*"
Else
    RetMatch = RetMatch & IIf(recur_level = 0 _
    And lngLocalLongestMatch > 0 _
    And (lngMatchAt1 > 1 Or lngMatchAt2 > 1) _
    , "*", "")
End If


RetMatch = RetMatch & Mid$(FirstString, lngMatchAt1 + 1, lngLocalLongestMatch)


lngLongestMatch = lngLongestMatch _
+ Similarity_sub(lngMatchAt1 + lngLocalLongestMatch, end1, _
lngMatchAt2 + lngLocalLongestMatch, end2, _
b1, b2, _
FirstString, _
strRetMatch2, _
min_match, _
recur_level + 1)

If strRetMatch2 <> "" Then
    RetMatch = RetMatch & "*" & strRetMatch2
Else
    RetMatch = RetMatch & IIf(recur_level = 0 _
    And lngLocalLongestMatch > 0 _
    And ((lngMatchAt1 + lngLocalLongestMatch < end1) _
    Or (lngMatchAt2 + lngLocalLongestMatch < end2)) _
    , "*", "")
End If

Similarity_sub = lngLongestMatch

End Function
Function RemovePunctuation(Txt As String) As String
'https://www.extendoffice.com/documents/excel/3296-excel-remove-all-punctuation.html
    With CreateObject("VBScript.RegExp")
        .Pattern = "[^A-Z0-9 ]"
        .IgnoreCase = True
        .Global = True
        RemovePunctuation = .Replace(Txt, "")
        If Left(RemovePunctuation, 2) = "me" And Right(RemovePunctuation, 1) = "i" Then
            RemovePunctuation = Mid(RemovePunctuation, 3, Len(RemovePunctuation) - 3)
            Exit Function
        End If
        If (Left(RemovePunctuation, 4) = "meny" Or Left(RemovePunctuation, 4) = "peny") And Right(RemovePunctuation, 3) = "kan" Then
            RemovePunctuation = Mid(RemovePunctuation, 5, Len(RemovePunctuation) - 7)
            RemovePunctuation = "s" & RemovePunctuation
            Exit Function
        End If
        If Left(RemovePunctuation, 4) = "meng" And Right(RemovePunctuation, 3) = "kan" Then
            RemovePunctuation = Mid(RemovePunctuation, 5, Len(RemovePunctuation) - 7)
            Exit Function
        End If
        If Left(RemovePunctuation, 2) = "me" And Right(RemovePunctuation, 3) = "kan" Then
            RemovePunctuation = Mid(RemovePunctuation, 3, Len(RemovePunctuation) - 5)
            Exit Function
        End If
        If Left(RemovePunctuation, 3) = "ber" Then
            'RemovePunctuation = Mid(RemovePunctuation, 4, Len(RemovePunctuation) - 3)
            'Exit Function
        End If
        
    End With
    
End Function

Function FuzzyMatchByWord(ByVal lsPhrase1 As String, ByVal lsPhrase2 As String, Optional lbStripVowels As Boolean = False, Optional lbDiscardExtra As Boolean = False) As Double

'
' Compare two phrases and return a similarity value (between 0 and 100).
'
' Arguments:
'
' 1. Phrase1        String; any text string
' 2. Phrase2        String; any text string
' 3. StripVowels    Optional to strip all vowels from the phrases
' 4. DiscardExtra   Optional to discard any unmatched words
'
   
    'local variables
    Dim lsWord1() As String
    Dim lsWord2() As String
    Dim ldMatch() As Double
    Dim ldCur As Double
    Dim ldMax As Double
    Dim liCnt1 As Integer
    Dim liCnt2 As Integer
    Dim liCnt3 As Integer
    Dim lbMatched() As Boolean
    Dim lsNew As String
    Dim lsChr As String
    Dim lsKeep As String
   
    'set default value as failure
    FuzzyMatchByWord = 0
   
    'create list of characters to keep
    lsKeep = "BCDFGHJKLMNPQRSTVWXYZ0123456789 "
    If Not lbStripVowels Then
        lsKeep = lsKeep & "AEIOU"
    End If
   
    'clean up phrases by stripping undesired characters
    'phrase1
    lsPhrase1 = Trim$(UCase$(lsPhrase1))
    lsNew = ""
    For liCnt1 = 1 To Len(lsPhrase1)
        lsChr = Mid$(lsPhrase1, liCnt1, 1)
        If InStr(lsKeep, lsChr) <> 0 Then
            lsNew = lsNew & lsChr
        End If
    Next
    lsPhrase1 = lsNew
    lsPhrase1 = Replace(lsPhrase1, "  ", " ")
    lsWord1 = Split(lsPhrase1, " ")
    If UBound(lsWord1) = -1 Then
        Exit Function
    End If
    ReDim ldMatch(UBound(lsWord1))
    'phrase2
    lsPhrase2 = Trim$(UCase$(lsPhrase2))
    lsNew = ""
    For liCnt1 = 1 To Len(lsPhrase2)
        lsChr = Mid$(lsPhrase2, liCnt1, 1)
        If InStr(lsKeep, lsChr) <> 0 Then
            lsNew = lsNew & lsChr
        End If
    Next
    lsPhrase2 = lsNew
    lsPhrase2 = Replace(lsPhrase2, "  ", " ")
    lsWord2 = Split(lsPhrase2, " ")
    If UBound(lsWord2) = -1 Then
        Exit Function
    End If
    ReDim lbMatched(UBound(lsWord2))
   
    'exit if empty
    If Trim$(lsPhrase1) = "" Or Trim$(lsPhrase2) = "" Then
        Exit Function
    End If
   
    'compare words in each phrase
    For liCnt1 = 0 To UBound(lsWord1)
        ldMax = 0
        For liCnt2 = 0 To UBound(lsWord2)
            If Not lbMatched(liCnt2) Then
                ldCur = FuzzyMatch(lsWord1(liCnt1), lsWord2(liCnt2))
                If ldCur > ldMax Then
                    liCnt3 = liCnt2
                    ldMax = ldCur
                End If
            End If
        Next
        lbMatched(liCnt3) = True
        ldMatch(liCnt1) = ldMax
    Next
   
    'discard extra words
    ldMax = 0
    For liCnt1 = 0 To UBound(ldMatch)
        ldMax = ldMax + ldMatch(liCnt1)
    Next
    If lbDiscardExtra Then
        liCnt2 = 0
        For liCnt1 = 0 To UBound(lbMatched)
            If lbMatched(liCnt1) Then
                liCnt2 = liCnt2 + 1
            End If
        Next
    Else
        liCnt2 = UBound(lsWord2) + 1
    End If
   
    'return overall similarity
    FuzzyMatchByWord = 100 * (ldMax / liCnt2)
   
End Function

Function FuzzyMatch(Fstr As String, Sstr As String) As Double

'
' Code sourced from: http://www.mrexcel.com/pc07.shtml
' Credited to: Ed Acosta
' Modified: Joe Stanton
'

    Dim l, L1, L2, m, SC, T, r As Integer
   
    l = 0
    m = 0
    SC = 1
   
    L1 = Len(Fstr)
    L2 = Len(Sstr)
   
    Do While l < L1
        l = l + 1
        For T = SC To L1
            If Mid$(Sstr, l, 1) = Mid$(Fstr, T, 1) Then
                m = m + 1
                SC = T
                T = L1 + 1
            End If
        Next T
    Loop
   
    If L1 = 0 Then
        FuzzyMatch = 0
    Else
        FuzzyMatch = m / L1
    End If

End Function
Sub GetConcern()
    Dim i As Long, j As Long
    
    For i = 1 To 4
        For j = 1 To Len(Cells(i, 1).Value)
            If Cells(i, 1).Characters(Start:=j, Length:=1).Font.Color = vbRed Then
                Cells(i, 2).Value = Cells(i, 2).Value & Mid(Cells(i, 1), j, 1)
           End If
        Next j
    Next i
End Sub
Sub win()
Dim myWindow1 As Window, myWindow2 As Window
Set myWindow1 = ActiveWindow
Set myWindow2 = myWindow1.NewWindow
Dim w As Long, h As Long
w = GetSystemMetrics32(0)
h = GetSystemMetrics32(1)

With myWindow1
    .WindowState = xlNormal
    .Top = 0
    .Left = 0
    .Height = h '* 0.75 'Application.UsableHeight
    .Width = w '* 0.75 'Application.UsableWidth * 0.75
End With
With myWindow2
    .WindowState = xlNormal
    .Top = 0
    .Left = (Application.UsableWidth * 0.25) + 1
    .Height = h * 0.25 'Application.UsableHeight
    .Width = w * 0.25 'Application.UsableWidth * 0.75
End With
End Sub

