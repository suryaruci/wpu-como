Attribute VB_Name = "Support"
Option Explicit

'GetSystemMetrics32 info: http://msdn.microsoft.com/en-us/library/ms724385(VS.85).aspx
#If Win64 Then
    Private Declare PtrSafe Function GetSystemMetrics32 Lib "User32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
#ElseIf Win32 Then
    Private Declare Function GetSystemMetrics32 Lib "User32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
#End If

'VBA Wrappers:
Public Function dllGetMonitors() As Long
    Const SM_CMONITORS = 80
    dllGetMonitors = GetSystemMetrics32(SM_CMONITORS)
End Function

Public Function dllGetHorizontalResolution() As Long
    Const SM_CXVIRTUALSCREEN = 78
    dllGetHorizontalResolution = GetSystemMetrics32(SM_CXVIRTUALSCREEN)
End Function

Public Function dllGetVerticalResolution() As Long
    Const SM_CYVIRTUALSCREEN = 79
    dllGetVerticalResolution = GetSystemMetrics32(SM_CYVIRTUALSCREEN)
End Function

Public Sub ShowDisplayInfo()
'https://riptutorial.com/vba/example/31840/get-total-monitors-and-screen-resolution
    Debug.Print "Total monitors: " & vbTab & vbTab & dllGetMonitors
    Debug.Print "Horizontal Resolution: " & vbTab & dllGetHorizontalResolution
    Debug.Print "Vertical Resolution: " & vbTab & dllGetVerticalResolution

    'Total monitors:         1
    'Horizontal Resolution:  1920
    'Vertical Resolution:    1080
End Sub


Function GetNextAvailableName(ByVal nmFile As String) As String
'https://stackoverflow.com/questions/31703554/save-with-a-different-name-if-the-file-already-exists-in-directory?rq=1
'check existing file in same directory
    With CreateObject("Scripting.FileSystemObject")

        Dim strFolder As String, strBaseName As String, strExt As String, i As Long
        strFolder = .GetParentFolderName(nmFile)
        strBaseName = .getbaseName(nmFile)
        strExt = .GetExtensionName(nmFile)

        Do While .FileExists(nmFile)
            i = i + 1
            nmFile = .BuildPath(strFolder, strBaseName & " rev" & Format(i, "00") & "." & strExt)
        Loop

    End With

    GetNextAvailableName = nmFile

End Function
Sub TurnOffStuf()
    'speedup
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
End Sub
Sub TurnOnStuf()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
Sub removeBlankFromula(x)
    
    'Remove blank formula
    'https://stackoverflow.com/questions/42342709/vba-copy-only-non-blank-cells
    x.UsedRange.Cells.Replace What:="", Replacement:="pneumonoultramicroscopicsilicovolcanoconiosis", LookAt:= _
    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    x.UsedRange.Cells.Replace What:="pneumonoultramicroscopicsilicovolcanoconiosis", Replacement:="", LookAt:= _
    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    
End Sub
Function WorksheetExists2(WorksheetName As String, Optional wb As Workbook) As Boolean
'https://www.exceltip.com/files-workbook-and-worksheets-in-vba/determine-if-a-sheet-exists-in-a-workbook-using-vba-in-microsoft-excel.html
    If wb Is Nothing Then Set wb = ThisWorkbook
    With wb
        On Error Resume Next
        WorksheetExists2 = (.Sheets(WorksheetName).Name = WorksheetName)
        On Error GoTo 0
    End With
End Function
Sub NormalSetting()
    ' Remove the any existing filters and sort by index
    If wsData.FilterMode = True Then
        wsData.ShowAllData
    End If
    'Makesure sorting from header
    wsData.Rows("1:2").ClearContents
    
    'sorting by index
    wsData.UsedRange.Sort key1:=[P4], order1:=xlAscending, Header:=xlYes

End Sub
Sub FormatHeader()
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
End Sub

Function getnmFile(ByVal nmFile As String) As String
    'Get name file
    With CreateObject("Scripting.FileSystemObject")
        Dim temp() As String
        nmFile = CreateObject("Scripting.FileSystemObject").getbaseName(ThisWorkbook.Name)
        temp = Split(nmFile, " ")
        nmFile = ""
        For i = 2 To UBound(temp)
            If UBound(temp) = 2 Then
                nmFile = temp(i)
            Else
                Trim(nmFile) = nmFile & " " & temp(i)
            End If
        Next i
    End With
    getnmFile = nmFile
End Function

