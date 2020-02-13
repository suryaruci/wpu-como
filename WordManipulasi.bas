Attribute VB_Name = "WordManipulasi"


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
