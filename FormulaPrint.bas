Attribute VB_Name = "FormulaPrint"

Function GETFORMULA(formCell As Range, varCell As Range, Optional includeLeftHandSide = False) As String
    Application.Volatile
    Dim form As String
    Dim varCol As String
    Dim formCol As String
    Dim ws As Worksheet
    Set ws = formCell.Worksheet
    
    formCol = Left(formCell.Address(False, False), Len(formCell.Address(False, False)) - Len(Format(formCell.Row, "0")))
    varCol = Left(varCell.Address(False, False), Len(varCell.Address(False, False)) - Len(Format(varCell.Row, "0")))
    form = Replace(UCase(formCell.Formula), "$", "")
    
    ' create RegExp matches array for row references
    Dim re As Object, matches As Variant
    Set re = CreateObject("VBScript.RegExp") 'late binding
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = "[A-Z]+\$?(\d+)"
    Set matches = re.Execute(form)
    If matches.Count = 0 Then Exit Function

    ' add matches to array
    Dim arr() As Variant
    ReDim arr(matches.Count - 1)
    For i = 0 To matches.Count - 1
        arr(i) = matches(i).SubMatches(0)
    Next
    
    ' sort array
    Call QuickSort(arr, LBound(arr), UBound(arr))
    
    ' replace references with variables
    Dim var As String, r As String
    For i = UBound(arr) To LBound(arr) Step -1
        r = arr(i)
        r = Replace(r, "=", "")
        var = Replace(Trim(ws.Range(varCol & r).Value()), "=", "")
        form = Replace(form, formCol & r, var)
    Next i
    
    ' clean formula
    Call CleanFormula(form)
    If includeLeftHandSide Then form = varCell & form
    
    GETFORMULA = form
End Function

Private Function CleanFormula(ByRef form As String) As String
    
    For i = 0 To Len(form)
        ' ABS() to |n| format
        If InStr(form, "ABS(") Then
            parInd1 = InStr(1, form, "ABS(") + 3
            parInd2 = InStr(parInd1, form, ")")
            form = WorksheetFunction.Replace(form, parInd2, 1, "|")
            form = Replace(form, "ABS(", "|", 1, 1)
        End If
        
        ' PI() to pi character
        If InStr(form, "PI()") > 0 Then form = Replace(form, "PI()", ChrW(960))
        
    Next
End Function


' sorting function
' https://stackoverflow.com/questions/152319/vba-array-sort-function
Private Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub
