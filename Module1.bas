Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Function IsEmailValid(strEmail)
    Dim strArray As Variant
    Dim strItem As Variant
    Dim i As Long, c As String, blnIsItValid As Boolean
    blnIsItValid = True

    Dim pos%
    pos = InStr(strEmail, "@")
    If pos = 0 OR InStr(pos + 1, strEmail, "@") <> 0 Then IsEmailValid = False: Exit Function
    ReDim strArray(1 To 2)
    strArray(1) = Left(strEmail, pos - 1)
    strArray(2) = Right(strEmail, Len(strEmail) - pos)
    For Each strItem In strArray
        If Len(strItem) <= 0 Then
            blnIsItValid = False
            IsEmailValid = blnIsItValid
            Exit Function
        End If
        For i = 1 To Len(strItem)
            c = LCase(Mid(strItem, i, 1))
            If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
                blnIsItValid = False
                IsEmailValid = blnIsItValid
                Exit Function
            End If
        Next i
        If Left(strItem, 1) = "." Or Right(strItem, 1) = "." Then
            blnIsItValid = False
            IsEmailValid = blnIsItValid
            Exit Function
        End If
    Next strItem
    If InStr(strArray(2), ".") <= 0 Then
        blnIsItValid = False
        IsEmailValid = blnIsItValid
        Exit Function
    End If
'    i = Len(strArray(2)) - InStrRev(strArray(2), ".")
'    If i <> 2 And i <> 4 Then
'        blnIsItValid = False
'        IsEmailValid = blnIsItValid
'        Exit Function
'    End If
    If InStr(strEmail, "..") > 0 Then
        blnIsItValid = False
        IsEmailValid = blnIsItValid
        Exit Function
    End If
    IsEmailValid = blnIsItValid
End Function

