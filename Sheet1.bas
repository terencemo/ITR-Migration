Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Private Sub cmdCalculate_Click()
Module3.validate_xmls
Sheet5.cmdTax_Click
Sheet5.cmdTaxTransfer_Click
Sheet5.cmdInterest_Click
Sheet5.cmdInterestTransfer_Click

End Sub

Private Sub cmdGenerate_Click()
Module3.Create_XML



End Sub

Private Sub cmdImport_Click()
Module3.Import
End Sub


Private Sub cmdNext_Click()
Sheet3.Activate
End Sub

Private Sub cmdHelp_Click()
Sheet30.Visible = xlSheetVisible
Sheet30.Activate
'Sheet30.Range("i_general").Select
End Sub

Private Sub cmdPrint_Click()
Module3.PrintWorksheets
End Sub

Private Sub cmdValidate_Click()
Module3.printerrormessage_gen1
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo exit1


Application.EnableEvents = False
If Target.Validation.Type = 3 Then
'If (getRangeName(Target) = "sheet1.Status") Then
'Sheet1.Range("sheet1.Gender1") = "X"
'Sheet1.Range("sheet1.EmployerCategory1") = "NA"
'
'GoTo exit1
'End If


GoTo exit1
End If

If (getRangeName(Target) <> "sheet1.EmailAddress") Then
    Target.value = UCase(Target.value)
Else
Dim emailaddress As String
emailaddress = Target.value
If Trim(Len(emailaddress)) > 0 Then
   Range("sheet1.EmailAddress").Font.Underline = xlUnderlineStyleNone
End If
End If




If (getRangeName(Target) = "sheet1.PAN") Then
If Not (ValidatePAN(Target.value)) Then
    MsgBox "INVALID PAN"
End If
GoTo exit1
End If

If (getRangeName(Target) = "sheet1.DOB") Then
If Not (ValidateDOB_1) Then
    MsgBox "INVALID DATE"
End If
GoTo exit1
End If

If (getRangeName(Target) = "sheet1.OrigRetFiledDate") Then
If Not (ValidateOrigRetFiledDate_1) Then
    MsgBox "INVALID DATE"
End If
GoTo exit1
End If

    If (getRangeName(Target) = "sheet1.PinCode") Then
         If Not (ValidatePinCode_1()) Then
             MsgBox "INVALID PinCode"
         End If
         GoTo exit1
     End If

     If (getRangeName(Target) = "sheet1.STDcode") Then
         If Not (ValidateSTDcode_1()) Then
             MsgBox "INVALID STDcode"
         End If
         GoTo exit1
     End If

     If (getRangeName(Target) = "sheet1.PhoneNo") Then
         If Not (ValidatePhoneNo_1()) Then
             MsgBox "INVALID PhoneNo"
         End If
         GoTo exit1
     End If

If (getRangeName(Target) = "sheet1.ReceiptNo") Or (getRangeName(Target) = "sheet1.OrigRetFiledDate") Then
If Not isrevised Then
 If Len(Target.value) > 0 Then
 MsgBox "These fields are only for revised returns"
 Range(Target.Address).Select
GoTo exit1
End If
End If
GoTo exit1
End If

exit1:
Application.EnableEvents = True

End Sub

Function getRangeName(ByVal Target As Range) As String
Dim start As Integer
start = Len(Target.Name)

If InStr(1, Target.Name.Name, "'") > 0 Then
start = start + 3 + 1
Else
start = start + 1 + 1
End If

If InStr(1, Target.Name.Name, ThisComponent.CurrentController.getActiveSheet.Name) = 0 Then
start = 1
End If
getRangeName = Mid(Target.Name.Name, start)

End Function

Sub just()
Application.EnableEvents = True

End Sub

Function isrevised() As Boolean
If Mid(Range("sheet1.ReturnType1").value, 1, 1) = "R" Then
isrevised = True
Else
isrevised = False
End If
End Function

Function is44ab() As Boolean
If Mid(Range("sheet1.LiableSec44ABflg").value, 1, 1) = "Y" Then
is44ab = True
Else
is44ab = False
End If
End Function

Function isrep() As Boolean
If Mid(Range("sheet1.AsseseeRepFlg").value, 1, 1) = "Y" Then
isrep = True
Else
isrep = False
End If
End Function

