Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Private Sub cmdGenerate_Click()
Module3.Create_XML
End Sub
Private Sub cmdNext_Click()
    Sheet2.Activate

End Sub


Private Sub cmdPrev_Click()
    Sheet1.Activate

End Sub

Private Sub cmdHelp_Click()
Sheet30.Visible = xlSheetVisible
Sheet30.Activate
'Sheet30.Range("i_tds").Select
End Sub

Private Sub cmdPrint_Click()
Module3.PrintWorksheets
End Sub

Private Sub cmdValidate_Click()
Module3.printerrormessage_TDSal
End Sub
' For Sheet : TDS
Private Sub Worksheet_Change(ByVal Target As Range)

On Error GoTo exit1

Application.EnableEvents = False
If Target.Validation.Type = 3 Then
     GoTo exit1
End If

Target.value = UCase(Target.value)

     If (getRangeName(Target) = "TDSal.TAN") Then
         If Not (ValidateTAN_TDSal()) Then
             MsgBox "INVALID TAN"
         End If
         GoTo exit1
     End If


     If (getRangeName(Target) = "TDSoth.TAN") Then
         If Not (ValidateTAN_TDSoth()) Then
             MsgBox "INVALID TAN"
         End If
         GoTo exit1
     End If



     If (getRangeName(Target) = "TaxP.DateDep") Then
         If Not (ValidateDateDep_TaxP()) Then
             MsgBox "INVALID DateDep"
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

