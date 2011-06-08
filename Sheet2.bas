Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Private Sub cmdGenerate_Click()
Module3.Create_XML
End Sub

Private Sub cmdNext_Click()
    Sheets("TDS").Activate
End Sub
Private Sub cmdPrev_Click()
Sheet3.Activate
End Sub

Private Sub cmdHelp_Click()
Sheet30.Visible = xlSheetVisible
Sheet30.Activate
'Sheet30.Range("i_general2").Select
End Sub


Private Sub cmdPrint_Click()
Module3.PrintWorksheets
End Sub

Private Sub cmdValidate_Click()
Module3.printerrormessage_IncD
End Sub

Private Sub CommandButton1_Click()
Module3.PrintWorksheets
End Sub

Private Sub CommandButton4_Click()
Module3.printerrormessage_IncD
End Sub

Private Sub cmdVallidate_Click()
Module3.printerrormessage_IncD
End Sub

' For Sheet : GENERAL2

Private Sub Worksheet_Change(ByVal Target As Range)

On Error GoTo exit1

Application.EnableEvents = False
If Target.Validation.Type = 3 Then
     GoTo exit1
End If

If (getRangeName(Target) = "Ver.PAN") Then
    Target.value = UCase(Target.value)
End If

Target.value = UCase(Target.value)
Target.Formula = UCase(Target.Formula)

If (Val(Sheet2.Range("IncD.RefundDue"))) > 0 Then
  If (Range("IncD.EcsRequired").value = "No") Then
      If (Not (Range("IncD.MICRCode").value = "")) Or (Not (Range("IncD.BankAccountType").value = "")) Then
            MsgBox "If refund is by cheque then MICR Code and Type of account must not be filled"
            Range("IncD.MICRCode").value = ""
            Range("IncD.BankAccountType").value = ""
       End If
  End If
End If
'Else
'    MsgBox "If no Refund due then Bank details not required "
'    Range("IncD.BankAccountNumber").Value = ""
'    Range("IncD.EcsRequired").Value = "No"
'    Range("IncD.MICRCode").Value = ""
'    Range("IncD.BankAccountType").Value = ""
'GoTo exit1
'End If


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
start = Len(Name)

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

