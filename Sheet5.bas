Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Public Sub cmdInterest_Click()
On Error Resume Next
msginit21 = Module3.getmsgstate
ThisComponent.CurrentController.getActiveSheet.Unprotect msginit21 + "*"
Module4.COMPUTE_INTEREST
ThisComponent.CurrentController.getActiveSheet.Protect msginit21 + "*"
End Sub
Public Sub cmdInterestTransfer_Click()
On Error Resume Next
ThisComponent.Sheets(1-1).getCellRangeByName("IncD.IntrstPayUs234A").value = ThisComponent.Sheets(5-1).getCellRangeByName("Calc_234A").value
ThisComponent.Sheets(1-1).getCellRangeByName("IncD.IntrstPayUs234B").value = ThisComponent.Sheets(5-1).getCellRangeByName("Calc_234B").value
ThisComponent.Sheets(1-1).getCellRangeByName("IncD.IntrstPayUs234C").value = ThisComponent.Sheets(5-1).getCellRangeByName("Calc_234C").value
ThisComponent.Sheets(1-1).getCellRangeByName("IncD.TotalIntrstPay").value = ThisComponent.Sheets(5-1).getCellRangeByName("Calc_234A").value + _
    ThisComponent.Sheets(5-1).getCellRangeByName("Calc_234B").value + ThisComponent.Sheets(5-1).getCellRangeByName("Calc_234C").value
msginit21 = Module3.getmsgstate
ThisComponent.Sheets(0).Protect msginit21 + "*"
ThisComponent.CurrentController.getActiveSheet.Protect msginit21 + "*"
End Sub
Public Sub cmdTax_Click()
On Error Resume Next
msginit21 = Module3.getmsgstate
ThisComponent.CurrentController.getActiveSheet.Unprotect msginit21 + "*"
Module2.calc_TaxatNormalRate
ThisComponent.CurrentController.getActiveSheet.Protect msginit21 + "*"
End Sub
Public Sub cmdTaxTransfer_Click()
On Error Resume Next
ThisComponent.Sheets(1-1).getCellRangeByName("IncD.TotalTaxPayable").value = ThisComponent.Sheets(5-1).getCellRangeByName("TXN_Calc").value
ThisComponent.Sheets(1-1).getCellRangeByName("IncD.RebateOnAgriInc").value = 0
ThisComponent.Sheets(1-1).getCellRangeByName("IncD.RebateOnAgriInc").value = 0
ThisComponent.Sheets(1-1).getCellRangeByName("IncD.SurchargeOnTaxPayable").value = 0
ThisComponent.Sheets(1-1).getCellRangeByName("IncD.EducationCess").value = ThisComponent.Sheets(5-1).getCellRangeByName("Calc_ED").value
End Sub
Private Sub cndBacj_Click()

End Sub
