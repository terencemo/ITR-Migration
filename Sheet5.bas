Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Public Sub cmdInterest_Click()
On Error Resume Next
msginit21 = Module3.getmsgstate
Sheet5.Unprotect msginit21 + "*"
Module4.COMPUTE_INTEREST
Sheet5.Protect msginit21 + "*"
End Sub
Public Sub cmdInterestTransfer_Click()
On Error Resume Next
ThisComponent.Sheets.getByName("Sheet1").Range("IncD.IntrstPayUs234A").value = ThisComponent.Sheets.getByName("Sheet5").Range("Calc_234A").value
ThisComponent.Sheets.getByName("Sheet1").Range("IncD.IntrstPayUs234B").value = ThisComponent.Sheets.getByName("Sheet5").Range("Calc_234B").value
ThisComponent.Sheets.getByName("Sheet1").Range("IncD.IntrstPayUs234C").value = ThisComponent.Sheets.getByName("Sheet5").Range("Calc_234C").value
ThisComponent.Sheets.getByName("Sheet1").Range("IncD.TotalIntrstPay").value = ThisComponent.Sheets.getByName("Sheet5").Range("Calc_234A").value + _
    ThisComponent.Sheets.getByName("Sheet5").Range("Calc_234B").value + ThisComponent.Sheets.getByName("Sheet5").Range("Calc_234C").value
msginit21 = Module3.getmsgstate
Sheet1.Protect msginit21 + "*"
Sheet5.Protect msginit21 + "*"
End Sub
Public Sub cmdTax_Click()
On Error Resume Next
msginit21 = Module3.getmsgstate
Sheet5.Unprotect msginit21 + "*"
Module2.calc_TaxatNormalRate
Sheet5.Protect msginit21 + "*"
End Sub
Public Sub cmdTaxTransfer_Click()
On Error Resume Next
ThisComponent.Sheets.getByName("Sheet1").Range("IncD.TotalTaxPayable").value = ThisComponent.Sheets.getByName("Sheet5").Range("TXN_Calc").value
ThisComponent.Sheets.getByName("Sheet1").Range("IncD.RebateOnAgriInc").value = 0
ThisComponent.Sheets.getByName("Sheet1").Range("IncD.RebateOnAgriInc").value = 0
ThisComponent.Sheets.getByName("Sheet1").Range("IncD.SurchargeOnTaxPayable").value = 0
ThisComponent.Sheets.getByName("Sheet1").Range("IncD.EducationCess").value = ThisComponent.Sheets.getByName("Sheet5").Range("Calc_ED").value
End Sub
Private Sub cndBacj_Click()

End Sub
