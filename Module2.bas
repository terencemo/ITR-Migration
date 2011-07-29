Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Sub calc_TaxatNormalRate()

Dim ResStatus As String
Dim isSenior As Boolean
Dim gender As String
Dim dob As String
Dim Status As String
Const seniorDate As String = "01/04/1946"

Status = Mid(ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.Status").value, 1, 1)
ResStatus = Mid(ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.ResidentialStatus1").value, 1, 3)
gender = Mid(ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.Gender1").value, 1, 1)
dob = Dformat1(ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.DOB"), "yyyy-mm-dd")
If CheckDateMaxDDMMYYYY(dob, 31, 3, 1946) = True Then
    isSenior = True
Else
  isSenior = False
End If

'Sheet5.Range("Calc_SplRate").value = Round(Sheet17.Range("SI.TotSplRateIncTax2").value)

If Status = "H" Then

            'Sheet17.Range("THRESOLD").Value = Round(Sheet5.Range("HUF_TH").Value)
            ThisComponent.Sheets(5-1).getCellRangeByName("TXN_CALC").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("HUF").value)
            ThisComponent.Sheets(5-1).getCellRangeByName("avgratetax").value = ThisComponent.Sheets(5-1).getCellRangeByName("HUF_AVG").value
            ThisComponent.Sheets(5-1).getCellRangeByName("Rebate_AgriInc_Calc").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("HUF_rebate").value)
            ThisComponent.Sheets(5-1).getCellRangeByName("Calc_SplRate").value = ThisComponent.Sheets(17-1).getCellRangeByName("SI.TotSplRateIncTax").value
            ThisComponent.Sheets(5-1).getCellRangeByName("Sur_Calc").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("HUF_Surcharge").value)
            ThisComponent.Sheets(5-1).getCellRangeByName("Clac_MR").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("HUF_MR").value)
            ThisComponent.Sheets(5-1).getCellRangeByName("Calc_NetSur").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("HUF_NetSur").value)
            ThisComponent.Sheets(5-1).getCellRangeByName("Calc_ED").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("HUF_ED").value)
            ThisComponent.Sheets(5-1).getCellRangeByName("avgratetax").value = ThisComponent.Sheets(5-1).getCellRangeByName("HUF_AVG").value
            
ElseIf ResStatus = "RES" And isSenior Then
    'Sheet17.Range("THRESOLD").Value = Round(Sheet5.Range("RES_senior_TH").Value)
    ThisComponent.Sheets(5-1).getCellRangeByName("TXN_Calc").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("RES_senior").value)
    'Sheet5.Range("Calc_SplRate").Value = Sheet17.Range("SI.TotSplRateIncTax").Value
    ThisComponent.Sheets(5-1).getCellRangeByName("Rebate_AgriInc_Calc").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("RES_senior_rebate").value)
    'Sheet5.Range("Calc_SplRate").Value = Sheet17.Range("SI.TotSplRateIncTax").Value
    ThisComponent.Sheets(5-1).getCellRangeByName("Sur_Calc").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("RES_senior_Surcharge").value)
    ThisComponent.Sheets(5-1).getCellRangeByName("Clac_MR").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("RES_senior_MR").value)
    ThisComponent.Sheets(5-1).getCellRangeByName("Calc_NetSur").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("RES_senior_NetSur").value)
    ThisComponent.Sheets(5-1).getCellRangeByName("Calc_ED").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("RES_senior_ED").value)
    
    ThisComponent.Sheets(5-1).getCellRangeByName("avgratetax").value = ThisComponent.Sheets(5-1).getCellRangeByName("RES_senior_AVG").value
  ElseIf ResStatus = "RES" And gender = "F" Then
        'Sheet17.Range("THRESOLD").Value = Round(Sheet5.Range("Res_F_TH").Value)
        ThisComponent.Sheets(5-1).getCellRangeByName("TXN_Calc").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("Res_F_TXN").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("Rebate_AgriInc_Calc").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("Res_F_rebate").value)
        'Sheet5.Range("Calc_SplRate").Value = Sheet17.Range("SI.TotSplRateIncTax").Value
        ThisComponent.Sheets(5-1).getCellRangeByName("Sur_Calc").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("Res_F_Surcharge").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("Clac_MR").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("Res_F_MR").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("Calc_NetSur").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("Res_F_NetSur").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("Calc_ED").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("Res_F_ED").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("avgratetax").value = ThisComponent.Sheets(5-1).getCellRangeByName("Res_F_AVG").value
    ElseIf ResStatus = "RES" Then
        'Sheet17.Range("THRESOLD").Value = Round(Sheet5.Range("Res_M_TH").Value)
        ThisComponent.Sheets(5-1).getCellRangeByName("TXN_Calc").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("Res_M_TXN").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("Rebate_AgriInc_Calc").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("Res_M_rebate").value)
        'Sheet5.Range("Calc_SplRate").Value = Sheet17.Range("SI.TotSplRateIncTax").Value
        ThisComponent.Sheets(5-1).getCellRangeByName("Sur_Calc").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("Res_M_Surcharge").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("Clac_MR").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("Res_M_MR").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("Calc_NetSur").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("Res_M_NetSur").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("Calc_ED").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("Res_M_ED").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("avgratetax").value = ThisComponent.Sheets(5-1).getCellRangeByName("Res_M_AVG").value
    ElseIf ResStatus = "NRI" Then
        'Sheet17.Range("THRESOLD").Value = Round(Sheet5.Range("NRI_TH").Value)
        ThisComponent.Sheets(5-1).getCellRangeByName("TXN_Calc").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("NRI").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("Rebate_AgriInc_Calc").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("NRI_rebate").value)
        'Sheet5.Range("Calc_SplRate").Value = Sheet17.Range("SI.TotSplRateIncTax").Value
        ThisComponent.Sheets(5-1).getCellRangeByName("Sur_Calc").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("NRI_Surcharge").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("Clac_MR").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("NRI_MR").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("Calc_NetSur").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("NRI_NetSur").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("Calc_ED").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("NRI_ED").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("avgratetax").value = ThisComponent.Sheets(5-1).getCellRangeByName("NRI_AVG").value
    ElseIf ResStatus = "NOR" Then
        'Sheet17.Range("THRESOLD").Value = Round(Sheet5.Range("NRI_TH").Value)
        ThisComponent.Sheets(5-1).getCellRangeByName("TXN_Calc").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("NRI").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("Rebate_AgriInc_Calc").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("NRI_rebate").value)
        'Sheet5.Range("Calc_SplRate").Value = Sheet17.Range("SI.TotSplRateIncTax").Value
        ThisComponent.Sheets(5-1).getCellRangeByName("Sur_Calc").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("NRI_Surcharge").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("Clac_MR").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("NRI_MR").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("Calc_NetSur").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("NRI_NetSur").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("Calc_ED").value = Round(ThisComponent.Sheets(5-1).getCellRangeByName("NRI_ED").value)
        ThisComponent.Sheets(5-1).getCellRangeByName("avgratetax").value = ThisComponent.Sheets(5-1).getCellRangeByName("NRI_AVG").value
End If


End Sub

Function CheckDateMaxDDMMYYYY(ByVal dt As String, ByVal maxday As Integer, ByVal maxmonth As Integer, maxyear As Integer) As Boolean
CheckDateMaxDDMMYYYY = True
If (Val(Mid(dt, 1, 4)) > maxyear) Then
    CheckDateMaxDDMMYYYY = False
    'MsgBox "INVALID DATE, " & errormsg
    GoTo exit1
End If
If (Val(Mid(dt, 1, 4)) = maxyear) And (Val(Mid(dt, 6, 2)) > maxmonth) Then
          CheckDateMaxDDMMYYYY = False
          'MsgBox "INVALID DATE, " & errormsg
    GoTo exit1
End If

If (Val(Mid(dt, 1, 4)) = maxyear) And (Val(Mid(dt, 6, 2)) = maxmonth) And (Val(Mid(dt, 9, 2)) > maxday) Then
          CheckDateMaxDDMMYYYY = False
          'MsgBox "INVALID DATE, " & errormsg
    GoTo exit1
End If

exit1:
End Function
Function Dformat1(dt As Variant, timepass As String) As String
'yyyy-mm-dd'

Dim formateddate As String
Dim day As String
Dim month As String
Dim year As String
If Len(dt) > 0 Then

year = Mid(dt, 7, 4)
month = Mid(dt, 4, 2)
day = Mid(dt, 1, 2)
formateddate = year & "-" & month & "-" & day
Dformat1 = formateddate

Else
Dformat1 = ""
End If
End Function
