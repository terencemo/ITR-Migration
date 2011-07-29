Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Dim i, count As Integer
Dim s As String
Dim fff As String

'Start dim
Dim msginit As String
Dim msginit21 As String
Dim FirstName_1 As Variant
Dim MiddleName_1 As Variant
Dim SurNameOrOrgName_1 As Variant
Dim PAN_1 As Variant
Dim ResidenceNo_1 As Variant
Dim ResidenceName_1 As Variant
Dim RoadOrStreet_1 As Variant
Dim LocalityOrArea_1 As Variant
Dim CityOrTownOrDistrict_1 As Variant
Dim StateCode_1 As Variant
Dim PinCode_1 As Variant
Dim STDcode_1 As Variant
Dim PhoneNo_1 As Variant
Dim MobileNo_1 As Variant
Dim EmailAddress_1 As Variant
Dim DOB_1 As Variant
Dim EmployerCategory_1 As Variant
Dim Gender_1 As Variant
Dim DesigOfficerWardorCircle_1 As Variant
Dim ReturnFileSec_1 As Variant
Dim ReturnType_1 As Variant
Dim ReceiptNo_1 As Variant
Dim OrigRetFiledDate_1 As Variant
Dim ResidentialStatus_1 As Variant
Dim AsseseeRepFlg_1 As Variant
Dim RepName_1 As Variant
Dim RepAddress_1 As Variant
Dim RepPAN_1 As Variant
Dim LiableSec44AAflg_1 As Variant
Dim LiableSec44ABflg_1 As Variant
Dim AuditorName_1 As Variant
Dim AuditorMemNo_1 As Variant
Dim AudFrmName_1 As Variant
Dim AudFrmPAN_1 As Variant
Dim AuditDate_1 As Variant
Dim msgValidateSheet1 As String
Dim msgValidateSheet1Blanks As String
Dim msgValidateSheet1specialcharacters As String
Dim msgstats As String
Dim msgstate As String

Dim status_1 As Variant
Dim verPAN As Variant

 'Variable Declaration
Dim IncomeFromSal_IncD As Variant
Dim IncomeFromHP_IncD As Variant
Dim FamPension_IncD As Variant
Dim IndInterest_IncD As Variant
Dim IncomeFromOS_IncD As Variant
Dim GrossTotIncome_IncD As Variant
Dim Section80C_IncD As Variant
Dim Section80CCC_IncD As Variant
Dim Section80CCD_IncD As Variant
Dim Section80CCF_IncD As Variant
Dim Section80D_IncD As Variant
Dim Section80DD_IncD As Variant
Dim Section80DDB_IncD As Variant
Dim Section80E_IncD As Variant
Dim Section80G_IncD As Variant
Dim Section80GG_IncD As Variant
Dim Section80GGA_IncD As Variant
Dim Section80GGC_IncD As Variant
Dim Section80U_IncD As Variant
Dim TotalChapVIADeductions_IncD As Variant
Dim TotalIncome_IncD As Variant
Dim NetAgriculturalIncome_IncD As Variant
Dim AggregateIncome_IncD As Variant
Dim TaxOnAggregateInc_IncD As Variant
Dim RebateOnAgriInc_IncD As Variant
Dim TotalTaxPayable_IncD As Variant
Dim SurchargeOnTaxPayable_IncD As Variant
Dim EducationCess_IncD As Variant
Dim GrossTaxLiability_IncD As Variant
Dim Section89_IncD As Variant
Dim Section90and91_IncD As Variant
Dim NetTaxLiability_IncD As Variant
Dim IntrstPayUs234A_IncD As Variant
Dim IntrstPayUs234B_IncD As Variant
Dim IntrstPayUs234C_IncD As Variant
Dim TotalIntrstPay_IncD As Variant
Dim TotTaxPlusIntrstPay_IncD As Variant
Dim AdvanceTax_IncD As Variant
Dim TDS_IncD As Variant
Dim SelfAssessmentTax_IncD As Variant
Dim TotalTaxesPaid_IncD As Variant
Dim BalTaxPayable_IncD As Variant
Dim RefundDue_IncD As Variant
Dim BankAccountNumber_IncD As Variant
Dim EcsRequired_IncD As Variant
Dim MICRCode_IncD As Variant
Dim BankAccountType_IncD As Variant
Dim TAN_TDSal As Variant
Dim EmployerOrDeductorOrCollecterName_TDSal As Variant
Dim AddrDetail_TDSal As Variant
Dim CityOrTownOrDistrict_TDSal As Variant
Dim StateCode_TDSal As Variant
Dim PinCode_TDSal As Variant
Dim IncChrgSal_TDSal As Variant
Dim DeductUnderChapVIA_TDSal As Variant
Dim TaxPayIncluSurchEdnCes_TDSal As Variant
Dim TotalTDSSal_TDSal As Variant
Dim TaxPayRefund_TDSal As Variant
Dim TAN_TDSoth As Variant
Dim EmployerOrDeductorOrCollecterName_TDSoth As Variant
Dim AddrDetail_TDSoth As Variant
Dim CityOrTownOrDistrict_TDSoth As Variant
Dim StateCode_TDSoth As Variant
Dim PinCode_TDSoth As Variant
Dim AmtPaid_TDSoth As Variant
Dim DatePayCred_TDSoth As Variant
Dim TotTDSOnAmtPaid_TDSoth As Variant
Dim ClaimOutOfTotTDSOnAmtPaid_TDSoth As Variant
Dim NameOfBank_TaxP As Variant
Dim NameOfBranch_TaxP As Variant
Dim BSRCode_TaxP As Variant
Dim DateDep_TaxP As Variant
Dim SrlNoOfChaln_TaxP As Variant
Dim Amt_TaxP As Variant
Dim Code001_AIR As Variant
Dim Code002_AIR As Variant
Dim Code003_AIR As Variant
Dim Code004_AIR As Variant
Dim Code005_AIR As Variant
Dim Code006_AIR As Variant
Dim Code007_AIR As Variant
Dim Code008_AIR As Variant
Dim TaxExmpIntInc_AIR As Variant
Dim AssesseeVerName_Ver As Variant
Dim FatherName_Ver As Variant
Dim Place_Ver As Variant
Dim Date_Ver As Variant
Dim IdentificationNoOfTRP_Ver As Variant
Dim NameOfTRP_Ver As Variant
Dim ReImbFrmGov_Ver As Variant

 'Variable and Function Declaration for TableType=T1 and RptFrm

Dim msgValidateSheetIncD As String


Dim rngname_TDSal As Variant
Dim end_TDSal As Variant
Dim incBy_TDSal As Variant
Dim msgValidateSheetTDSal As String

Dim rngname_TDSal2 As Variant
Dim end_TDSal2 As Variant

Dim rngname_TDSoth As Variant
Dim end_TDSoth As Variant
Dim incBy_TDSoth As Variant
Dim msgValidateSheetTDSoth As String

Dim rngname_TDSoth2 As Variant
Dim end_TDSoth2 As Variant

Dim rngname_TaxP As Variant
Dim end_TaxP As Variant
Dim incBy_TaxP As Variant
Dim msgValidateSheetTaxP As String

Dim rngname_TaxP2 As Variant
Dim end_TaxP2 As Variant


Dim msgValidateSheetAIR As String
Const msgstats1 As String = "BE"

Dim msgValidateSheetVer As String

Dim Total_Count As Integer

Dim secapplicable() As Variant
Dim secapplicable1() As Variant
Dim sectionname() As Variant


Function getmsgstate() As String

'msgstate = Sheet6.Range(msgstats1 + "1")
msgstate = ""
msgstate = msgstate + "(21)"
getmsgstate = msgstate
End Function
Sub AddRows_TDSal()

 setTblinfo_TDSal
 Dim numberofrows As Integer
 SelectLastRow ("TDSal.TAN")
 numberofrows = InsertRowsAndFillFormulas()
 Call ExendRangeNameToTable(numberofrows, rngname_TDSal)
End Sub


Sub setTblinfo_TDSal()
 Dim rangecells As Object
 Dim mIntCells As Integer
 Dim mIntCtr As Integer
 Dim ccount As Integer
 ccount = 0
 Set rangecells = ThisComponent.Sheets(1).getCellRangeByName("TDSal.TAN_1")
 mIntCells = rangecells.ComputeFunction(com.sun.star.sheet.GeneralFunction.COUNT)
 For mIntCtr = 0 To mIntCells - 1
     If rangecells.getCellByPosition(0, mIntCtr).Type = EMPTY Then
         ccount = ccount + 1
     End If
 Next
 end_TDSal = ccount
 rngname_TDSal = "TDSal.TAN;TDSal.EmployerOrDeductorOrCollecterName;TDSal.IncChrgSal;TDSal.TotalTDSSal;"
 End Sub
Sub setTblinfo_TDSal2()
 Dim rangecells As Object
 Dim mIntCells As Integer
 Dim mIntCtr As Integer
 Dim ccount As Integer
 ccount = 0
 Set rangecells = ThisComponent.Sheets(1).getCellRangeByName("TDSal.TotalTDSSal")
 mIntCells = rangecells.ComputeFunction(com.sun.star.sheet.GeneralFunction.COUNT)
 For mIntCtr = 0 To mIntCells - 1
     If rangecells.getCellByPosition(0, mIntCtr).Type = EMPTY Then
         ccount = ccount + 1
     End If
 Next
 end_TDSal2 = ccount
 rngname_TDSal2 = "TDSal.TAN;TDSal.EmployerOrDeductorOrCollecterName;TDSal.IncChrgSal;TDSal.TotalTDSSal;"
 End Sub


Sub AddRows_TDSoth()
 setTblinfo_TDSoth
 Dim numberofrows As Integer
 SelectLastRow ("TDSoth.TAN")
 numberofrows = InsertRowsAndFillFormulas()
 Call ExendRangeNameToTable(numberofrows, rngname_TDSoth)
End Sub


Sub setTblinfo_TDSoth()
 Dim rangecells As Object
 Dim mIntCells As Integer
 Dim mIntCtr As Integer
 Dim ccount As Integer
 ccount = 0
 Set rangecells = ThisComponent.Sheets(1).getCellRangeByName("TDSoth.TAN")
 mIntCells = rangecells.ComputeFunction(com.sun.star.sheet.GeneralFunction.COUNT)
 For mIntCtr = 0 To mIntCells - 1
     If rangecells.getCellByPosition(0, mIntCtr).Type = EMPTY Then
         ccount = ccount + 1
     End If
 Next
 end_TDSoth = ccount
 rngname_TDSoth = "TDSoth.TAN;TDSoth.EmployerOrDeductorOrCollecterName;TDSoth.TotTDSOnAmtPaid;TDSoth.ClaimOutOfTotTDSOnAmtPaid;"
 End Sub

Sub setTblinfo_TDSoth2()
 Dim rangecells As Object
 Dim mIntCells As Integer
 Dim mIntCtr As Integer
 Dim ccount As Integer
 ccount = 0
 Set rangecells = ThisComponent.Sheets(1).getCellRangeByName("TDSoth.ClaimOutOfTotTDSOnAmtPaid")
 mIntCells = rangecells.ComputeFunction(com.sun.star.sheet.GeneralFunction.COUNT)
 For mIntCtr = 0 To mIntCells - 1
     If rangecells.getCellByPosition(0, mIntCtr).Type = EMPTY Then
         ccount = ccount + 1
     End If
 Next
 end_TDSoth2 = ccount
 rngname_TDSoth2 = "TDSoth.TAN;TDSoth.EmployerOrDeductorOrCollecterName;TDSoth.TotTDSOnAmtPaid;TDSoth.ClaimOutOfTotTDSOnAmtPaid;"
 End Sub

Sub AddBlockCall_salrptfrm()
     'setTblinfo_salrptfrm
     Call addblock(rngname_salrptfrm, frmRngname_salrptfrm, cntrRng_salrptfrm, frmsize_salrptfrm)
End Sub
Function cmdFileDialog() As String
cmdFileDialog = ""
   Dim fDialog As Office.FileDialog
   Dim varFile As Variant

     Set fDialog = Application.FileDialog(msoFileDialogFilePicker)

   With fDialog

      .AllowMultiSelect = False
      .Filters.Clear
      .Filters.add "Microsoft Office Excel Workbook", "*.xls"
        If .Show = True Then
            For Each varFile In .SelectedItems
               cmdFileDialog = varFile
            Next
        End If
   End With
End Function

Function InsertRowsToImport(Optional vRows As Integer = 0)
   
   Dim x As Long
   msginit21 = Module3.getmsgstate

   strpassword = msginit21 + "*"
   ThisComponent.CurrentController.getActiveSheet.Unprotect Password:=strpassword

    
   ActiveCell.EntireRow.Select  'So you do not have to preselect entire row
    
   Dim sht As Worksheet, shts() As String, i As Integer
   ReDim shts(1 To Worksheets.Application.ActiveWorkbook. _
       Windows(1).SelectedSheets.count)
   i = 0
   For Each sht In _
       Application.ActiveWorkbook.Windows(1).SelectedSheets
    Sheets(sht.Name).Select
    i = i + 1
    shts(i) = sht.Name

    x = Sheets(sht.Name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow. _
     Resize(rowsize:=vRows).Insert Shift:=xlDown

    Selection.AutoFill Selection.Resize( _
     rowsize:=vRows + 1), xlFillDefault

    On Error Resume Next
    
Selection.Offset(1).Resize(vRows).EntireRow. _
 SpecialCells(xlCellTypeAllValidation).ClearContents
'
   Next sht
ThisComponent.CurrentController.getActiveSheet.Protect Password:=strpassword
   
End Function


Sub Import()

Dim flag As Boolean
Dim Filename As Variant
Dim dfilename, ndfilename As Variant
Dim add As Variant
Dim destadd As Variant

Dim a As Integer
Dim AB As Integer


flag = True
On Error Resume Next

MsgBox "Please Check the final ITR1 after importing to ensure all rows are inserted. "
'MsgBox "Nature of Business, enter Code manually after importing. Especially, check S.80G after importing to ensure correctness."

Filename = cmdFileDialog()
If Not Filename = "" Then
        Filename = Split(Filename, "\")
        newfilename = Filename(UBound(Filename))
        Dim DestBook As Workbook, SrcBook As Workbook
        Dim cnt As Integer
        cnt = 0
    Application.ScreenUpdating = False
    Set SrcBook = Workbooks.Open(newfilename)
    Set DestBook = ThisWorkbook
   dfilename = Split((DestBook.FullName), "\")
   ndfilename = dfilename(UBound(dfilename))
   
   If newfilename <> ndfilename Then

    For Each RNAME In Workbooks(newfilename).Names
    
        If RNAME.Name = "TDSal.TAN" Then
            sfirstbound = SrcBook.Sheets("TDS").getCellRangeByName(RNAME.Name).Address
            sTEMP = Split(sfirstbound, "$")
            supperbound = UBound(sTEMP)
            sTEMP = sTEMP(UBound(sTEMP))

            dfirstbound = DestBook.Sheets("TDS").getCellRangeByName(RNAME.Name).Address
            dTemp = Split(dfirstbound, "$")
            dupperbound = UBound(dTemp)
            ddTemp = dTemp(UBound(dTemp))

            cnt = SrcBook.Sheets("TDS").getCellRangeByName(RNAME.Name).ComputeFunction(com.sun.star.sheet.GeneralFunction.COUNT)
            dcnt = DestBook.Sheets("TDS").getCellRangeByName(RNAME.Name).ComputeFunction(com.sun.star.sheet.GeneralFunction.COUNT)
            DestBook.Sheets("TDS").Activate
                If (cnt - dcnt) > 0 Then
                    DestBook.Sheets("TDS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                    setTblinfo_TDSal
                    InsertRowsToImport (cnt - dcnt)
                    Call ExendRangeNameToTable(cnt - dcnt, rngname_TDSal)
                    SrcBook.Sheets("TDS").Range(RNAME.Name).Copy
                    DestBook.Sheets("TDS").Range(RNAME.Name).PasteSpecial xlValues
                Else
                    SrcBook.Sheets("TDS").Range(RNAME.Name).Copy
                    DestBook.Sheets("TDS").Range(RNAME.Name).PasteSpecial xlValues
                End If
         End If
         
         
         If RNAME.Name = "TDSoth.TAN" Then

            sfirstbound = SrcBook.Sheets("TDS").Range(RNAME.Name).Address
            sTEMP = Split(sfirstbound, "$")
            supperbound = UBound(sTEMP)
            sTEMP = sTEMP(UBound(sTEMP))

            dfirstbound = DestBook.Sheets("TDS").Range(RNAME.Name).Address
            dTemp = Split(dfirstbound, "$")
            dupperbound = UBound(dTemp)
            ddTemp = dTemp(UBound(dTemp))

            cnt = SrcBook.Sheets("TDS").getCellRangeByName(RNAME.Name).ComputeFunction(com.sun.star.sheet.GeneralFunction.COUNT)
            dcnt = DestBook.Sheets("TDS").getCellRangeByName(RNAME.Name).ComputeFunction(com.sun.star.sheet.GeneralFunction.COUNT)
            DestBook.Sheets("TDS").Activate
                If (cnt - dcnt) > 0 Then
                    DestBook.Sheets("TDS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                    setTblinfo_TDSoth
                    InsertRowsToImport (cnt - dcnt)
                    Call ExendRangeNameToTable(cnt - dcnt, rngname_TDSoth)
                    SrcBook.Sheets("TDS").Range(RNAME.Name).Copy
                    DestBook.Sheets("TDS").Range(RNAME.Name).PasteSpecial xlValues
                Else
                    SrcBook.Sheets("TDS").Range(RNAME.Name).Copy
                    DestBook.Sheets("TDS").Range(RNAME.Name).PasteSpecial xlValues
                End If
         End If
         
         If RNAME.Name = "TaxP.BSRCode" Then

            sfirstbound = SrcBook.Sheets("TDS").Range(RNAME.Name).Address
            sTEMP = Split(sfirstbound, "$")
            supperbound = UBound(sTEMP)
            sTEMP = sTEMP(UBound(sTEMP))

            dfirstbound = DestBook.Sheets("TDS").Range(RNAME.Name).Address
            dTemp = Split(dfirstbound, "$")
            dupperbound = UBound(dTemp)
            ddTemp = dTemp(UBound(dTemp))

            cnt = SrcBook.Sheets("TDS").getCellRangeByName(RNAME.Name).ComputeFunction(com.sun.star.sheet.GeneralFunction.COUNT)
            dcnt = DestBook.Sheets("TDS").getCellRangeByName(RNAME.Name).ComputeFunction(com.sun.star.sheet.GeneralFunction.COUNT)
            DestBook.Sheets("TDS").Activate
                If (cnt - dcnt) > 0 Then
                    DestBook.Sheets("TDS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                    setTblinfo_TaxP
                    InsertRowsToImport (cnt - dcnt)
                    Call ExendRangeNameToTable(cnt - dcnt, rngname_TaxP)
                    SrcBook.Sheets("TDS").Range(RNAME.Name).Copy
                    DestBook.Sheets("TDS").Range(RNAME.Name).PasteSpecial xlValues
                Else
                    SrcBook.Sheets("TDS").Range(RNAME.Name).Copy
                    DestBook.Sheets("TDS").Range(RNAME.Name).PasteSpecial xlValues
                End If
         End If

    Next
      a = 1
      AB = 1
    ReDim add(SrcBook.Worksheets.count)
    ReDim destadd(SrcBook.Worksheets.count)
    
    For Each ws In SrcBook.Worksheets
        add(a) = ws.Name
        a = a + 1
    Next
    For Each ws In DestBook.Worksheets
        destadd(AB) = ws.Name
        AB = AB + 1
    Next
    
    Application.EnableEvents = False
    Dim NEWRNAME As Variant
    Dim NEWSHEETNAME As Variant
    For Each RNAME In Workbooks(newfilename).Names
    NEWRNAME = RNAME.Name
    NEWSHEETNAME = ""
    If InStr(1, NEWRNAME, "!") > 0 Then
    NEWRNAME = Mid(RNAME.Name, InStr(1, RNAME.Name, "!") + 1)
    NEWSHEETNAME = Mid(RNAME.Name, 1, InStr(1, RNAME.Name, "!") - 1)
    If UCase(NEWSHEETNAME) = "GENERAL" Then
        NEWSHEETNAME = "Income Details"
    End If
    If UCase(NEWSHEETNAME) = "GENERAL2" Then
         NEWSHEETNAME = "Taxes paid and Verification"
    End If
    'NEWRNAME = NEWSHEETNAME & NEWRNAME
    End If
    
       
        For a = 1 To UBound(add)
                    DestBook.Worksheets(destadd(a)).Activate
                    SrcBook.Worksheets(add(a)).Range(RNAME.Name).Copy
                    DestBook.Worksheets(destadd(a)).Range(NEWRNAME).PasteSpecial xlValues
        Next
    Next

    Application.EnableEvents = True
      MsgBox "Import Completed"
      DestBook.Save
      DestBook.Worksheets("Income Details").Select
      Set SrcBook = Nothing
    'On Error GoTo 0
Else
   MsgBox "Source File must Not Have Same Name As Destination File"
End If
End If
End Sub


Sub AddRows_TaxP()
 setTblinfo_TaxP
 Dim numberofrows As Integer
 SelectLastRow ("TaxP.BSRCode")
 numberofrows = InsertRowsAndFillFormulas()
 Call ExendRangeNameToTable(numberofrows, rngname_TaxP)
End Sub


Sub setTblinfo_TaxP()
 Dim rangecells As Object
 Dim mIntCells As Integer
 Dim mIntCtr As Integer
 Dim ccount As Integer
 ccount = 0
 Set rangecells = ThisComponent.Sheets(1).getCellRangeByName("TaxP.BSRCode")
 mIntCells = rangecells.ComputeFunction(com.sun.star.sheet.GeneralFunction.COUNT)
 For mIntCtr = 0 To mIntCells - 1
     If rangecells.getCellByPosition(0, mIntCtr).Type = EMPTY Then
         ccount = ccount + 1
     End If
 Next
 end_TaxP = ccount
 rngname_TaxP = "TaxP.BSRCode;TaxP.DateDep;TaxP.SrlNoOfChaln;TaxP.Amt;IT.FormulaOFS;"
 End Sub

Sub setTblinfo_TaxP2()
 Dim rangecells As Object
 Dim mIntCells As Integer
 Dim mIntCtr As Integer
 Dim ccount As Integer
 ccount = 0
 Set rangecells = ThisComponent.Sheets(1).getCellRangeByName("TaxP.Amt")
 mIntCells = rangecells.ComputeFunction(com.sun.star.sheet.GeneralFunction.COUNT)
 For mIntCtr = 0 To mIntCells - 1
     If rangecells.getCellByPosition(0, mIntCtr).Type = EMPTY Then
         ccount = ccount + 1
     End If
 Next
 end_TaxP2 = ccount
 rngname_TaxP2 = "TaxP.BSRCode;TaxP.DateDep;TaxP.SrlNoOfChaln;TaxP.Amt;IT.FormulaOFS;"
 End Sub


'---ValidateBlock
 Sub Validateshts()

If Not Validatesheet1 Then
    ThisComponent.CurrentController.setActiveSheet(ThisComponent.Sheets(1-1))
    MsgBox (msgValidateSheet1)
    EndProcessing
End If

 If Not ValidatesheetIncD Then
     ThisComponent.CurrentController.setActiveSheet(ThisComponent.Sheets(2-1))
     MsgBox (msgValidateSheetIncD)
     EndProcessing
 End If


 If Not ValidatesheetTDSal Then
     ThisComponent.CurrentController.setActiveSheet(ThisComponent.Sheets(3-1))
     MsgBox (msgValidateSheetTDSal)
     EndProcessing
 End If


 If Not ValidatesheetTDSoth Then
     ThisComponent.CurrentController.setActiveSheet(ThisComponent.Sheets(3-1))
     MsgBox (msgValidateSheetTDSoth)
     EndProcessing
 End If


 If Not ValidatesheetTaxP Then
     ThisComponent.CurrentController.setActiveSheet(ThisComponent.Sheets(3-1))
     MsgBox (msgValidateSheetTaxP)
     EndProcessing
 End If


 If Not ValidatesheetAIR Then
     ThisComponent.CurrentController.setActiveSheet(ThisComponent.Sheets(2-1))
     MsgBox (msgValidateSheetAIR)
     EndProcessing
 End If


 If Not ValidatesheetVer Then
     ThisComponent.CurrentController.setActiveSheet(ThisComponent.Sheets(2-1))
     MsgBox (msgValidateSheetVer)
     EndProcessing
 End If


 End Sub
'------ End ValidateBlock
'---Printerrormessage Functions

 Sub printerrormessage_IncD()
 If Not ValidatesheetIncD Then
     ThisComponent.CurrentController.setActiveSheet(ThisComponent.Sheets(2-1))
     MsgBox (msgValidateSheetIncD)
     EndProcessing
'Else
'MsgBox (" Sheet is ok ")
 End If
 
  printerrormessage_AIR
 printerrormessage_Ver
 End Sub


 Sub printerrormessage_TDSal()
 If Not ValidatesheetTDSal Then
     ThisComponent.CurrentController.setActiveSheet(ThisComponent.Sheets(3-1))
     MsgBox (msgValidateSheetTDSal)
     EndProcessing
'Else
'MsgBox (" Sheet is ok ")
 End If
 printerrormessage_TDSoth
 printerrormessage_TaxP
 End Sub


 Sub printerrormessage_TDSoth()
 If Not ValidatesheetTDSoth Then
     ThisComponent.CurrentController.setActiveSheet(ThisComponent.Sheets(3-1))
     MsgBox (msgValidateSheetTDSoth)
     EndProcessing
'Else
'MsgBox (" Sheet is ok ")
 End If
 End Sub


 Sub printerrormessage_TaxP()
 If Not ValidatesheetTaxP Then
     ThisComponent.CurrentController.setActiveSheet(ThisComponent.Sheets(3-1))
     MsgBox (msgValidateSheetTaxP)
     EndProcessing
Else
MsgBox (" Sheet is ok ")
 End If
 End Sub


 Sub printerrormessage_AIR()
 If Not ValidatesheetAIR Then
     ThisComponent.CurrentController.setActiveSheet(ThisComponent.Sheets(2-1))
     MsgBox (msgValidateSheetAIR)
     EndProcessing
'Else
'MsgBox (" Sheet is ok ")
 End If
 End Sub


 Sub printerrormessage_Ver()
 If Not ValidatesheetVer Then
     ThisComponent.CurrentController.setActiveSheet(ThisComponent.Sheets(2-1))
     MsgBox (msgValidateSheetVer)
     EndProcessing
Else
MsgBox (" Sheet is ok ")
 End If
 End Sub

'------ End Printerrormessage Functions

Function msgbox_IncD(strmsg As String) As String
     msgValidateSheetIncD = msgValidateSheetIncD & strmsg & Chr(13)
End Function


Function msgbox_TDSal(strmsg As String) As String
     msgValidateSheetTDSal = msgValidateSheetTDSal & strmsg & Chr(13)
End Function


Function msgbox_TDSoth(strmsg As String) As String
     msgValidateSheetTDSoth = msgValidateSheetTDSoth & strmsg & Chr(13)
End Function


Function msgbox_TaxP(strmsg As String) As String
     msgValidateSheetTaxP = msgValidateSheetTaxP & strmsg & Chr(13)
End Function


Function msgbox_AIR(strmsg As String) As String
     msgValidateSheetAIR = msgValidateSheetAIR & strmsg & Chr(13)
End Function


Function msgbox_Ver(strmsg As String) As String
     msgValidateSheetVer = msgValidateSheetVer & strmsg & Chr(13)
End Function

'Main Validation Function

Function ValidatesheetIncD() As Boolean
     ValidatesheetIncD = True

'If (Val(Sheet2.Range("IncD.TotalTaxesPaid")) > 0) Then
     If Not ValidateAdvanceTax_IncD() Then ValidatesheetIncD = False
     If Not ValidateTDS_IncD() Then ValidatesheetIncD = False
     If Not ValidateSelfAssessmentTax_IncD() Then ValidatesheetIncD = False
     If Not ValidateTotalTaxesPaid_IncD() Then ValidatesheetIncD = False
     If Not ValidateBalTaxPayable_IncD() Then ValidatesheetIncD = False
'End If
If Not ValidateRefundDue_IncD() Then ValidatesheetIncD = False
If (ThisComponent.Sheets(2-1).getCellRangeByName("IncD.RefundDue").Value > 0) Then
     If Not ValidateBankAccountNumber_IncD() Then ValidatesheetIncD = False
     If Not ValidateEcsRequired_IncD() Then ValidatesheetIncD = False
     If ThisComponent.Sheets(2-1).getCellRangeByName("IncD.EcsRequired") = "Yes" Then
        If Not ValidateMICRCode_IncD() Then ValidatesheetIncD = False
        If Not ValidateBankAccountType_IncD() Then ValidatesheetIncD = False
     End If
 End If
End Function


Function ValidatesheetTDSal() As Boolean
     ValidatesheetTDSal = True
     If Not ValidateTAN_TDSal() Then ValidatesheetTDSal = False
 If (Len(ThisComponent.Sheets(1).getCellRangeByName("TDSal.TAN").getCellByPosition(0, 1).String) > 0) Then
     If Not ValidateEmployerOrDeductorOrCollecterName_TDSal() Then ValidatesheetTDSal = False
     If Not ValidateAddrDetail_TDSal() Then ValidatesheetTDSal = False
     If Not ValidateCityOrTownOrDistrict_TDSal() Then ValidatesheetTDSal = False
     If Not ValidateStateCode_TDSal() Then ValidatesheetTDSal = False
     If Not ValidatePinCode_TDSal() Then ValidatesheetTDSal = False
     If Not ValidateIncChrgSal_TDSal() Then ValidatesheetTDSal = False
     If Not ValidateDeductUnderChapVIA_TDSal() Then ValidatesheetTDSal = False
     If Not ValidateTaxPayIncluSurchEdnCes_TDSal() Then ValidatesheetTDSal = False
     If Not ValidateTotalTDSSal_TDSal() Then ValidatesheetTDSal = False
     If Not ValidateTaxPayRefund_TDSal() Then ValidatesheetTDSal = False
End If

setTblinfo_TDSal2
'setTblinfo_TDSal
If (end_TDSal2 <> end_TDSal) Then
         msgbox_TDSal ("Enter compulsory fields for Sch TDS from Salary ")
         ValidatesheetTDSal = False
         Exit Function

End If
End Function

Function ValidatesheetTDSoth() As Boolean
     ValidatesheetTDSoth = True
     If Not ValidateTAN_TDSoth() Then ValidatesheetTDSoth = False
 If (Len(ThisComponent.Sheets(1).getCellRangeByName("TDSoth.TAN").getCellByPosition(0, 1).String) > 0) Then
     If Not ValidateEmployerOrDeductorOrCollecterName_TDSoth() Then ValidatesheetTDSoth = False
     If Not ValidateAddrDetail_TDSoth() Then ValidatesheetTDSoth = False
     If Not ValidateCityOrTownOrDistrict_TDSoth() Then ValidatesheetTDSoth = False
     If Not ValidateStateCode_TDSoth() Then ValidatesheetTDSoth = False
     If Not ValidatePinCode_TDSoth() Then ValidatesheetTDSoth = False
     If Not ValidateAmtPaid_TDSoth() Then ValidatesheetTDSoth = False
     If Not ValidateDatePayCred_TDSoth() Then ValidatesheetTDSoth = False
     If Not ValidateTotTDSOnAmtPaid_TDSoth() Then ValidatesheetTDSoth = False
     If Not ValidateClaimOutOfTotTDSOnAmtPaid_TDSoth() Then ValidatesheetTDSoth = False
 End If
 setTblinfo_TDSoth2

If (end_TDSoth2 <> end_TDSoth) Then
         msgbox_TDSoth ("Enter all compulsory fields for Sch TDS on Interest")
         ValidatesheetTDSoth = False
         Exit Function

End If
End Function


Function ValidatesheetTaxP() As Boolean
     ValidatesheetTaxP = True
     If Not ValidateBSRCode_TaxP() Then ValidatesheetTaxP = False
 If (Len(ThisComponent.Sheets(1).getCellRangeByName("TaxP.BSRCode").getCellByPosition(0, 1).String) > 0) Then
     If Not ValidateNameOfBank_TaxP() Then ValidatesheetTaxP = False
     If Not ValidateNameOfBranch_TaxP() Then ValidatesheetTaxP = False
     'If Not ValidateBSRCode_TaxP() Then ValidatesheetTaxP = False
     If Not ValidateDateDep_TaxP() Then ValidatesheetTaxP = False
     If Not ValidateSrlNoOfChaln_TaxP() Then ValidatesheetTaxP = False
     If Not ValidateAmt_TaxP() Then ValidatesheetTaxP = False
 End If
setTblinfo_TaxP2

If (end_TaxP2 <> end_TaxP) Then
         msgbox_TaxP ("Details of Adv Tax and Self Asst Tax in Sheet TDS is Compulsory")
         ValidatesheetTaxP = False
         Exit Function

End If
End Function

Function ValidatesheetAIR() As Boolean
     ValidatesheetAIR = True
     If Not ValidateCode001_AIR() Then ValidatesheetAIR = False
     If Not ValidateCode002_AIR() Then ValidatesheetAIR = False
     If Not ValidateCode003_AIR() Then ValidatesheetAIR = False
     If Not ValidateCode004_AIR() Then ValidatesheetAIR = False
     If Not ValidateCode005_AIR() Then ValidatesheetAIR = False
     If Not ValidateCode006_AIR() Then ValidatesheetAIR = False
     If Not ValidateCode007_AIR() Then ValidatesheetAIR = False
     If Not ValidateCode008_AIR() Then ValidatesheetAIR = False
     If Not ValidateTaxExmpIntInc_AIR() Then ValidatesheetAIR = False
End Function

Function ValidatesheetVer() As Boolean
     ValidatesheetVer = True
     If Not ValidateAssesseeVerName_Ver() Then ValidatesheetVer = False
 'If (Len(Sheet2.Range("Ver.AssesseeVerName")) > 0) Then
     If Not ValidateFatherName_Ver() Then ValidatesheetVer = False
     If Not ValidatePAN_Ver() Then ValidatesheetVer = False
     If Not ValidatePlace_Ver() Then ValidatesheetVer = False
     If Not ValidateDate_Ver() Then ValidatesheetVer = False
 'End If
 If (Len(ThisComponent.Sheets(2-1).getCellRangeByName("Ver.IdentificationNoOfTRP").String) > 0) Then
     If Not ValidateIdentificationNoOfTRP_Ver() Then ValidatesheetVer = False
     If Not ValidateNameOfTRP_Ver() Then ValidatesheetVer = False
     If Not ValidateReImbFrmGov_Ver() Then ValidatesheetVer = False
 End If
End Function

Function ValidateIncomeFromSal_IncD() As Boolean
 ValidateIncomeFromSal_IncD = True
 IncomeFromSal_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.IncomeFromSal").Value
End Function
Function checkhpresponse() As Integer

Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Msg = "If the property is Occupied by you (Self occupied) /then the maximum you can claim is Rs -1,50,000. Is your property Self Occupied?"
Style = vbYesNo + vbInformation + vbDefaultButton1    ' Define buttons.
Title = "House Property"    ' Define title.
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
checkhpresponse = 6
Else
checkhpresponse = 7
End If

End Function
Function ValidateIncomeFromHP_IncD() As Boolean
Dim hpresponse As Integer

 ValidateIncomeFromHP_IncD = True
 IncomeFromHP_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.IncomeFromHP").Value
 
 If IncomeFromHP_IncD < -150000 Then
hpresponse = checkhpresponse
 If hpresponse = 6 Then
 IncomeFromHP_IncD = -150000
 ThisComponent.Sheets(1-1).getCellRangeByName("IncD.IncomeFromHP").Value = IncomeFromHP_IncD
 End If
End If

 
End Function


Function ValidateFamPension_IncD() As Boolean
 ValidateFamPension_IncD = True
 FamPension_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.FamPension").Value
End Function


Function ValidateIndInterest_IncD() As Boolean
 ValidateIndInterest_IncD = True
 IndInterest_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.IndInterest").Value
End Function


Function ValidateIncomeFromOS_IncD() As Boolean
 ValidateIncomeFromOS_IncD = True
 IncomeFromOS_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.IncomeFromOS").Value
End Function


Function ValidateGrossTotIncome_IncD() As Boolean
 ValidateGrossTotIncome_IncD = True
 GrossTotIncome_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.GrossTotIncome").Value
End Function


Function ValidateSection80C_IncD() As Boolean
 ValidateSection80C_IncD = True
 Section80C_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.Section80C").Value
End Function


Function ValidateSection80CCC_IncD() As Boolean
 ValidateSection80CCC_IncD = True
 Section80CCC_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.Section80CCC").Value
End Function


Function ValidateSection80CCD_IncD() As Boolean
 ValidateSection80CCD_IncD = True
 Section80CCD_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.Section80CCD").Value
End Function
Function ValidateSection80CCF_IncD() As Boolean
 ValidateSection80CCF_IncD = True
 Section80CCF_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.Section80CCF").Value
End Function


Function ValidateSection80D_IncD() As Boolean
 ValidateSection80D_IncD = True
 Section80D_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.Section80D").Value
End Function


Function ValidateSection80DD_IncD() As Boolean
 ValidateSection80DD_IncD = True
 Section80DD_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.Section80DD").Value
End Function


Function ValidateSection80DDB_IncD() As Boolean
 ValidateSection80DDB_IncD = True
 Section80DDB_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.Section80DDB").Value
End Function


Function ValidateSection80E_IncD() As Boolean
 ValidateSection80E_IncD = True
 Section80E_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.Section80E").Value
End Function


Function ValidateSection80G_IncD() As Boolean
 ValidateSection80G_IncD = True
 Section80G_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.Section80G").Value
End Function


Function ValidateSection80GG_IncD() As Boolean
 ValidateSection80GG_IncD = True
 Section80GG_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.Section80GG").Value
End Function


Function ValidateSection80GGA_IncD() As Boolean
 ValidateSection80GGA_IncD = True
 Section80GGA_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.Section80GGA").Value
End Function


Function ValidateSection80GGC_IncD() As Boolean
 ValidateSection80GGC_IncD = True
 Section80GGC_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.Section80GGC").Value
End Function


Function ValidateSection80U_IncD() As Boolean
 ValidateSection80U_IncD = True
 Section80U_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.Section80U").Value
End Function


Function ValidateTotalChapVIADeductions_IncD() As Boolean
 ValidateTotalChapVIADeductions_IncD = True
 TotalChapVIADeductions_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.TotalChapVIADeductions").Value
End Function


Function ValidateTotalIncome_IncD() As Boolean
 ValidateTotalIncome_IncD = True
 TotalIncome_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.TotalIncome").Value
End Function


Function ValidateNetAgriculturalIncome_IncD() As Boolean
 ValidateNetAgriculturalIncome_IncD = True
 NetAgriculturalIncome_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.NetAgriculturalIncome").Value
End Function


Function ValidateAggregateIncome_IncD() As Boolean
 ValidateAggregateIncome_IncD = True
 AggregateIncome_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.AggregateIncome").Value
End Function


Function ValidateTaxOnAggregateInc_IncD() As Boolean
 ValidateTaxOnAggregateInc_IncD = True
 TaxOnAggregateInc_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.TaxOnAggregateInc").Value
End Function


Function ValidateRebateOnAgriInc_IncD() As Boolean
 ValidateRebateOnAgriInc_IncD = True
 RebateOnAgriInc_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.RebateOnAgriInc").Value
End Function


Function ValidateTotalTaxPayable_IncD() As Boolean
 ValidateTotalTaxPayable_IncD = True
 TotalTaxPayable_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.TotalTaxPayable").Value
End Function


Function ValidateSurchargeOnTaxPayable_IncD() As Boolean
 ValidateSurchargeOnTaxPayable_IncD = True
 SurchargeOnTaxPayable_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.SurchargeOnTaxPayable").Value
End Function


Function ValidateEducationCess_IncD() As Boolean
 ValidateEducationCess_IncD = True
 EducationCess_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.EducationCess").Value
End Function


Function ValidateGrossTaxLiability_IncD() As Boolean
 ValidateGrossTaxLiability_IncD = True
 GrossTaxLiability_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.GrossTaxLiability").Value
End Function


Function ValidateSection89_IncD() As Boolean
 ValidateSection89_IncD = True
 Section89_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.Section89").Value
End Function


Function ValidateSection90and91_IncD() As Boolean
 ValidateSection90and91_IncD = True
 Section90and91_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.Section90and91").Value
End Function


Function ValidateNetTaxLiability_IncD() As Boolean
 ValidateNetTaxLiability_IncD = True
 NetTaxLiability_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.NetTaxLiability").Value
End Function


Function ValidateIntrstPayUs234A_IncD() As Boolean
 ValidateIntrstPayUs234A_IncD = True
 IntrstPayUs234A_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.IntrstPayUs234A").Value
End Function


Function ValidateIntrstPayUs234B_IncD() As Boolean
 ValidateIntrstPayUs234B_IncD = True
 IntrstPayUs234B_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.IntrstPayUs234B").Value
End Function


Function ValidateIntrstPayUs234C_IncD() As Boolean
 ValidateIntrstPayUs234C_IncD = True
 IntrstPayUs234C_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.IntrstPayUs234C").Value
End Function


Function ValidateTotalIntrstPay_IncD() As Boolean
 ValidateTotalIntrstPay_IncD = True
 TotalIntrstPay_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.TotalIntrstPay").Value
End Function


Function ValidateTotTaxPlusIntrstPay_IncD() As Boolean
 ValidateTotTaxPlusIntrstPay_IncD = True
 TotTaxPlusIntrstPay_IncD = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.TotTaxPlusIntrstPay").Value
End Function


Function ValidateAdvanceTax_IncD() As Boolean
 ValidateAdvanceTax_IncD = True
 AdvanceTax_IncD = ThisComponent.Sheets(2-1).getCellRangeByName("IncD.AdvanceTax").Value
End Function


Function ValidateTDS_IncD() As Boolean
 ValidateTDS_IncD = True
 TDS_IncD = ThisComponent.Sheets(2-1).getCellRangeByName("IncD.TDS").Value
End Function


Function ValidateSelfAssessmentTax_IncD() As Boolean
 ValidateSelfAssessmentTax_IncD = True
 SelfAssessmentTax_IncD = ThisComponent.Sheets(2-1).getCellRangeByName("IncD.SelfAssessmentTax").Value
End Function


Function ValidateTotalTaxesPaid_IncD() As Boolean
 ValidateTotalTaxesPaid_IncD = True
 TotalTaxesPaid_IncD = ThisComponent.Sheets(2-1).getCellRangeByName("IncD.TotalTaxesPaid").Value
End Function


Function ValidateBalTaxPayable_IncD() As Boolean
 ValidateBalTaxPayable_IncD = True
 BalTaxPayable_IncD = ThisComponent.Sheets(2-1).getCellRangeByName("IncD.BalTaxPayable").Value
End Function


Function ValidateRefundDue_IncD() As Boolean
 ValidateRefundDue_IncD = True
 RefundDue_IncD = ThisComponent.Sheets(2-1).getCellRangeByName("IncD.RefundDue").Value
End Function

Function ValidateBankAccountNumber_IncD() As Boolean
 
 ValidateBankAccountNumber_IncD = True
 BankAccountNumber_IncD = ThisComponent.Sheets(2-1).getCellRangeByName("IncD.BankAccountNumber").String
     If Not chkCompulsory(BankAccountNumber_IncD) Then
         msgbox_IncD ("BankAccountNumber in Sheet Taxes paid and Verification  is Compulsory")
         ValidateBankAccountNumber_IncD = False
         Exit Function
     End If
    If Not IsNumeric(BankAccountNumber_IncD) Then
         msgbox_IncD ("BankAccountNumber in Sheet Taxes paid and Verification only digits 0 to 9 allowed")
          ValidateBankAccountNumber_IncD = False
          Exit Function
     End If
    If BankAccountNumber_IncD < 0 Then
         msgbox_IncD ("BankAccountNumber in Sheet Taxes paid and Verification only digits 0 to 9 allowed")
          ValidateBankAccountNumber_IncD = False
          Exit Function
     End If
     
End Function
 

Function ValidateEcsRequired_IncD() As Boolean
  ValidateEcsRequired_IncD = True
 EcsRequired_IncD = ThisComponent.Sheets(2-1).getCellRangeByName("IncD.EcsRequired").Value
 EcsRequired_IncD = Mid(EcsRequired_IncD, 1, 1)
End Function

Function ValidateMICRCode_IncD() As Boolean
 ValidateMICRCode_IncD = True
 MICRCode_IncD = ThisComponent.Sheets(2-1).getCellRangeByName("IncD.MICRCode").Value
     If Not chkCompulsory(MICRCode_IncD) Then
         msgbox_IncD ("MICRCode in Sheet Taxes paid and Verification  is Compulsory")
         ValidateMICRCode_IncD = False
         Exit Function
     End If
     If Not chkNumeric(MICRCode_IncD) Then
         msgbox_IncD ("MICRCode  in Sheet Taxes paid and Verification  only digits 0 to 9 allowed ")
         ValidateMICRCode_IncD = False
         Exit Function
     End If
End Function


Function ValidateBankAccountType_IncD() As Boolean
  ValidateBankAccountType_IncD = True
 BankAccountType_IncD = ThisComponent.Sheets(2-1).getCellRangeByName("IncD.BankAccountType").Value
 BankAccountType_IncD = Mid(BankAccountType_IncD, 1, 3)
     If Not chkCompulsory(BankAccountType_IncD) Then
         msgbox_IncD ("BankAccountType in Sheet Taxes paid and Verification  is Compulsory")
         ValidateBankAccountType_IncD = False
         Exit Function
     End If
End Function

Function ValidateTAN_TDSal() As Boolean
    ValidateTAN_TDSal = True
    setTblinfo_TDSal
    Dim rangecells As Object
    Set rangecells = ThisComponent.Sheets(1).getCellRangeByName("TDSal.TAN")
    ReDim TAN_TDSal(end_TDSal)
    For i = 0  To end_TDSal - 1
        TAN_TDSal(i) = rangecells.getCellByPosition(0, i).String
 If Not Len(TAN_TDSal(i)) = 0 Then
     If Not ValidateTantype_text(Mid(TAN_TDSal(i), 1, 4)) Then
         msgbox_TDSal ("TAN at Sr. No  " & i & " in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet")
         ValidateTAN_TDSal = False
         Exit Function
     End If
     If Not IsNumeric(Mid(TAN_TDSal(i), 5, 5)) Then
         msgbox_TDSal ("TAN at Sr. No  " & i & "  in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet")
         ValidateTAN_TDSal = False
         Exit Function
     End If
     If Not ValidateTantype_text(Right(TAN_TDSal(i), 1)) Then
         msgbox_TDSal ("TAN at Sr. No  " & i & "  in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet")
         ValidateTAN_TDSal = False
         Exit Function
     End If
 ElseIf Not chkCompulsory(TAN_TDSal(i)) Then
         msgbox_TDSal ("TAN at Sr. No  " & i & "  in Sheet TDS  is Compulsory")
     ValidateTAN_TDSal = False
     Exit Function
 End If
 Next
End Function

Function ValidateEmployerOrDeductorOrCollecterName_TDSal() As Boolean
 
    ValidateEmployerOrDeductorOrCollecterName_TDSal = True
    setTblinfo_TDSal
    Dim rangecells As Object
    Set rangecells = ThisComponent.Sheets(1).getCellRangeByName("TDSal.EmployerOrDeductorOrCollecterName")
    ReDim EmployerOrDeductorOrCollecterName_TDSal(end_TDSal)
    For i = 0  To end_TDSal - 1
        EmployerOrDeductorOrCollecterName_TDSal(i) = rangecells.getCellByPosition(0, i).String
     If Not chkCompulsory(EmployerOrDeductorOrCollecterName_TDSal(i)) Then
         msgbox_TDSal ("EmployerOrDeductorOrCollecterName at Sr. No  " & i & "  in Sheet TDS  is Compulsory")
         ValidateEmployerOrDeductorOrCollecterName_TDSal = False
         Exit Function
     End If
    If Not checkfieldspecialcharacter(EmployerOrDeductorOrCollecterName_TDSal(i)) Then
         msgbox_TDSal ("EmployerOrDeductorOrCollecterName at Sr. No  " & i & " in Sheet TDS  characters < > & ' " & Chr(34) & " are not allowed")
          ValidateEmployerOrDeductorOrCollecterName_TDSal = False
          Exit Function
     End If
 Next
End Function


 
Function ValidateAddrDetail_TDSal() As Boolean
 
    ValidateAddrDetail_TDSal = True
End Function
 
Function ValidateCityOrTownOrDistrict_TDSal() As Boolean
 
    ValidateCityOrTownOrDistrict_TDSal = True
End Function
 

Function ValidateStateCode_TDSal() As Boolean
    ValidateStateCode_TDSal = True
End Function

Function ValidatePinCode_TDSal() As Boolean
    ValidatePinCode_TDSal = True
End Function


Function ValidateIncChrgSal_TDSal() As Boolean
    ValidateIncChrgSal_TDSal = True
    setTblinfo_TDSal
    Dim rangecells As Object
    Set rangecells = ThisComponent.Sheets(1).getCellRangeByName("TDSal.IncChrgSal")
    ReDim IncChrgSal_TDSal(end_TDSal)
    For i = 0  To end_TDSal - 1
        IncChrgSal_TDSal(i) = rangecells.getCellByPosition(0, i).String
 Next
End Function


Function ValidateDeductUnderChapVIA_TDSal() As Boolean
    ValidateDeductUnderChapVIA_TDSal = True
End Function

Function ValidateTaxPayIncluSurchEdnCes_TDSal() As Boolean
    ValidateTaxPayIncluSurchEdnCes_TDSal = True
End Function


Function ValidateTotalTDSSal_TDSal() As Boolean
    ValidateTotalTDSSal_TDSal = True
    setTblinfo_TDSal
    Dim rangecells As Object
    Set rangecells = ThisComponent.Sheets(1).getCellRangeByName("TDSal.TotalTDSSal")
    ReDim TotalTDSSal_TDSal(end_TDSal)
    For i = 0  To end_TDSal - 1
        TotalTDSSal_TDSal(i) = rangecells.getCellByPosition(0, i).String
 Next
End Function


Function ValidateTaxPayRefund_TDSal() As Boolean
    ValidateTaxPayRefund_TDSal = True
End Function


Function ValidateTAN_TDSoth() As Boolean
    ValidateTAN_TDSoth = True
    setTblinfo_TDSoth
    Dim rangecells As Object
    Set rangecells = ThisComponent.Sheets(1).getCellRangeByName("TDSoth.TAN")
    ReDim TAN_TDSoth(end_TDSoth)
    For i = 0  To end_TDSoth - 1
        TAN_TDSoth(i) = rangecells.getCellByPosition(0, i).String
 If Not Len(TAN_TDSoth(i)) = 0 Then
     If Not ValidateTantype_text(Mid(TAN_TDSoth(i), 1, 4)) Then
         msgbox_TDSoth ("TAN at Sr. No  " & i & " in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet")
         ValidateTAN_TDSoth = False
         Exit Function
     End If
     If Not IsNumeric(Mid(TAN_TDSoth(i), 5, 5)) Then
         msgbox_TDSoth ("TAN at Sr. No  " & i & "  in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet")
         ValidateTAN_TDSoth = False
         Exit Function
     End If
     If Not ValidateTantype_text(Right(TAN_TDSoth(i), 1)) Then
         msgbox_TDSoth ("TAN at Sr. No  " & i & "  in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet")
         ValidateTAN_TDSoth = False
         Exit Function
     End If
 ElseIf Not chkCompulsory(TAN_TDSoth(i)) Then
         msgbox_TDSoth ("TAN at Sr. No  " & i & "  in Sheet TDS  is Compulsory")
     ValidateTAN_TDSoth = False
     Exit Function
 End If
 Next
End Function

Function ValidateEmployerOrDeductorOrCollecterName_TDSoth() As Boolean
 
    ValidateEmployerOrDeductorOrCollecterName_TDSoth = True
    setTblinfo_TDSoth
    Dim rangecells As Object
    Set rangecells = ThisComponent.Sheets(1).getCellRangeByName("TDSoth.EmployerOrDeductorOrCollecterName")
    ReDim EmployerOrDeductorOrCollecterName_TDSoth(end_TDSoth)
    For i = 0  To end_TDSoth - 1
        EmployerOrDeductorOrCollecterName_TDSoth(i) = rangecells.getCellByPosition(0, i).String
     If Not chkCompulsory(EmployerOrDeductorOrCollecterName_TDSoth(i)) Then
         msgbox_TDSoth ("EmployerOrDeductorOrCollecterName at Sr. No  " & i & "  in Sheet TDS  is Compulsory")
         ValidateEmployerOrDeductorOrCollecterName_TDSoth = False
         Exit Function
     End If
    If Not checkfieldspecialcharacter(EmployerOrDeductorOrCollecterName_TDSoth(i)) Then
         msgbox_TDSoth ("EmployerOrDeductorOrCollecterName at Sr. No  " & i & " in Sheet TDS  characters < > & ' " & Chr(34) & " are not allowed")
          ValidateEmployerOrDeductorOrCollecterName_TDSoth = False
          Exit Function
     End If
 Next
End Function
 
Function ValidateAddrDetail_TDSoth() As Boolean
 
    ValidateAddrDetail_TDSoth = True
End Function
 
Function ValidateCityOrTownOrDistrict_TDSoth() As Boolean
 
    ValidateCityOrTownOrDistrict_TDSoth = True
End Function
 

Function ValidateStateCode_TDSoth() As Boolean
    ValidateStateCode_TDSoth = True
End Function

Function ValidatePinCode_TDSoth() As Boolean
    ValidatePinCode_TDSoth = True
End Function


Function ValidateAmtPaid_TDSoth() As Boolean
    ValidateAmtPaid_TDSoth = True
End Function

                                                                    
Function ValidateDatePayCred_TDSoth() As Boolean
    ValidateDatePayCred_TDSoth = True
End Function


Function ValidateTotTDSOnAmtPaid_TDSoth() As Boolean
    ValidateTotTDSOnAmtPaid_TDSoth = True
    setTblinfo_TDSoth
    Dim rangecells As Object
    Set rangecells = ThisComponent.Sheets(1).getCellRangeByName("TDSoth.TotTDSOnAmtPaid")
    ReDim TotTDSOnAmtPaid_TDSoth(end_TDSoth)
    For i = 0  To end_TDSoth - 1
        TotTDSOnAmtPaid_TDSoth(i) = rangecells.getCellByPosition(0, i).String
 Next
End Function


Function ValidateClaimOutOfTotTDSOnAmtPaid_TDSoth() As Boolean
    ValidateClaimOutOfTotTDSOnAmtPaid_TDSoth = True
    setTblinfo_TDSoth
    Dim rangecells As Object
    Set rangecells = ThisComponent.Sheets(1).getCellRangeByName("TDSoth.ClaimOutOfTotTDSOnAmtPaid")
    ReDim ClaimOutOfTotTDSOnAmtPaid_TDSoth(end_TDSoth)
    Dim msgtdsothwarning As Boolean
    Dim msgtdsothwarningmessage As String
    msgtdsothwarningmessage = "Typically amount in Col 7 of TDS Other than Salary would be the same as amount in Col 6. However, in some rows the amount in Col 7 has been left blank or is less than the amount in Col 6. Please verify and change if required."
    For i = 0  To end_TDSoth - 1
        ClaimOutOfTotTDSOnAmtPaid_TDSoth(i) = rangecells.getCellByPosition(0, i).String
        
        If (TotTDSOnAmtPaid_TDSoth(i) = "") Then
        TotTDSOnAmtPaid_TDSoth(i) = 0
        End If
        
            If ClaimOutOfTotTDSOnAmtPaid_TDSoth(i) > TotTDSOnAmtPaid_TDSoth(i) Then
                msgbox_TDSoth ("(7) cannot be greater than (6) at Sr. No  " & i & " in Sheet TDS ")
                ValidateClaimOutOfTotTDSOnAmtPaid_TDSoth = False
                Exit Function
            End If
            If ClaimOutOfTotTDSOnAmtPaid_TDSoth(i) = "" Then
                msgtdsothwarning = True
            End If
            If ClaimOutOfTotTDSOnAmtPaid_TDSoth(i) = 0 Then
                msgtdsothwarning = True
            End If
 
 Next
 If msgtdsothwarning = True Then
    MsgBox msgtdsothwarningmessage, vbOKOnly + vbExclamation, "TDS Other than Salary"
 End If
 
 
End Function

Function ValidateNameOfBank_TaxP() As Boolean
    ValidateNameOfBank_TaxP = True
End Function
 
Function ValidateNameOfBranch_TaxP() As Boolean
    ValidateNameOfBranch_TaxP = True
End Function
 

Function ValidateBSRCode_TaxP() As Boolean
    ValidateBSRCode_TaxP = True
    setTblinfo_TaxP
    Dim rangecells As Object
    Set rangecells = ThisComponent.Sheets(1).getCellRangeByName("TaxP.BSRCode")
    ReDim BSRCode_TaxP(end_TaxP)
    For i = 0  To end_TaxP - 1
        BSRCode_TaxP(i) = rangecells.getCellByPosition(0, i).String
     If Not chkCompulsory(BSRCode_TaxP(i)) Then
         msgbox_TaxP ("BSRCode at Sr. No  " & i & "  in Sheet TDS  is Compulsory")
         ValidateBSRCode_TaxP = False
         Exit Function
     End If
     If Not chkNumeric(BSRCode_TaxP(i)) Then
         msgbox_TaxP ("BSRCode at Sr. No  " & i & "  in Sheet TDS  only digits 0 to 9 allowed")
         ValidateBSRCode_TaxP = False
         Exit Function
     End If
 Next
End Function

                                                                    
Function ValidateDateDep_TaxP() As Boolean
    ValidateDateDep_TaxP = True
    setTblinfo_TaxP
    Dim rangecells As Object
    Set rangecells = ThisComponent.Sheets(1).getCellRangeByName("TaxP.DateDep")
    ReDim DateDep_TaxP(end_TaxP)
    For i = 0  To end_TaxP - 1
        DateDep_TaxP(i) = rangecells.getCellByPosition(0, i).String
If Not chkCompulsory(DateDep_TaxP(i)) Then
         msgbox_TaxP ("DateDep at Sr. No  " & i & "  in Sheet TDS  is Compulsory")
    ValidateDateDep_TaxP = False
 Exit Function
End If
If Not CheckDateddmmyyyy(DateDep_TaxP(i)) Then
    ValidateDateDep_TaxP = False
         msgbox_TaxP ("DateDep at Sr. No  " & i & "  in Sheet TDS  is invalid. Pl enter in dd/mm/yyyy format")
 Exit Function
Else
  DateDep_TaxP(i) = Dformat(DateDep_TaxP(i))
     'changed
     If Not ChkMinInclusiveDate(DateDep_TaxP(i), "2010-04-01") Then
         ValidateDateDep_TaxP = False
         'changed
         msgbox_TaxP ("DateDep at Sr. No  " & i & " in Sheet TDS  should not be prior to 2010-04-01")
         Exit Function
     End If

End If
 
 Next
End Function


Function ValidateSrlNoOfChaln_TaxP() As Boolean
    ValidateSrlNoOfChaln_TaxP = True
    setTblinfo_TaxP
    Dim rangecells As Object
    Set rangecells = ThisComponent.Sheets(1).getCellRangeByName("TaxP.SrlNoOfChaln")
    ReDim SrlNoOfChaln_TaxP(end_TaxP)
    For i = 0  To end_TaxP - 1
        SrlNoOfChaln_TaxP(i) = rangecells.getCellByPosition(0, i).String
     If Not chkCompulsory(SrlNoOfChaln_TaxP(i)) Then
         msgbox_TaxP ("SrlNoOfChaln at Sr. No  " & i & "  in Sheet TDS  is Compulsory")
         ValidateSrlNoOfChaln_TaxP = False
         Exit Function
     End If
     If Not chkNumeric(SrlNoOfChaln_TaxP(i)) Then
         msgbox_TaxP ("SrlNoOfChaln at Sr. No  " & i & "  in Sheet TDS  only digits 0 to 9 allowed")
         ValidateSrlNoOfChaln_TaxP = False
         Exit Function
     End If
 Next
End Function


Function ValidateAmt_TaxP() As Boolean
    ValidateAmt_TaxP = True
    setTblinfo_TaxP
    Dim rangecells As Object
    Set rangecells = ThisComponent.Sheets(1).getCellRangeByName("TaxP.Amt")
    ReDim Amt_TaxP(end_TaxP)
    For i = 0  To end_TaxP - 1
        Amt_TaxP(i) = rangecells.getCellByPosition(0, i).String
 Next
End Function


Function ValidateCode001_AIR() As Boolean
 ValidateCode001_AIR = True
 
End Function


Function ValidateCode002_AIR() As Boolean
 ValidateCode002_AIR = True

End Function


Function ValidateCode003_AIR() As Boolean
 ValidateCode003_AIR = True

End Function


Function ValidateCode004_AIR() As Boolean
 ValidateCode004_AIR = True

End Function


Function ValidateCode005_AIR() As Boolean
 ValidateCode005_AIR = True

End Function


Function ValidateCode006_AIR() As Boolean
 ValidateCode006_AIR = True

End Function


Function ValidateCode007_AIR() As Boolean
 ValidateCode007_AIR = True

End Function


Function ValidateCode008_AIR() As Boolean
 ValidateCode008_AIR = True

End Function


Function ValidateTaxExmpIntInc_AIR() As Boolean
 ValidateTaxExmpIntInc_AIR = True
 TaxExmpIntInc_AIR = ThisComponent.Sheets(2-1).getCellRangeByName("AIR.TaxExmpIntInc").Value
End Function

Function ValidateAssesseeVerName_Ver() As Boolean
 
 ValidateAssesseeVerName_Ver = True
 AssesseeVerName_Ver = ThisComponent.Sheets(2-1).getCellRangeByName("Ver.AssesseeVerName").String
     If Not chkCompulsory(AssesseeVerName_Ver) Then
         msgbox_Ver ("AssesseeVerName in Sheet Taxes paid and Verification  is Compulsory")
         ValidateAssesseeVerName_Ver = False
         Exit Function
     End If
    If Not checkfieldspecialcharacter(AssesseeVerName_Ver) Then
         msgbox_Ver ("AssesseeVerName in Sheet Taxes paid and Verification  characters < > & ' " & Chr(34) & " are not allowed")
          ValidateAssesseeVerName_Ver = False
          Exit Function
     End If
End Function
 
Function ValidateFatherName_Ver() As Boolean
 
 ValidateFatherName_Ver = True
 FatherName_Ver = ThisComponent.Sheets(2-1).getCellRangeByName("Ver.FatherName").String
     If Not chkCompulsory(FatherName_Ver) Then
         msgbox_Ver ("Fathers name in Sheet Taxes paid and Verification  is Compulsory")
         ValidateFatherName_Ver = False
         Exit Function
     End If
 If Len(FatherName_Ver) > 0 Then
    If Not checkfieldspecialcharacter(FatherName_Ver) Then
         msgbox_Ver ("FatherName in Sheet Taxes paid and Verification  characters < > & ' " & Chr(34) & " are not allowed")
          ValidateFatherName_Ver = False
          Exit Function
     End If
 End If
End Function
 
Function ValidatePlace_Ver() As Boolean
 
 ValidatePlace_Ver = True
 Place_Ver = ThisComponent.Sheets(2-1).getCellRangeByName("Ver.Place").String
     If Not chkCompulsory(Place_Ver) Then
         msgbox_Ver ("Place in Sheet Taxes paid and Verification  is Compulsory")
         ValidatePlace_Ver = False
         Exit Function
     End If
    If Not checkfieldspecialcharacter(Place_Ver) Then
         msgbox_Ver ("Place in Sheet Taxes paid and Verification  characters < > & ' " & Chr(34) & " are not allowed")
          ValidatePlace_Ver = False
          Exit Function
     End If
End Function
 
                                                                    
Function ValidateDate_Ver() As Boolean
 ValidateDate_Ver = True
 Date_Ver = ThisComponent.Sheets(2-1).getCellRangeByName("Ver.Date").Value
If Not chkCompulsory(Date_Ver) Then
         msgbox_Ver ("Date in Sheet Taxes paid and Verification  is Compulsory")
    ValidateDate_Ver = False
 Exit Function
End If
If Not CheckDateddmmyyyy(Date_Ver) Then
    ValidateDate_Ver = False
    msgbox_Ver ("Date in Sheet Taxes paid and Verification  is invalid. Pl enter in dd/mm/yyyy format")
    Exit Function
Else
  Date_Ver = Dformat(Date_Ver)
     If Not ChkMinInclusiveDate(Date_Ver, "2011-04-01") Then
         msgbox_Ver ("Date in Sheet Taxes paid and Verification  should not be less than 2011-04-01")
         ValidateDate_Ver = False
         Exit Function
     End If
End If
End Function

Function ValidateIdentificationNoOfTRP_Ver() As Boolean
 
 ValidateIdentificationNoOfTRP_Ver = True
 IdentificationNoOfTRP_Ver = ThisComponent.Sheets(2-1).getCellRangeByName("Ver.IdentificationNoOfTRP").String
     If Not chkCompulsory(IdentificationNoOfTRP_Ver) Then
         msgbox_Ver ("IdentificationNoOfTRP in Sheet Taxes paid and Verification  is Compulsory")
         ValidateIdentificationNoOfTRP_Ver = False
         Exit Function
     End If
    If Not checkfieldspecialcharacter(IdentificationNoOfTRP_Ver) Then
         msgbox_Ver ("IdentificationNoOfTRP in Sheet Taxes paid and Verification  characters < > & ' " & Chr(34) & " are not allowed")
          ValidateIdentificationNoOfTRP_Ver = False
          Exit Function
     End If
End Function
Function ValidatePAN_Ver() As Boolean
 
 ValidatePAN_Ver = True
 verPAN = ThisComponent.Sheets(2-1).getCellRangeByName("Ver.PAN").String
 PAN_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.PAN").String
 Dim tempPAN_1 As String
 tempPAN_1 = PAN_1
 Dim tempPAN_ver As String
 tempPAN_ver = verPAN
     If Len(verPAN) > 10 Then
        msgbox_Ver ("PAN in Sheet : Taxes paid and Verification  should be 10 digits")
        ValidatePAN_Ver = False
        Exit Function
     End If
     If Not chkCompulsory(verPAN) Then
         msgbox_Ver ("PAN in Sheet Taxes paid and Verification  is Compulsory")
         ValidatePAN_Ver = False
         Exit Function
     End If
     If Not ValidatePAN(tempPAN_ver) Then
        msgbox_Ver ("PAN in Sheet : Taxes paid and Verification  is invalid (10 digits valid PAN)")
        ValidatePAN_Ver = False
        Exit Function
     End If
     If Not checkfieldspecialcharacter(verPAN) Then
         msgbox_Ver ("PAN in Sheet Taxes paid and Verification  characters < > & ' " & Chr(34) & " are not allowed")
          ValidatePAN_Ver = False
          Exit Function
     End If
'     If Not verPAN = tempPAN_1 Then
'         msgbox_Ver ("PAN must be same in both Sheet General and General2")
'          ValidatePAN_Ver = False
'          Exit Function
'     End If
End Function
 
Function ValidateNameOfTRP_Ver() As Boolean
 
 ValidateNameOfTRP_Ver = True
 NameOfTRP_Ver = ThisComponent.Sheets(2-1).getCellRangeByName("Ver.NameOfTRP").String
     If Not chkCompulsory(NameOfTRP_Ver) Then
         msgbox_Ver ("NameOfTRP in Sheet Taxes paid and Verification  is Compulsory")
         ValidateNameOfTRP_Ver = False
         Exit Function
     End If
    If Not checkfieldspecialcharacter(NameOfTRP_Ver) Then
         msgbox_Ver ("NameOfTRP in Sheet Taxes paid and Verification  characters < > & ' " & Chr(34) & " are not allowed")
          ValidateNameOfTRP_Ver = False
          Exit Function
     End If
End Function
 

Function ValidateReImbFrmGov_Ver() As Boolean
 ValidateReImbFrmGov_Ver = True
 ReImbFrmGov_Ver = ThisComponent.Sheets(2-1).getCellRangeByName("Ver.ReImbFrmGov").Value
 If Not chkCompulsory(ReImbFrmGov_Ver) Then
         msgbox_Ver ("ReImbFrmGov in Sheet Taxes paid and Verification  is Compulsory")
         ValidateReImbFrmGov_Ver = False
         Exit Function
End If
End Function

Function DefaultIncomeFromSal_IncD() As String
DefaultIncomeFromSal_IncD = "0"
End Function
Function DefaultIncomeFromHP_IncD() As String
DefaultIncomeFromHP_IncD = "0"
End Function
Function DefaultFamPension_IncD() As String
DefaultFamPension_IncD = "0"
End Function
Function DefaultIndInterest_IncD() As String
DefaultIndInterest_IncD = "0"
End Function
Function DefaultIncomeFromOS_IncD() As String
DefaultIncomeFromOS_IncD = "0"
End Function
Function DefaultGrossTotIncome_IncD() As String
DefaultGrossTotIncome_IncD = "0"
End Function
Function DefaultSection80C_IncD() As String
DefaultSection80C_IncD = "0"
End Function
Function DefaultSection80CCC_IncD() As String
DefaultSection80CCC_IncD = "0"
End Function
Function DefaultSection80CCD_IncD() As String
DefaultSection80CCD_IncD = "0"
End Function
Function DefaultSection80CCF_IncD() As String
DefaultSection80CCF_IncD = "0"
End Function
Function DefaultSection80D_IncD() As String
DefaultSection80D_IncD = "0"
End Function
Function DefaultSection80DD_IncD() As String
DefaultSection80DD_IncD = "0"
End Function
Function DefaultSection80DDB_IncD() As String
DefaultSection80DDB_IncD = "0"
End Function
Function DefaultSection80E_IncD() As String
DefaultSection80E_IncD = "0"
End Function
Function DefaultSection80G_IncD() As String
DefaultSection80G_IncD = "0"
End Function
Function DefaultSection80GG_IncD() As String
DefaultSection80GG_IncD = "0"
End Function
Function DefaultSection80GGA_IncD() As String
DefaultSection80GGA_IncD = "0"
End Function
Function DefaultSection80GGC_IncD() As String
DefaultSection80GGC_IncD = "0"
End Function
Function DefaultSection80U_IncD() As String
DefaultSection80U_IncD = "0"
End Function
Function DefaultTotalChapVIADeductions_IncD() As String
DefaultTotalChapVIADeductions_IncD = "0"
End Function
Function DefaultTotalIncome_IncD() As String
DefaultTotalIncome_IncD = "0"
End Function
Function DefaultNetAgriculturalIncome_IncD() As String
DefaultNetAgriculturalIncome_IncD = "0"
End Function
Function DefaultAggregateIncome_IncD() As String
DefaultAggregateIncome_IncD = "0"
End Function
Function DefaultTaxOnAggregateInc_IncD() As String
DefaultTaxOnAggregateInc_IncD = "0"
End Function
Function DefaultRebateOnAgriInc_IncD() As String
DefaultRebateOnAgriInc_IncD = "0"
End Function
Function DefaultTotalTaxPayable_IncD() As String
DefaultTotalTaxPayable_IncD = "0"
End Function
Function DefaultSurchargeOnTaxPayable_IncD() As String
DefaultSurchargeOnTaxPayable_IncD = "0"
End Function
Function DefaultEducationCess_IncD() As String
DefaultEducationCess_IncD = "0"
End Function
Function DefaultGrossTaxLiability_IncD() As String
DefaultGrossTaxLiability_IncD = "0"
End Function
Function DefaultSection89_IncD() As String
DefaultSection89_IncD = "0"
End Function
Function DefaultSection90and91_IncD() As String
DefaultSection90and91_IncD = "0"
End Function
Function DefaultNetTaxLiability_IncD() As String
DefaultNetTaxLiability_IncD = "0"
End Function
Function DefaultIntrstPayUs234A_IncD() As String
DefaultIntrstPayUs234A_IncD = "0"
End Function
Function DefaultIntrstPayUs234B_IncD() As String
DefaultIntrstPayUs234B_IncD = "0"
End Function
Function DefaultIntrstPayUs234C_IncD() As String
DefaultIntrstPayUs234C_IncD = "0"
End Function
Function DefaultTotalIntrstPay_IncD() As String
DefaultTotalIntrstPay_IncD = "0"
End Function
Function DefaultTotTaxPlusIntrstPay_IncD() As String
DefaultTotTaxPlusIntrstPay_IncD = "0"
End Function
Function DefaultAdvanceTax_IncD() As String
DefaultAdvanceTax_IncD = "0"
End Function
Function DefaultTDS_IncD() As String
DefaultTDS_IncD = "0"
End Function
Function DefaultSelfAssessmentTax_IncD() As String
DefaultSelfAssessmentTax_IncD = "0"
End Function
Function DefaultTotalTaxesPaid_IncD() As String
DefaultTotalTaxesPaid_IncD = "0"
End Function
Function DefaultBalTaxPayable_IncD() As String
DefaultBalTaxPayable_IncD = "0"
End Function
Function DefaultRefundDue_IncD() As String
DefaultRefundDue_IncD = "0"
End Function
Function DefaultEcsRequired_IncD() As String
DefaultEcsRequired_IncD = "N"
End Function
Function DefaultTDSonSalaries_TDSal() As String
DefaultTDSonSalaries_TDSal = "0"
End Function
Function DefaultTDSonSalary_TDSal() As String
DefaultTDSonSalary_TDSal = "0"
End Function
Function DefaultIncChrgSal_TDSal() As String
DefaultIncChrgSal_TDSal = "0"
End Function
Function DefaultDeductUnderChapVIA_TDSal() As String
DefaultDeductUnderChapVIA_TDSal = "0"
End Function
Function DefaultTaxPayIncluSurchEdnCes_TDSal() As String
DefaultTaxPayIncluSurchEdnCes_TDSal = "0"
End Function
Function DefaultTotalTDSSal_TDSal() As String
DefaultTotalTDSSal_TDSal = "0"
End Function
Function DefaultTaxPayRefund_TDSal() As String
DefaultTaxPayRefund_TDSal = "0"
End Function
Function DefaultTDSonOthThanSals_TDSoth() As String
DefaultTDSonOthThanSals_TDSoth = "0"
End Function
Function DefaultTDSonOthThanSal_TDSoth() As String
DefaultTDSonOthThanSal_TDSoth = "0"
End Function
Function DefaultEmployerOrDeductorOrCollectDetl_TDSoth() As String
DefaultEmployerOrDeductorOrCollectDetl_TDSoth = "0"
End Function
Function DefaultAmtPaid_TDSoth() As String
DefaultAmtPaid_TDSoth = "0"
End Function
Function DefaultTotTDSOnAmtPaid_TDSoth() As String
DefaultTotTDSOnAmtPaid_TDSoth = "0"
End Function
Function DefaultClaimOutOfTotTDSOnAmtPaid_TDSoth() As String
DefaultClaimOutOfTotTDSOnAmtPaid_TDSoth = "0"
End Function
Function DefaultTaxPayments_TaxP() As String
DefaultTaxPayments_TaxP = "0"
End Function
Function DefaultTaxPayment_TaxP() As String
DefaultTaxPayment_TaxP = "0"
End Function
Function DefaultNameOfBankAndBranch_TaxP() As String
DefaultNameOfBankAndBranch_TaxP = "0"
End Function
Function DefaultAmt_TaxP() As String
DefaultAmt_TaxP = "0"
End Function
Function DefaultCode001_AIR() As String
DefaultCode001_AIR = "0"
End Function
Function DefaultCode002_AIR() As String
DefaultCode002_AIR = "0"
End Function
Function DefaultCode003_AIR() As String
DefaultCode003_AIR = "0"
End Function
Function DefaultCode004_AIR() As String
DefaultCode004_AIR = "0"
End Function
Function DefaultCode005_AIR() As String
DefaultCode005_AIR = "0"
End Function
Function DefaultCode006_AIR() As String
DefaultCode006_AIR = "0"
End Function
Function DefaultCode007_AIR() As String
DefaultCode007_AIR = "0"
End Function
Function DefaultCode008_AIR() As String
DefaultCode008_AIR = "0"
End Function
Function DefaultTaxExmpIntInc_AIR() As String
DefaultTaxExmpIntInc_AIR = "0"
End Function
Function DefaultReImbFrmGov_Ver() As String
DefaultReImbFrmGov_Ver = "0"
End Function



 ' Start Declaration  For Hide Unhide



 '  End Declaration  For Hide Unhide

Sub intVariables()
     IncomeFromSal_IncD = ""
     IncomeFromHP_IncD = ""
     FamPension_IncD = ""
     IndInterest_IncD = ""
     IncomeFromOS_IncD = ""
     GrossTotIncome_IncD = ""
     Section80C_IncD = ""
     Section80CCC_IncD = ""
     Section80CCD_IncD = ""
     Section80CCF_IncD = ""
     Section80D_IncD = ""
     Section80DD_IncD = ""
     Section80DDB_IncD = ""
     Section80E_IncD = ""
     Section80G_IncD = ""
     Section80GG_IncD = ""
     Section80GGA_IncD = ""
     Section80GGC_IncD = ""
     Section80U_IncD = ""
     TotalChapVIADeductions_IncD = ""
     TotalIncome_IncD = ""
     NetAgriculturalIncome_IncD = ""
     AggregateIncome_IncD = ""
     TaxOnAggregateInc_IncD = ""
     RebateOnAgriInc_IncD = ""
     TotalTaxPayable_IncD = ""
     SurchargeOnTaxPayable_IncD = ""
     EducationCess_IncD = ""
     GrossTaxLiability_IncD = ""
     Section89_IncD = ""
     Section90and91_IncD = ""
     NetTaxLiability_IncD = ""
     IntrstPayUs234A_IncD = ""
     IntrstPayUs234B_IncD = ""
     IntrstPayUs234C_IncD = ""
     TotalIntrstPay_IncD = ""
     TotTaxPlusIntrstPay_IncD = ""
     AdvanceTax_IncD = ""
     TDS_IncD = ""
     SelfAssessmentTax_IncD = ""
     TotalTaxesPaid_IncD = ""
     BalTaxPayable_IncD = ""
     RefundDue_IncD = ""
     BankAccountNumber_IncD = ""
     EcsRequired_IncD = ""
     MICRCode_IncD = ""
     BankAccountType_IncD = ""
     TAN_TDSal = ""
     EmployerOrDeductorOrCollecterName_TDSal = ""
     AddrDetail_TDSal = ""
     CityOrTownOrDistrict_TDSal = ""
     StateCode_TDSal = ""
     PinCode_TDSal = ""
     IncChrgSal_TDSal = ""
     DeductUnderChapVIA_TDSal = ""
     TaxPayIncluSurchEdnCes_TDSal = ""
     TotalTDSSal_TDSal = ""
     TaxPayRefund_TDSal = ""
     TAN_TDSoth = ""
     EmployerOrDeductorOrCollecterName_TDSoth = ""
     AddrDetail_TDSoth = ""
     CityOrTownOrDistrict_TDSoth = ""
     StateCode_TDSoth = ""
     PinCode_TDSoth = ""
     AmtPaid_TDSoth = ""
     DatePayCred_TDSoth = ""
     TotTDSOnAmtPaid_TDSoth = ""
     ClaimOutOfTotTDSOnAmtPaid_TDSoth = ""
     NameOfBank_TaxP = ""
     NameOfBranch_TaxP = ""
     BSRCode_TaxP = ""
     DateDep_TaxP = ""
     SrlNoOfChaln_TaxP = ""
     Amt_TaxP = ""
     Code001_AIR = ""
     Code002_AIR = ""
     Code003_AIR = ""
     Code004_AIR = ""
     Code005_AIR = ""
     Code006_AIR = ""
     Code007_AIR = ""
     Code008_AIR = ""
     TaxExmpIntInc_AIR = ""
     AssesseeVerName_Ver = ""
     FatherName_Ver = ""
     Place_Ver = ""
     Date_Ver = ""
     IdentificationNoOfTRP_Ver = ""
     NameOfTRP_Ver = ""
     ReImbFrmGov_Ver = ""
     msgValidateSheetIncD = ""
     rngname_TDSal = ""
     end_TDSal = ""
     incBy_TDSal = ""
     msgValidateSheetTDSal = ""
     rngname_TDSal2 = ""
     end_TDSal2 = ""
     rngname_TDSoth = ""
     end_TDSoth = ""
     incBy_TDSoth = ""
     msgValidateSheetTDSoth = ""
     rngname_TDSoth2 = ""
     end_TDSoth2 = ""
     rngname_TaxP = ""
     end_TaxP = ""
     incBy_TaxP = ""
     msgValidateSheetTaxP = ""
     rngname_TaxP2 = ""
     end_TaxP2 = ""
     msgValidateSheetAIR = ""
     msgValidateSheetVer = ""

End Sub
Sub Auto_Open()


    With Application
        .Calculation = xlAutomatic
        

    End With
    
Application.EnableEvents = True
'hidegrids
msginit = "Green cells are for data entry" & Chr(13)
msginit = msginit & "CAUTION : DO NOT USE CTRL X OR CUT PASTE WHILE DATAENTRY. " & Chr(13)
msginit = msginit & "Red labels indicate compulsory fields" & Chr(13)
'msginit = msginint & "CLICK THE BUTTON ON TOP LEFT (With Tick Mark) TO CHECK THE ERRORS IN A SHEET" & Chr(13)
MsgBox ("This utility is for Assessment Year 2011-2012")
MsgBox (msginit)
msgint = "Please enable macros to be able to use the excel utility" + Chr(10) + Chr(13)
msgint = msgint + "To enable your macros, please follow these steps (steps may change with Versions of Excel) " + Chr(10) + Chr(13)
msgint = msgint + "Select Tools -> Macros -> Security and Select Low / Medium . Restart Excel" + Chr(10) + Chr(13)
msgint = msgint + "If prompted to enable macros, select Yes and then open the utility."

MsgBox msgint
ThisComponent.CurrentController.setActiveSheet(ThisComponent.Sheets(1-1))

End Sub


Sub validate_xmls()
On Error Resume Next
intVariables
Validateshts
End Sub
Sub Create_XML()
On Error Resume Next
intVariables
Validateshts
Sheet4.Visible = xlSheetVisible
ThisComponent.CurrentController.setActiveSheet(ThisComponent.Sheets(4-1))
msginit21 = Module3.getmsgstate

strpassword = msginit21 + "*"
ThisComponent.CurrentController.getActiveSheet.Unprotect Password:=strpassword
ThisComponent.Sheets(4-1).getCellRangeByName("tds1").Value = UBound(TAN_TDSal)
ThisComponent.Sheets(4-1).getCellRangeByName("tds2").Value = UBound(TAN_TDSoth)
ThisComponent.Sheets(4-1).getCellRangeByName("tp").Value = UBound(BSRCode_TaxP)

ThisComponent.CurrentController.getActiveSheet.Protect Password:=strpassword
MsgBox "To compute Tax and Interest using this utllity, you must click on Compute Tax button and verify the figures before saving the XML. If you have not done so, please do it and then again Generate XML", vbInformation, "Compute Tax"
End Sub
Sub Create_XML_FINAL()
On Error Resume Next
intVariables

stCon = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & ThisWorkbook.FullName & ";" & _
    "Extended Properties=""Excel 8.0;HDR=Yes"";"
    
Dim Msg As Variant
Dim XMLFileName As String

Validateshts

'Total no of entries in the various tabs are summarised. Please cross check before submitting.
 msgrecords = "The following is a summary of the number of entries" & Chr(13)
 msgrecords = msgrecords & "against the various sections of Form 1 submitted by you." & Chr(13)
 msgrecords = msgrecords & "Please verify this summary and press Cancel to go back to correct" & Chr(13)
 msgrecords = msgrecords & "If all summaries are correct, please press OK to generate the XML" & Chr(13) & Chr(13)
 msgrecords = msgrecords & GetRecordCounts
 
 msgrecords1 = GetRecordCounts1
 
' userentry = MsgBox(msgrecords, vbOKCancel, "Do you want to continue generation of XML?")
' If userentry = 2 Then
'' MsgBox ("Please check again your entries, " & Chr(13) & "ensure that all entries per sheet are continuous with no missing 'rows")
' EndProcessing
' End If
 
 'userentry2 = MsgBox(msgrecords1, vbOKCancel, "Do you want to continue rest of XML?")
' If userentry2 = 2 Then
' 'MsgBox ("Please check again your entries, " & Chr(13) & "ensure that all entries per sheet are continuous with no missing 'rows")
' EndProcessing
' End If
'
    PANCODE = PAN_1
    XMLFileName = ThisWorkbook.Path & "\ITR1_" & ThisComponent.Sheets.getByName("Sheet1").Range("sheet1.PAN").value & ".xml"
    
    Open XMLFileName For Output As #1
            XMLHeader
            Form01Header
        
            ConstructSheet1
            ConstructITR1
            Form01Footer
            XMLFooter
        
    Close #1
'MsgBox "This utility is for Assessment Year 2011-2012"
MsgBox "File Saved " & XMLFileName & " Please print acknoledgement from Website"
'PrintWorksheets
End Sub


Function ValidateText_1(strname As Variant, maxsize As Integer) As Boolean
ValidateText_1 = True
If Len(strname) > maxsize Then
   ValidateText_1 = False
End If
End Function

Sub writeXML(line As String)
line = Replace(line, "&", "-")

Print #1, line
End Sub


Function GetRecordCounts1() As String
GetRecordCounts1 = ""
count_5 = countsheet5
GetRecordCounts1 = ""
GetRecordCounts1 = GetRecordCounts1 & " Schedule IT  - Advance Tax and Self Assessment Tax" & count_5 & Chr(13)

End Function

Function XMLHeader()
Print #1, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "ISO-8859-1" & Chr(34) & "?>"
Print #1, "<ITRETURN:ITR xsi:schemaLocation=" & Chr(34) & "http://incometaxindiaefiling.gov.in/main ITRMain11.xsd" & Chr(34) & " xmlns:ITR1FORM=" & Chr(34) & "http://incometaxindiaefiling.gov.in/ITR1" & Chr(34) & " xmlns:ITR2FORM=" & Chr(34) & "http://incometaxindiaefiling.gov.in/ITR2" & Chr(34) & " xmlns:ITR3FORM=" & Chr(34) & "http://incometaxindiaefiling.gov.in/ITR3" & Chr(34) & " xmlns:ITR4FORM=" & Chr(34) & "http://incometaxindiaefiling.gov.in/ITR4" & Chr(34) & " xmlns:ITRETURN=" & Chr(34) & "http://incometaxindiaefiling.gov.in/main" & Chr(34) & " xmlns:ITRForm=" & Chr(34) & "http://incometaxindiaefiling.gov.in/master" & Chr(34) & " xmlns:xsi=" & Chr(34) & "http://www.w3.org/2001/XMLSchema-instance" & Chr(34) & ">"
'Print #1, "<ITR1FORM:ITR1 xsi:schemaLocation=" & Chr(34) & "http://incometaxindiaefiling.gov.in/ITR1 ITR1s09.xsd" & Chr(34) & " xmlns:ITR1FORM=" & Chr(34) & "http://incometaxindiaefiling.gov.in/ITR1" & Chr(34) & " xmlns:ITR2FORM=" & Chr(34) & "http://incometaxindiaefiling.gov.in/ITR2" & Chr(34) & " xmlns:ITR3FORM=" & Chr(34) & "http://incometaxindiaefiling.gov.in/ITR3" & Chr(34) & " xmlns:ITR4FORM=" & Chr(34) & "http://incometaxindiaefiling.gov.in/ITR4" & Chr(34) & " xmlns:ITRETURN=" & Chr(34) & "http://incometaxindiaefiling.gov.in/main" & Chr(34) & " xmlns:ITRForm=" & Chr(34) & "http://incometaxindiaefiling.gov.in/master" & Chr(34) & " xmlns:xsi=" & Chr(34) & "http://www.w3.org/2001/XMLSchema-instance" & Chr(34) & ">"
End Function
Function XMLFooter()
'Print #1, "</ITR1FORM:ITR1>"
Print #1, "</ITRETURN:ITR>"
End Function
Function Form01Header()
Print #1, "<ITR1FORM:ITR1> "
Print #1, "    <ITRForm:CreationInfo>"
'Print #1, "        <ITRForm:SWVersionNo>" & Range("sheet1.SwVersionNo") & "</ITRForm:SWVersionNo>"
Print #1, "        <ITRForm:SWVersionNo>R2</ITRForm:SWVersionNo>"
Print #1, "        <ITRForm:SWCreatedBy>DITEXCEL</ITRForm:SWCreatedBy>"
Print #1, "        <ITRForm:XMLCreatedBy>DIT</ITRForm:XMLCreatedBy>"
Print #1, "        <ITRForm:XMLCreationDate>2011-05-05</ITRForm:XMLCreationDate>"
Print #1, "        <ITRForm:IntermediaryCity>Delhi</ITRForm:IntermediaryCity>"
Print #1, "    </ITRForm:CreationInfo>"
Print #1, "    <ITRForm:Form_ITR1>"
Print #1, "        <ITRForm:FormName>ITR-1</ITRForm:FormName>"
Print #1, "        <ITRForm:Description>For Indls having Income from Salary, Pension, family pension and Interest</ITRForm:Description>"
Print #1, "        <ITRForm:AssessmentYear>2011</ITRForm:AssessmentYear>"
Print #1, "        <ITRForm:SchemaVer>Ver1.0</ITRForm:SchemaVer>"
Print #1, "        <ITRForm:FormVer>Ver1.0</ITRForm:FormVer>"
Print #1, "    </ITRForm:Form_ITR1>"
End Function

Function Form01Footer()
Print #1, "</ITR1FORM:ITR1>"
End Function

Function validateBlank(str As Variant) As Variant
If str = "" Then
    validateBlank = 0
Else
    validateBlank = str
End If
End Function

Sub insertrowstofillformula()
'strpassword = msginit21 + "*"
'ActiveSheet.Unprotect Password:=strPassword
Call InsertRowsAndFillFormulas
ThisComponent.CurrentController.getActiveSheet.Protect Password:=strpassword

End Sub

Sub EndProcessing()
'PrintWorksheets
End
End Sub

Sub EndPrintProcessing()
'PrintWorksheets
End
End Sub

Function CheckAtoZ(chr1) As Boolean
CheckAtoZ = True
If ((Asc(chr1) < 65) Or (Asc(chr1) > 90)) Then
CheckAtoZ = False
End If
End Function

Function CheckDateddmmyyyy(dt As Variant) As Boolean
CheckDateddmmyyyy = True
If Len(dt) > 0 Then
If Mid(dt, 3, 1) <> "/" Then
    If Mid(dt, 3, 1) <> "\" Then
        If Mid(dt, 3, 1) <> "-" Then
            If Mid(dt, 3, 1) <> "." Then
                CheckDateddmmyyyy = False
            Else
            dt = Mid(dt, 1, 2) & "/" & Mid(dt, 4, 7)
            End If
        Else
            dt = Mid(dt, 1, 2) & "/" & Mid(dt, 4, 7)
        End If
    Else
        dt = Mid(dt, 1, 2) & "/" & Mid(dt, 4, 7)
    End If
End If
If Mid(dt, 6, 1) <> "/" Then
    If Mid(dt, 6, 1) <> "-" Then
        If Mid(dt, 6, 1) <> "\" Then
            If Mid(dt, 6, 1) <> "." Then
                CheckDateddmmyyyy = False
            Else
                dt = Mid(dt, 1, 5) & "/" & Mid(dt, 7, 4)
            End If
        Else
            dt = Mid(dt, 1, 5) & "/" & Mid(dt, 7, 4)
        End If
    Else
        dt = Mid(dt, 1, 5) & "/" & Mid(dt, 7, 4)
    End If
End If

If Not IsDate(dt) Then CheckDateddmmyyyy = False
If Val(Mid(dt, 1, 2)) < 0 Then CheckDateddmmyyyy = False
If Val(Mid(dt, 1, 2)) > 31 Then CheckDateddmmyyyy = False
If Val(Mid(dt, 4, 2)) < 0 Then CheckDateddmmyyyy = False
If Val(Mid(dt, 4, 2)) > 12 Then CheckDateddmmyyyy = False
If Val(Mid(dt, 7, 4)) < 1500 Then CheckDateddmmyyyy = False
If Val(Mid(dt, 7, 4)) > 3000 Then CheckDateddmmyyyy = False

End If

End Function

Function ValidatePAN(panentry As String) As Boolean
' arun
ValidatePAN = True
'pan = Range("PAN").Value
If Len(panentry) > 0 Then
If Not IsNumeric(Mid(panentry, 6, 4)) Then
    ValidatePAN = False
    Exit Function
End If
If Not CheckAtoZ(Mid(panentry, 1, 1)) Then
ValidatePAN = False
Exit Function
End If
If Not CheckAtoZ(Mid(panentry, 2, 1)) Then
ValidatePAN = False
Exit Function
End If
If Not CheckAtoZ(Mid(panentry, 3, 1)) Then
ValidatePAN = False
Exit Function
End If
If Not CheckAtoZ(Mid(panentry, 4, 1)) Then
ValidatePAN = False
Exit Function
End If
If Not CheckAtoZ(Mid(panentry, 5, 1)) Then
ValidatePAN = False
Exit Function
End If
If Not CheckAtoZ(Mid(panentry, 10, 1)) Then
ValidatePAN = False
Exit Function
End If
End If

End Function

Function ValidateTantype_text(strname As Variant) As Boolean
ValidateTantype_text = True
Dim len1 As Integer
Dim s1 As String
len1 = Len(strname)
For j = 1 To len1
s1 = Mid(strname, j, 1)
If (((Asc(s1) >= 65) And (Asc(s1) <= 90)) Or (Asc(s1) = 45)) Then
Else
ValidateTantype_text = False
End If
Next
End Function


Function DValidate(ByVal field As Variant, ByVal fieldlabel As String, ByVal sheetlabel As String) As Boolean
'DValidate = True
'Dim arr As Variant
'Dim sheetno As String
'sheetno = ""
'arr = Array("&", """", "'", ">", "<")
'For i = 1 To Len(field)
'    For j = 0 To UBound(arr)
'    If Mid(field, i, 1) = arr(j) Then
'        DValidate = False
'        msgfieldspecialcharacter fieldlabel, sheetlabel, sheetno
'        Exit Function
'    End If
'    Next
'Next
End Function


Function CheckDateMaxDDMMYYYY(ByVal dt As String, ByVal maxday As Integer, ByVal maxmonth As Integer, maxyear As Integer, ByVal errormsg As String)
CheckDateMaxDDMMYYYY = True
If (Val(Mid(dt, 1, 4)) > maxyear) Then
    CheckDateMaxDDMMYYYY = False
    MsgBox "INVALID DATE, " & errormsg
    GoTo exit1
End If
If (Val(Mid(dt, 1, 4)) = maxyear) And (Val(Mid(dt, 6, 2)) > maxmonth) Then
          CheckDateMaxDDMMYYYY = False
          MsgBox "INVALID DATE, " & errormsg
    GoTo exit1
End If

If (Val(Mid(dt, 1, 4)) = maxyear) And (Val(Mid(dt, 6, 2)) = maxmonth) And (Val(Mid(dt, 9, 2)) = maxday) Then
          CheckDateMaxDDMMYYYY = False
          MsgBox "INVALID DATE, " & errormsg
    GoTo exit1
End If

exit1:
End Function

Function CheckDateMinDDMMYYYY(ByVal dt As String, ByVal minday As Integer, ByVal minmonth As Integer, ByVal minyear As Integer, ByVal errormsg As String) As Boolean
CheckDateMinDDMMYYYY = True


If (Val(Mid(dt, 1, 4)) < minyear) Then
    validatedt = False
    CheckDateMinDDMMYYYY = False
    MsgBox "INVALID DATE, " & errormsg
    GoTo exit1
End If

If (Val(Mid(dt, 1, 4)) = minyear) And (Val(Mid(dt, 6, 2)) < minmonth) Then
    CheckDateMinDDMMYYYY = False
    MsgBox "INVALID DATE, " & errormsg
    GoTo exit1
End If
If (Val(Mid(dt, 1, 4)) = minyear) And (Val(Mid(dt, 6, 2)) < minmonth) And (Val(Mid(dt, 9, 2)) < minday) Then
    CheckDateMinDDMMYYYY = False
    MsgBox "INVALID DATE, " & errormsg
    GoTo exit1
End If

exit1:
End Function



Function Dformat(dt As Variant) As String
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
Dformat = formateddate

Else
Dformat = ""
End If
End Function


'from rupam new functions

Function chkCompulsory(field As Variant) As Boolean
chkCompulsory = True
If Len(Trim(field)) <= 0 Then
    chkCompulsory = False
End If
If IsEmpty(field) Then
chkCompulsory = False
End If
End Function

Function chkMaxsize(field As Variant, maxsize As Integer) As Boolean
chkMaxsize = True
If Len(field) > maxsize Then
    chkMaxsize = False
End If
End Function

Function chkMinsize(field As Variant, minsize As Integer) As Boolean
chkMinsize = True
If Len(field) < minsize Then
    chkMinsize = False
End If
End Function

Function chkNumeric(field As Variant) As Boolean
chkNumeric = True
Dim k As Integer
Dim chkchar As String
For k = 1 To Len(field)
chkchar = Mid(field, k, 1)
    If Not IsNumeric(chkchar) Then
        chkNumeric = False
        Exit Function
    End If
Next
End Function

Function chkLong(field As Variant) As Boolean
chkLong = True
Dim k As Integer
Dim chkchar As String
chkchar = Mid(field, 1, 1)
If (chkchar = "-" Or IsNumeric(chkchar)) Then
    For k = 2 To Len(field)
    chkchar = Mid(field, k, 1)
        If Not IsNumeric(chkchar) Then
           chkLong = False
           Exit Function
        End If
    Next
 Else
 chkLong = False
 End If

End Function

Function chkDecimal(field As Variant) As Boolean
Dim k As Integer
Dim chkchar As String
chkDecimal = True
If Not IsNumeric(field) Then
chkDecimal = False
End If
End Function

Function addblock(prevrangename As Variant, blockname As Variant, frmcounter As Variant, totalblocksize As Variant)
Dim i As Integer
Dim newnamerefersto As Variant
Application.EnableEvents = False
msginit21 = Module3.getmsgstate
strpassword = msginit21 + "*"
ThisComponent.CurrentController.getActiveSheet.Unprotect Password:=strpassword

gapbtn = 0
dcounter = 0
Counter = Range(frmcounter).value
If Counter > 0 Then
    
    
    
    destinationrowindex = Range(blockname & Counter).Row
    destinationrowindex = destinationrowindex + gapbtn + totalblocksize
    destinationcolumnindex = Range(blockname & Counter).Column
    
    newnamerefersto = Range(blockname & Counter).Address

    lastdestinationrow = destinationrowindex + totalblocksize - 1
    Cells(CInt(destinationrowindex - 1), CInt(destinationcolumnindex)).Select
    
    Call InsertBlock(totalblocksize)
    
       For i = 1 To Len(newnamerefersto)
        If (Mid(newnamerefersto, i, 1) = "$") Then
             dcounter = dcounter + 1

             If dcounter = 2 Then
                 lenn = (InStr(1, newnamerefersto, ":") - 1) - i
                 startrow = Mid(newnamerefersto, i + 1, lenn)
                 newnamerefersto = Replace(newnamerefersto, startrow, destinationrowindex)
             End If

             If dcounter = 4 Then
                 endrow = Mid(newnamerefersto, i + 1, Len(newnamerefersto) - i)
                 newnamerefersto = Replace(newnamerefersto, endrow, lastdestinationrow)
             End If

        End If
    Next

    Counter = Counter + 1
    ThisWorkbook.Names.add Name:=blockname & Counter, _
            RefersTo:="=" & newnamerefersto, Visible:=True
    Application.EnableEvents = False
    Range(frmcounter).value = Counter
'miki
Application.EnableEvents = False
   msginit21 = Module3.getmsgstate
   strpassword = msginit21 + "*"
   ThisComponent.CurrentController.getActiveSheet.Unprotect Password:=strpassword

    Range(blockname & Counter - 1).Copy Destination:=Cells(destinationrowindex, destinationcolumnindex)
End If

'' add all range name to block

commindex = 2
j = 0

'For i = 1 To Len(prevrangename)
'    If (Mid(prevrangename, i, 1) = ";") Then
'        j = j + 1
'        commindex = i
'    End If
'Next
prevrangename = Split(prevrangename, ";")
rangecount = UBound(prevrangename)

ReDim rangearr(rangecount)
commindex = 0
j = 1

For i = 0 To UBound(prevrangename)
    rangearr(i) = prevrangename(i)
Next

ReDim newranges(rangecount)
For i = 0 To UBound(rangearr)
   newranges(i) = Replace(rangearr(i), CStr(1), CStr(Counter))
Next

If Counter > 2 Then
    ReDim precrange(rangecount)
    For i = 0 To UBound(rangearr)
         precrange(i) = Replace(rangearr(i), CStr(1), CStr(Counter - 1))
    Next

    ReDim oldrangeaddress(rangecount)
    For i = 0 To UBound(precrange)
        If Not (precrange(i) = "") Then
           oldrangeaddress(i) = Range(precrange(i)).Address
        End If
    Next
Else
    ReDim oldrangeaddress(rangecount)
    For i = 0 To UBound(rangearr)
        If Not (rangearr(i) = "") Then
            oldrangeaddress(i) = Range(rangearr(i)).Address
        End If
    Next
End If

ReDim newrangeaddress(rangecount)

For i = 0 To UBound(oldrangeaddress)
    dcounter = 0
    For k = 1 To Len(oldrangeaddress(i))
        If (Mid(oldrangeaddress(i), k, 1) = "$") Then
            dcounter = dcounter + 1
            If (dcounter = 2) Then
            endrow = Mid(oldrangeaddress(i), k + 1, Len(oldrangeaddress(i)) - k)
            newrangeaddress(i) = Replace(oldrangeaddress(i), endrow, (endrow + gapbtn + totalblocksize))
            End If
        End If
    Next
  Next


'For b = 0 To UBound(newrangeaddress)
'
'    If Not newranges(b) = "" Then
'        ThisWorkbook.Names.Add Name:=newranges(b), _
'                    RefersTo:="=" & newrangeaddress(b), Visible:=True
'
'        If (Not Range(newranges(b)).Formula = "") Or IsEmpty(Range(newranges(b)).Formula) Then
'
'           newformula = Replace(Range(newranges(b)).Formula, CStr(counter - 1), CStr(counter))
'           Range(newranges(b)).Formula = newformula 'Application.ConvertFormula(newformula, xlA1)
'        Else
'           Range(newranges(b)).ClearContents
'        End If
'    End If
'Next



For b = 0 To UBound(newrangeaddress)

    If Not newranges(b) = "" Then
        ThisWorkbook.Names.add Name:=newranges(b), _
                    RefersTo:="=" & newrangeaddress(b), Visible:=True
        If Range(newranges(b)).Interior.ColorIndex = 35 Then
            Range(newranges(b)).ClearContents
        End If
    End If
Next
ThisComponent.CurrentController.getActiveSheet.Protect Password:=strpassword
Application.EnableEvents = True
End Function


Function InsertBlock(vRows1 As Variant)
   Dim x As Long
Application.EnableEvents = False
msginit21 = Module3.getmsgstate
strpassword = msginit21 + "*"
ThisComponent.CurrentController.getActiveSheet.Unprotect Password:=strpassword
'

   
   ActiveCell.EntireRow.Select  'So you do not have to preselect entire row
'   If vRows = 0 Then
'    vRows = Application.InputBox(prompt:= _
'      "Enter the number of rows you want to add below selected cell", Title:="Add Rows below the selected cell", _
'      Default:=1, Type:=1) 'Default for 1 row, type 1 is number
'        If vRows = False Then
'            InsertRowsAndFillFormulas = 0
'            Exit Function
'        End If
'   End If
'
    
    Dim vRows As Long
    vRows = CLng(vRows1)
   Dim sht As Worksheet, shts() As String, i As Integer
   ReDim shts(1 To Worksheets.Application.ActiveWorkbook. _
       Windows(1).SelectedSheets.count)
   i = 0
   For Each sht In _
       Application.ActiveWorkbook.Windows(1).SelectedSheets
    Sheets(sht.Name).Select
    i = i + 1
    shts(i) = sht.Name

    x = Sheets(sht.Name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow. _
     Resize(rowsize:=vRows).Insert Shift:=xlDown

    Selection.AutoFill Selection.Resize( _
     rowsize:=vRows + 1), xlFillDefault

    On Error Resume Next
    
    Selection.Offset(1).Resize(vRows).EntireRow. _
     SpecialCells(xlConstants).ClearContents
   Next sht
   
   Worksheets(shts).Select
   'InsertRowsAndFillFormulas = vRows
   
ThisComponent.CurrentController.getActiveSheet.Protect Password:=strpassword
Application.EnableEvents = True
   
End Function
Function GetIncomeOfHP() As Variant
If ValidateIncomeOfHP_HP Then
GetIncomeOfHP = IncomeOfHP_HP
End If
End Function





Function InsertRowsAndFillFormulas(Optional vRows As Long = 0) As Integer
   
   Dim x As Long
   



   ActiveCell.EntireRow.Select  'So you do not have to preselect entire row
   If vRows = 0 Then
    vRows = Application.InputBox(prompt:= _
      "Enter the number of rows you want to add below selected cell", Title:="Add Rows below the selected cell", _
      Default:=1, Type:=1) 'Default for 1 row, type 1 is number
        If vRows = False Then
            InsertRowsAndFillFormulas = 0
            Exit Function
        End If
   End If
    
  msginit21 = Module3.getmsgstate
strpassword = msginit21 + "*"
ThisComponent.CurrentController.getActiveSheet.Unprotect Password:=strpassword
   
   Dim sht As Worksheet, shts() As String, i As Integer
   ReDim shts(1 To Worksheets.Application.ActiveWorkbook. _
       Windows(1).SelectedSheets.count)
   i = 0
   For Each sht In _
       Application.ActiveWorkbook.Windows(1).SelectedSheets
    Sheets(sht.Name).Select
    i = i + 1
    shts(i) = sht.Name

    x = Sheets(sht.Name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow. _
     Resize(rowsize:=vRows).Insert Shift:=xlDown

    Selection.AutoFill Selection.Resize( _
     rowsize:=vRows + 1), xlFillDefault

    On Error Resume Next
    
Selection.Offset(1).Resize(vRows).EntireRow. _
     SpecialCells(xlCellTypeAllValidation).ClearContents
'
'    Selection.Offset(1).Resize(vRows).EntireRow. _
'     SpecialCells(xlConstants).ClearContents
   Next sht
   Worksheets(shts).Select
   InsertRowsAndFillFormulas = vRows
ThisComponent.CurrentController.getActiveSheet.Protect Password:=strpassword
   
End Function



Sub ExendRangeNameToTable(numberofrows As Integer, rangenamestring As Variant)
Dim i As Integer
    rangenamestring = Split(rangenamestring, ";")
    
    For i = 0 To UBound(rangenamestring) - 1
        
        firstbound = Range(rangenamestring(i)).Address
        TEMP = Split(firstbound, "$")
        upperbound = UBound(TEMP)
        TEMP = TEMP(UBound(TEMP))
        x = CInt(TEMP) + numberofrows
        lastbound = Replace(firstbound, TEMP, x)
        
        If upperbound < 3 Then
            rangeaddress = firstbound & ":" & lastbound
        Else
            rangeaddress = lastbound
        End If
        
        ThisWorkbook.Names.add Name:=rangenamestring(i), _
                 RefersTo:="=" & rangeaddress, Visible:=True
        
        
    Next
   
End Sub




Function ChkMinInclusiveDate(Mininclusive As Variant, Mininclusivedate As Variant) As Boolean
'' both date must be in format yyyy-mm-dd
    ChkMinInclusiveDate = True
    If Len(Mininclusive) > 0 Then
        If Mid(Mininclusive, 1, 4) < Mid(Mininclusivedate, 1, 4) Then
            ChkMinInclusiveDate = False
            Exit Function
        Else
            If Mid(Mininclusive, 1, 4) = Mid(Mininclusivedate, 1, 4) Then
                If (Mid(Mininclusive, 6, 2) < Mid(Mininclusivedate, 6, 2)) Then
                    ChkMinInclusiveDate = False
                    Exit Function
                ElseIf ((Mid(Mininclusive, 6, 2) = Mid(Mininclusivedate, 6, 2))) Then
                    If (Mid(Mininclusive, 9, 2) < Mid(Mininclusivedate, 9, 2)) Then
                        ChkMinInclusiveDate = False
                        Exit Function
                   End If
                End If
            End If
        End If
    End If
  
End Function


Function ChkMaxInclusiveDate(Maxinclusive As Variant, Maxinclusivedate As Variant) As Boolean
'' both date must be in format yyyy-mm-dd
    ChkMaxInclusiveDate = True
    If Len(Maxinclusive) > 0 Then
        If Mid(Maxinclusive, 1, 4) > Mid(Maxinclusivedate, 1, 4) Then
            ChkMaxInclusiveDate = False
            Exit Function
        Else
            If Mid(Maxinclusive, 1, 4) = Mid(Maxinclusivedate, 1, 4) Then
                If (Mid(Maxinclusive, 6, 2) > Mid(Maxinclusivedate, 6, 2)) Then
                    ChkMaxInclusiveDate = False
                    Exit Function
                ElseIf ((Mid(Maxinclusive, 6, 2) = Mid(Maxinclusivedate, 6, 2))) Then
                    If (Mid(Maxinclusive, 9, 2) > Mid(Maxinclusivedate, 9, 2)) Then
                        ChkMaxInclusiveDate = False
                        Exit Function
                   End If
                End If
            End If
        End If
    End If
  
End Function

Function checkfieldspecialcharacter(field As Variant) As Boolean
checkfieldspecialcharacter = True
Dim arr As Variant
Dim i As Integer
arr = Array("&", """", "'", ">", "<")
For i = 1 To Len(field)
    For j = 0 To UBound(arr)
    If Mid(field, i, 1) = arr(j) Then
        checkfieldspecialcharacter = False
        Exit Function
    End If
    Next
Next
End Function

Sub createarr()

Total_Count = 0

For i = 2 To 6
s = i
If Not Sheet23.getCellRangeByName("A" + s).String = "" Then
   Total_Count = Total_Count + 1
End If
Next
Total_Count = Total_Count + 1

ReDim sectionname(Total_Count)
For i = 2 To Total_Count
s = i
sectionname(i - 2) = Sheet4.getCellRangeByName("B" + s).String
Next


ReDim secapplicable(Total_Count)
For i = 2 To Total_Count
s = i
secapplicable(i - 2) = Sheet4.getCellRangeByName("D" + s).String
Next

ReDim secapplicable1(Total_Count)
For i = 2 To Total_Count
s = i
secapplicable1(i - 2) = Sheet4.getCellRangeByName("E" + s).String
Next

End Sub
Sub PrintWorksheets()
On Error Resume Next
Application.ScreenUpdating = False

 'Call createarr
 

        With Sheet1.PageSetup
        
            .BlackAndWhite = True
            .CenterHorizontally = True
            .CenterVertically = False
            .LeftMargin = 0.4
            .RightMargin = 0
            .TopMargin = 0
            .BottomMargin = 0
            .PaperSize = xlPaperA4
            .FitToPagesTall = 1
            .FitToPagesWide = 1
            .Orientation = xlPortrait

        End With
        
        With Sheet2.PageSetup

            .BlackAndWhite = True
            .CenterHorizontally = True
            .CenterVertically = False
            .LeftMargin = 0.4
            .RightMargin = 0
            .TopMargin = 0
            .BottomMargin = 0
            .PaperSize = xlPaperA4
            .Orientation = xlPortrait

        End With
        
        With Sheets("TDS").PageSetup

            .BlackAndWhite = True
            .CenterHorizontally = True
            .CenterVertically = False
            .LeftMargin = 0.4
            .RightMargin = 0
            .TopMargin = 0
            .BottomMargin = 0
            .PaperSize = xlPaperA4
            .Orientation = xlLandscape

        End With
        

 printmsg = MsgBox("Do you want to preview the ITR for printing?", vbOKCancel, "Print preview")
 If printmsg = 2 Then
'EndPrintProcessing
 Else
 'Sheets(Array("Income Details", "Taxes paid and Verification", "TDS")).PrintPreview
 Sheets(Array(2, 3, 4)).PrintPreview
End If

Sheet1.Select
Application.ScreenUpdating = True


'Worksheets.PrintPreview

End Sub

Sub HideSheets()

Application.ScreenUpdating = False

For Each ws In ActiveWorkbook.Worksheets
    ws.Visible = xlSheetVisible
Next

Call createarr
For i = 0 To Total_Count - 2
    If secapplicable(i) = "N" Then
        If Not (sectionname(i) = "80") Then
           ActiveWorkbook.Sheets(sectionname(i)).Visible = xlSheetVeryHidden
        Else
           Sheet16.Visible = xlSheetVeryHidden
        End If
    End If
Next

Application.ScreenUpdating = True

End Sub

'//////end prexml////////

Function Validatesheet1() As Boolean
Validatesheet1 = True
If Not ValidateFirstName_1() Then Validatesheet1 = False
If Not ValidateMiddleName_1() Then Validatesheet1 = False
If Not ValidateSurNameOrOrgName_1() Then Validatesheet1 = False
If Not ValidatePAN_1() Then Validatesheet1 = False
If Not ValidateResidenceNo_1() Then Validatesheet1 = False
If Not ValidateResidenceName_1() Then Validatesheet1 = False
If Not ValidateRoadOrStreet_1() Then Validatesheet1 = False
If Not ValidateLocalityOrArea_1() Then Validatesheet1 = False
If Not ValidateCityOrTownOrDistrict_1() Then Validatesheet1 = False
If Not ValidateStateCode_1() Then Validatesheet1 = False
If Not ValidatePinCode_1() Then Validatesheet1 = False
If Not ValidateSTDcode_1() Then Validatesheet1 = False
If Not ValidatePhoneNo_1() Then Validatesheet1 = False
If Not ValidateMobileNo_1() Then Validatesheet1 = False
If Not ValidateEmailAddress_1() Then Validatesheet1 = False
If Not ValidateDOB_1() Then Validatesheet1 = False
If Not ValidateEmployerCategory_1() Then Validatesheet1 = False
If Not ValidateGender_1() Then Validatesheet1 = False
If Not ValidateDesigOfficerWardorCircle_1() Then Validatesheet1 = False
If Not ValidateReturnFileSec_1() Then Validatesheet1 = False
If Not ValidateReturnType_1() Then Validatesheet1 = False
If Not ValidateReceiptNo_1() Then Validatesheet1 = False
If Not ValidateOrigRetFiledDate_1() Then Validatesheet1 = False
If Not ValidateResidentialStatus_1() Then Validatesheet1 = False
If Not ValidateStatus_1() Then Validatesheet1 = False
'If Not ValidateAsseseeRepFlg_1() Then Validatesheet1 = False
'If Not ValidateRepName_1() Then Validatesheet1 = False
'If Not ValidateRepAddress_1() Then Validatesheet1 = False
'If Not ValidateRepPAN_1() Then Validatesheet1 = False
 If Not ValidateIncomeFromSal_IncD() Then Validatesheet1 = False
 If Not ValidateIncomeFromHP_IncD() Then Validatesheet1 = False
 'If (Val(Sheet2.Range("IncD.FamPension")) > 0) Then
     'If Not ValidateFamPension_IncD() Then Validatesheet1 = False
     'If Not ValidateIndInterest_IncD() Then Validatesheet1 = False
     If Not ValidateIncomeFromOS_IncD() Then Validatesheet1 = False
 'End If
     If Not ValidateGrossTotIncome_IncD() Then Validatesheet1 = False
     If Not ValidateSection80C_IncD() Then Validatesheet1 = False
     If Not ValidateSection80CCC_IncD() Then Validatesheet1 = False
     If Not ValidateSection80CCD_IncD() Then Validatesheet1 = False
     If Not ValidateSection80CCF_IncD() Then Validatesheet1 = False
     If Not ValidateSection80D_IncD() Then Validatesheet1 = False
     If Not ValidateSection80DD_IncD() Then Validatesheet1 = False
     If Not ValidateSection80DDB_IncD() Then Validatesheet1 = False
     If Not ValidateSection80E_IncD() Then Validatesheet1 = False
     If Not ValidateSection80G_IncD() Then Validatesheet1 = False
     If Not ValidateSection80GG_IncD() Then Validatesheet1 = False
     If Not ValidateSection80GGA_IncD() Then Validatesheet1 = False
     If Not ValidateSection80GGC_IncD() Then Validatesheet1 = False
     If Not ValidateSection80U_IncD() Then Validatesheet1 = False
     If Not ValidateTotalChapVIADeductions_IncD() Then Validatesheet1 = False
     If Not ValidateTotalIncome_IncD() Then Validatesheet1 = False
     If Not ValidateNetAgriculturalIncome_IncD() Then Validatesheet1 = False
     If Not ValidateAggregateIncome_IncD() Then Validatesheet1 = False
     If Not ValidateTaxOnAggregateInc_IncD() Then Validatesheet1 = False
     If Not ValidateRebateOnAgriInc_IncD() Then Validatesheet1 = False
     If Not ValidateTotalTaxPayable_IncD() Then Validatesheet1 = False
     If Not ValidateSurchargeOnTaxPayable_IncD() Then Validatesheet1 = False
     If Not ValidateEducationCess_IncD() Then Validatesheet1 = False
     If Not ValidateGrossTaxLiability_IncD() Then Validatesheet1 = False
     If Not ValidateSection89_IncD() Then Validatesheet1 = False
     If Not ValidateSection90and91_IncD() Then Validatesheet1 = False
     If Not ValidateNetTaxLiability_IncD() Then Validatesheet1 = False
     If Not ValidateIntrstPayUs234A_IncD() Then Validatesheet1 = False
     If Not ValidateIntrstPayUs234B_IncD() Then Validatesheet1 = False
     If Not ValidateIntrstPayUs234C_IncD() Then Validatesheet1 = False
     If Not ValidateTotalIntrstPay_IncD() Then Validatesheet1 = False
     If Not ValidateTotTaxPlusIntrstPay_IncD() Then Validatesheet1 = False

End Function

Function ValidateFirstName_1() As Boolean
 
ValidateFirstName_1 = True
 FirstName_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.FirstName").String
If Len(FirstName_1) > 25 Then
 msgbox1 ("First Name in Sheet : Income Details  should not exceed 25 characters ")
ValidateFirstName_1 = False
Exit Function
End If
End Function
 
Function ValidateMiddleName_1() As Boolean
 
ValidateMiddleName_1 = True
 MiddleName_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.MiddleName").String
If Len(MiddleName_1) > 25 Then
 msgbox1 ("MiddleName in Sheet : Income Details  should not exceed 25 characters ")
ValidateMiddleName_1 = False
Exit Function
End If
End Function
 
Function ValidateSurNameOrOrgName_1() As Boolean
 
ValidateSurNameOrOrgName_1 = True
 SurNameOrOrgName_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.SurNameOrOrgName").String
If Len(SurNameOrOrgName_1) > 75 Then
 msgbox1 ("Last Name in Sheet : Income Details  should not exceed 75 characters ")
ValidateSurNameOrOrgName_1 = False
Exit Function
End If
If SurNameOrOrgName_1 = "" Or IsEmpty(SurNameOrOrgName_1) Then
msgbox1 ("Last Name in Sheet : Income Details  is Compulsory")
ValidateSurNameOrOrgName_1 = False
Exit Function
End If
End Function
 
 
Function ValidatePAN_1() As Boolean
ValidatePAN_1 = True
 PAN_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.PAN").String
Dim tempPAN As String
tempPAN = PAN_1
If Len(PAN_1) > 10 Then
msgbox1 ("PAN in Sheet : Income Details  should be 10 digits")
ValidatePAN_1 = False
Exit Function
End If
If PAN_1 = "" Or IsEmpty(PAN_1) Then
  msgbox1 ("PAN in Sheet : Income Details  is Compulsory")
ValidatePAN_1 = False
Exit Function
End If
If Not ValidatePAN(tempPAN) Then
  msgbox1 ("PAN in Sheet : Income Details  is invalid (10 digits valid PAN)")
ValidatePAN_1 = False
Exit Function
End If
 
If (Len(PAN_1) = 0) Then
  msgbox1 ("PAN in Sheet : Income Details  is Compulsory")
ValidatePAN_1 = False
Exit Function
End If
 
 
End Function

Function ValidateResidenceNo_1() As Boolean
 
ValidateResidenceNo_1 = True
 ResidenceNo_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.ResidenceNo").String
If Len(ResidenceNo_1) > 75 Then
 msgbox1 ("Flat/Door/Block No in Sheet : Income Details  should not exceed 75 characters ")
ValidateResidenceNo_1 = False
Exit Function
End If
If ResidenceNo_1 = "" Or IsEmpty(ResidenceNo_1) Then
msgbox1 ("Flat/Door/Block No in Sheet : Income Details  is Compulsory")
ValidateResidenceNo_1 = False
Exit Function
End If
End Function
 
Function ValidateResidenceName_1() As Boolean
 
ValidateResidenceName_1 = True
End Function
 
Function ValidateRoadOrStreet_1() As Boolean
 
ValidateRoadOrStreet_1 = True
 RoadOrStreet_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.RoadOrStreet").String
If Len(RoadOrStreet_1) > 75 Then
 msgbox1 ("RoadOrStreet in Sheet : Income Details  should not exceed 75 characters ")
ValidateRoadOrStreet_1 = False
Exit Function
End If
End Function
 
Function ValidateLocalityOrArea_1() As Boolean
 
ValidateLocalityOrArea_1 = True
 LocalityOrArea_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.LocalityOrArea").String
If Len(LocalityOrArea_1) > 75 Then
 msgbox1 ("LocalityOrArea in Sheet : Income Details  should not exceed 75 characters ")
ValidateLocalityOrArea_1 = False
Exit Function
End If
If LocalityOrArea_1 = "" Or IsEmpty(LocalityOrArea_1) Then
msgbox1 ("Area / Locality in Sheet : Income Details  is Compulsory")
ValidateLocalityOrArea_1 = False
Exit Function
End If
End Function
 
Function ValidateCityOrTownOrDistrict_1() As Boolean
 
ValidateCityOrTownOrDistrict_1 = True
 CityOrTownOrDistrict_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.CityOrTownOrDistrict").String
If Len(CityOrTownOrDistrict_1) > 75 Then
 msgbox1 ("CityOrTownOrDistrict in Sheet : Income Details  should not exceed 75 characters ")
ValidateCityOrTownOrDistrict_1 = False
Exit Function
End If
If CityOrTownOrDistrict_1 = "" Or IsEmpty(CityOrTownOrDistrict_1) Then
msgbox1 ("City/Town/District in Sheet : Income Details  is Compulsory")
ValidateCityOrTownOrDistrict_1 = False
Exit Function
End If
End Function
 

Function ValidateStateCode_1() As Boolean
ValidateStateCode_1 = True
 StateCode_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.StateCode1").Value
 StateCode_1 = Mid(StateCode_1, 1, 2)
If StateCode_1 = "" Or IsEmpty(StateCode_1) Then
msgbox1 ("State in Sheet : Income Details  is Compulsory")
ValidateStateCode_1 = False

End If
End Function

Function ValidatePinCode_1() As Boolean
ValidatePinCode_1 = True
PinCode_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.PinCode").Value
If Len(PinCode_1) > 6 Then
 msgbox1 ("PinCode in Sheet : Income Details  should be at 6 digits")
ValidatePinCode_1 = False
Exit Function
End If
If PinCode_1 = " " Or IsEmpty(PinCode_1) Then
  msgbox1 ("PinCode in Sheet : Income Details  is Compulsory")
ValidatePinCode_1 = False
Exit Function
End If
For i = 1 To Len(PinCode_1)
If Not IsNumeric(Mid(PinCode_1, i, 1)) Then
  msgbox1 ("PinCode in Sheet : Income Details  must contain only digits from 0 to 9")
ValidatePinCode_1 = False
Exit Function
End If
Next
End Function


Function ValidateSTDcode_1() As Boolean
ValidateSTDcode_1 = True
 STDcode_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.STDcode").Value
If Len(STDcode_1) > 5 Then
 msgbox1 ("STDcode in Sheet : Income Details  should be at most 5 digits")
ValidateSTDcode_1 = False
Exit Function
End If
'If STDcode_1 = " " Or IsEmpty(STDcode_1) Then
'  msgbox1 ("STDcode in Sheet : Income Details  is Compulsory")
'ValidateSTDcode_1 = False
'Exit Function
'End If
If Trim(STDcode_1) <> "" Then
    For i = 1 To Len(STDcode_1)
        If Not IsNumeric(Mid(STDcode_1, i, 1)) Then
            msgbox1 ("STDcode in Sheet : Income Details  must contain only digits from 0 to 9")
            ValidateSTDcode_1 = False
        Exit Function
        End If
    Next
End If

End Function


Function ValidatePhoneNo_1() As Boolean

ValidatePhoneNo_1 = True
 PhoneNo_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.PhoneNo").Value

If Len(PhoneNo_1) > 10 Then
 msgbox1 ("PhoneNo in Sheet : Income Details  should be at most 10 digits")
ValidatePhoneNo_1 = False
Exit Function
End If

If Trim(STDcode_1) <> "" Then
    If Trim(PhoneNo_1) = "" Or IsEmpty(PhoneNo_1) Then
      msgbox1 ("PhoneNo in Sheet : Income Details  is Compulsory")
        ValidatePhoneNo_1 = False
    Exit Function
    End If
End If

If Trim(PhoneNo_1) <> "" Then
    For i = 1 To Len(PhoneNo_1)
        If Not IsNumeric(Mid(PhoneNo_1, i, 1)) Then
          msgbox1 ("PhoneNo in Sheet : Income Details  must contain only digits from 0 to 9")
        ValidatePhoneNo_1 = False
        Exit Function
        End If
    Next
End If

If Trim(PhoneNo_1) <> "" Then
    If Trim(STDcode_1) = "" Or IsEmpty(STDcode_1) Then
      msgbox1 ("STD Code in Sheet : Income Details  is Compulsory if PhoneNo is entered")
        ValidatePhoneNo_1 = False
    Exit Function
    End If
End If

End Function
Function ValidateMobileNo_1() As Boolean

ValidateMobileNo_1 = True
 MobileNo_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.MobileNo").Value

If Trim(MobileNo_1) <> "" Then
If Len(MobileNo_1) <> 10 Then
 msgbox1 ("MobileNo in Sheet : Income Details  should be 10 digits")
ValidateMobileNo_1 = False
Exit Function
End If
End If

If Trim(MobileNo_1) <> "" Then
        If Mid(MobileNo_1, 1, 1) = "0" Then
        msgbox1 ("First digit of MobileNo in Sheet : Income Details  must contain only digits from 1 to 9")
        ValidateMobileNo_1 = False
        Exit Function
        
        End If
End If

If Trim(MobileNo_1) <> "" Then
    For i = 1 To Len(MobileNo_1)
        If Not IsNumeric(Mid(MobileNo_1, i, 1)) Then
          msgbox1 ("MobileNo in Sheet : Income Details  must contain only digits from 0 to 9")
        ValidateMobileNo_1 = False
        Exit Function
        End If
    Next
End If


End Function

Function ValidateEmailAddress_1() As Boolean
 
ValidateEmailAddress_1 = True
 EmailAddress_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.EmailAddress").String
If Len(EmailAddress_1) > 125 Then
 msgbox1 ("EmailAddress in Sheet : Income Details  should not exceed 125 characters ")
ValidateEmailAddress_1 = False
Exit Function
End If

If Not chkCompulsory(EmailAddress_1) Then
    msgbox1 ("Email Address in Sheet : Income Details  is Compulsory")
    ValidateEmailAddress_1 = False
    Exit Function
End If

'If InStr(1, EmailAddress_1, " ") > 0 Then
' msgbox1 ("EmailAddress in Sheet : Income Details  cannot have spaces. Pl correct and re enter")
'ValidateEmailAddress_1 = False
'Exit Function
'End If

If IsEmailValid(EmailAddress_1) Then
Else

 msgbox1 ("EmailAddress in Sheet : Income Details  is invalid")
ValidateEmailAddress_1 = False
Exit Function
End If
End Function
 
                                                                    
Function ValidateDOB_1() As Boolean
ValidateDOB_1 = True
 DOB_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.DOB").String
If DOB_1 = " " Or IsEmpty(DOB_1) Then
  msgbox1 ("Date of birth in Sheet : Income Details  is Compulsory")
    ValidateDOB_1 = False
 Exit Function
End If
If Not CheckDateddmmyyyy(DOB_1) Then
    ValidateDOB_1 = False
  msgbox1 ("Date of birth in Sheet : Income Details  must be a valid dd/mm/yyyy format")
 Exit Function
Else
  DOB_1 = Dformat(ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.DOB").String)
End If
If Not ChkMaxInclusiveDate(DOB_1, "2011-03-31") Then
         msgbox1 ("Date of birth in Sheet : Income Details  must not exceed 31/03/2011")
         ValidateDOB_1 = False
         Exit Function
Else
  DOB_1 = Dformat(ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.DOB").String)
End If

End Function
                                                                    

Function ValidateEmployerCategory_1() As Boolean
ValidateEmployerCategory_1 = True
 EmployerCategory_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.EmployerCategory1").String
If EmployerCategory_1 = "" Or IsEmpty(EmployerCategory_1) Then
msgbox1 ("EmployerCategory in Sheet : Income Details  is Compulsory")
End If
End Function

Function ValidateGender_1() As Boolean
ValidateGender_1 = True
 Gender_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.Gender1").String
 Gender_1 = Mid(Gender_1, 1, 1)
If Gender_1 = "" Or IsEmpty(Gender_1) Then
msgbox1 ("Gender in Sheet : Income Details  is Compulsory")
End If
End Function
Function ValidateDesigOfficerWardorCircle_1() As Boolean
 
ValidateDesigOfficerWardorCircle_1 = True
 DesigOfficerWardorCircle_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.DesigOfficerWardorCircle").String
If Len(DesigOfficerWardorCircle_1) > 40 Then
 msgbox1 ("DesigOfficerWardorCircle in Sheet : Income Details  should not exceed 40 characters ")
ValidateDesigOfficerWardorCircle_1 = False
Exit Function
End If
End Function
 

Function ValidateReturnFileSec_1() As Boolean
ValidateReturnFileSec_1 = True
 ReturnFileSec_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.ReturnFileSec1").String
 ReturnFileSec_1 = Mid(ReturnFileSec_1, 1, 2)
 If ReturnFileSec_1 = "" Or IsEmpty(ReturnFileSec_1) Then
msgbox1 ("ReturnFileSec in Sheet : Income Details  is Compulsory")
End If
End Function

Function ValidateReturnType_1() As Boolean
ValidateReturnType_1 = True
 ReturnType_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.ReturnType1").String
 ReturnType_1 = Mid(ReturnType_1, 1, 1)
If ReturnType_1 = " " Or IsEmpty(ReturnType_1) Then
msgbox1 ("ReturnType in Sheet : Income Details  is Compulsory")
End If

 retfilesec_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.ReturnFileSec1").String
 retfilesec_1 = Mid(ReturnFileSec_1, 1, 2)
 If (retfilesec_1 = "16") Then
    If ReturnType_1 = "O" Then
      msgbox1 ("ReturnType in Sheet : Income Details  must be Revised when Return filed under S.139(5) ")
    End If
Else
    If ReturnType_1 = "R" Then
          msgbox1 ("ReturnType in Sheet : Income Details  must be Original when Return NOT filed under S.139(5) ")
    End If

 End If
 
 

End Function
Function ValidateReceiptNo_1() As Boolean
 
ValidateReceiptNo_1 = True
 ReceiptNo_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.ReceiptNo").String
If Len(ReceiptNo_1) > 15 Then
 msgbox1 ("ReceiptNo in Sheet : Income Details  should be not exceed 15 digits")
ValidateReceiptNo_1 = False
Exit Function
End If
If UCase(ReturnType_1) = "O" Then
    If Len(ReceiptNo_1) > 0 Then
    msgbox1 ("Receipt No in Sheet : Income Details is only required for Revised returns")
    
    End If

End If
If UCase(ReturnType_1) = "R" Then
    If Len(ReceiptNo_1) = 0 Then
    msgbox1 ("Receipt No in Sheet : Income Details is required for Revised returns")
    ValidateReceiptNo_1 = False
    Exit Function
    End If

End If


End Function
 
                                                                    
Function ValidateOrigRetFiledDate_1() As Boolean
ValidateOrigRetFiledDate_1 = True
 OrigRetFiledDate_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.OrigRetFiledDate").String
If Not CheckDateddmmyyyy(OrigRetFiledDate_1) Then
    ValidateOrigRetFiledDate_1 = False
  msgbox1 ("OrigRetFiledDate in Sheet : Income Details  must be a valid dd/mm/yyyy format")
 Exit Function
Else
  OrigRetFiledDate_1 = Dformat(ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.OrigRetFiledDate").String)
 
  If Len(OrigRetFiledDate_1) > 0 Then
  If Not ChkMinInclusiveDate(OrigRetFiledDate_1, "2011-04-01") Then
         msgbox1 ("OrigRetFiledDate in Sheet : Income Details  cannot be less than 01/04/2011")
         ValidateOrigRetFiledDate_1 = False
         Exit Function
        'If Not CheckDateMinDDMMYYYY(OrigRetFiledDate_1, 1, 4, 2010, "Original return filing date cannot be less than 01/04/2010") Then
            'ValidateOrigRetFiledDate_1 = False
          
         'Exit Function
        Else
  OrigRetFiledDate_1 = Dformat(ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.OrigRetFiledDate").String)
        End If
  
  End If
  
End If


If UCase(ReturnType_1) = "O" Then
    If Len(OrigRetFiledDate_1) > 0 Then
    msgbox1 ("Return filed date in Sheet : Income Details is only required for Revised returns")
    ValidateOrigRetFiledDate_1 = False
    Exit Function
    End If

End If

If UCase(ReturnType_1) = "R" Then
    If Len(OrigRetFiledDate_1) = 0 Then
    msgbox1 ("Return filed date in Sheet : Income Details is required for Revised returns")
    End If

End If
End Function
                                                                    

Function ValidateResidentialStatus_1() As Boolean
ValidateResidentialStatus_1 = True
 ResidentialStatus_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.ResidentialStatus1").String
 ResidentialStatus_1 = Mid(ResidentialStatus_1, 1, 3)
If ResidentialStatus_1 = " " Or IsEmpty(ResidentialStatus_1) Then
msgbox1 ("ResidentialStatus in Sheet : Income Details  is Compulsory")
End If
End Function
Function ValidateStatus_1() As Boolean
ValidateStatus_1 = True
 status_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.Status").String
 status_1 = Mid(status_1, 1, 1)
If status_1 = "" Or IsEmpty(status_1) Then
msgbox1 ("Status in Sheet : Income Details  is Compulsory")
End If
End Function

Function ValidateAsseseeRepFlg_1() As Boolean
ValidateAsseseeRepFlg_1 = True
 AsseseeRepFlg_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.AsseseeRepFlg").String
 AsseseeRepFlg_1 = Mid(AsseseeRepFlg_1, 1, 1)
End Function
Function ValidateRepName_1() As Boolean
 
ValidateRepName_1 = True
 RepName_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.RepName").String
If Len(RepName_1) > 50 Then
 msgbox1 ("RepName in Sheet : Income Details  should not exceed 50 characters ")
ValidateRepName_1 = False
Exit Function
End If
If (AsseseeRepFlg_1 = "Y") Then

If RepName_1 = "" Or IsEmpty(RepName_1) Then
msgbox1 ("RepName in Sheet : Income Details  is Compulsory")
ValidateRepName_1 = False
Exit Function
End If

End If

End Function
 
Function ValidateRepAddress_1() As Boolean
 
ValidateRepAddress_1 = True
 RepAddress_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.RepAddress").String
If Len(RepAddress_1) > 75 Then
 msgbox1 ("RepAddress in Sheet : Income Details  should not exceed 75 characters ")
ValidateRepAddress_1 = False
Exit Function
End If
If (AsseseeRepFlg_1 = "Y") Then

If RepAddress_1 = "" Or IsEmpty(RepAddress_1) Then
msgbox1 ("RepAddress in Sheet : Income Details  is Compulsory")
ValidateRepAddress_1 = False
Exit Function
End If
End If

End Function
 
 
Function ValidateRepPAN_1() As Boolean
ValidateRepPAN_1 = True
 RepPAN_1 = ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.RepPAN").String
Dim tempPAN As String
tempPAN = RepPAN_1
If Len(RepPAN_1) > 10 Then
msgbox1 ("RepPAN in Sheet : Income Details  should not exceed 10 characters ")
ValidateRepPAN_1 = False
Exit Function
End If
If Not ValidatePAN(tempPAN) Then
  msgbox1 ("RepPAN in Sheet : Income Details  is invalid (10 digits valid PAN)")
ValidateRepPAN_1 = False
Exit Function
End If
 
If (AsseseeRepFlg_1 = "Y") Then

If (Len(RepPAN_1) = 0) Then
  msgbox1 ("RepPAN in Sheet : Income Details  is Compulsory")
ValidateRepPAN_1 = False
Exit Function
End If
 
End If

End Function

Function DefaultEmployerCategory_1() As String
DefaultEmployerCategory_1 = "OTH"
End Function
Function DefaultGender_1() As String
DefaultGender_1 = "M"
End Function
Function DefaultReturnFileSec_1() As String
DefaultReturnFileSec_1 = "11"
End Function
Function DefaultReturnType_1() As String
DefaultReturnType_1 = "O"
End Function
Function DefaultResidentialStatus_1() As String
DefaultResidentialStatus_1 = "RES"
End Function
Function DefaultStatus_1() As String
DefaultStatus_1 = "I"
End Function
Function DefaultAsseseeRepFlg_1() As String
DefaultAsseseeRepFlg_1 = "N"
End Function

Function ConstructSheet1()
'writeXML "    <ITRForm:PartA_GEN1>"
writeXML "    <ITRForm:PersonalInfo>"
writeXML "         <ITRForm:AssesseeName>"
If FirstName_1 = "" Then
writeXML "              <ITRForm:FirstName/>"
Else
writeXML "              <ITRForm:FirstName>" & UCase(FirstName_1) & "</ITRForm:FirstName>"
End If
 
If MiddleName_1 = "" Then
writeXML "              <ITRForm:MiddleName/>"
Else
writeXML "              <ITRForm:MiddleName>" & UCase(MiddleName_1) & "</ITRForm:MiddleName>"
End If
 
If SurNameOrOrgName_1 = "" Then
writeXML "              <ITRForm:SurNameOrOrgName/>"
Else
writeXML "              <ITRForm:SurNameOrOrgName>" & UCase(SurNameOrOrgName_1) & "</ITRForm:SurNameOrOrgName>"
End If
writeXML "         </ITRForm:AssesseeName>"
 
If PAN_1 = "" Then
writeXML "         <ITRForm:PAN/>"
Else
writeXML "         <ITRForm:PAN>" & UCase(PAN_1) & "</ITRForm:PAN>"
End If
writeXML "         <ITRForm:Address>"
 
If ResidenceNo_1 = "" Then
writeXML "              <ITRForm:ResidenceNo/>"
Else
writeXML "              <ITRForm:ResidenceNo>" & UCase(ResidenceNo_1) & "</ITRForm:ResidenceNo>"
End If
 
'If ResidenceName_1 = "" Then
'writeXML "              <ITRForm:ResidenceName/>"
'Else
'writeXML "              <ITRForm:ResidenceName>" & UCase(ResidenceName_1) & "</ITRForm:ResidenceName>"
'End If
 
If RoadOrStreet_1 = "" Then
writeXML "              <ITRForm:RoadOrStreet/>"
Else
writeXML "              <ITRForm:RoadOrStreet>" & UCase(RoadOrStreet_1) & "</ITRForm:RoadOrStreet>"
End If
 
If LocalityOrArea_1 = "" Then
writeXML "              <ITRForm:LocalityOrArea/>"
Else
writeXML "              <ITRForm:LocalityOrArea>" & UCase(LocalityOrArea_1) & "</ITRForm:LocalityOrArea>"
End If
 
If CityOrTownOrDistrict_1 = "" Then
writeXML "              <ITRForm:CityOrTownOrDistrict/>"
Else
writeXML "              <ITRForm:CityOrTownOrDistrict>" & UCase(CityOrTownOrDistrict_1) & "</ITRForm:CityOrTownOrDistrict>"
End If
 

If StateCode_1 = "" Then
writeXML "              <ITRForm:StateCode/>"
Else
writeXML "              <ITRForm:StateCode>" & UCase(StateCode_1) & "</ITRForm:StateCode>"
End If
 
If PinCode_1 = "" Then
writeXML "              <ITRForm:PinCode/>"
Else
writeXML "              <ITRForm:PinCode>" & UCase(PinCode_1) & "</ITRForm:PinCode>"
End If
 
If (Len(STDcode_1) > 0 Or Len(PhoneNo_1) > 0) Then
writeXML "              <ITRForm:Phone>"


writeXML "                    <ITRForm:STDcode>" & STDcode_1 & "</ITRForm:STDcode>"
writeXML "                    <ITRForm:PhoneNo>" & PhoneNo_1 & "</ITRForm:PhoneNo>"

writeXML "              </ITRForm:Phone>"
End If

 
If (Len(MobileNo_1) > 0) Then
writeXML "                    <ITRForm:MobileNo>" & MobileNo_1 & "</ITRForm:MobileNo>"
End If
 
 
If EmailAddress_1 = "" Then
'writeXML " <ITRForm:EmailAddress/>"
Else
writeXML "            <ITRForm:EmailAddress>" & EmailAddress_1 & "</ITRForm:EmailAddress>"
End If
 
writeXML "         </ITRForm:Address>"
 
If DOB_1 = "" Then
writeXML "         <ITRForm:DOB/>"
Else
writeXML "         <ITRForm:DOB>" & UCase(DOB_1) & "</ITRForm:DOB>"
End If
'xx
If EmployerCategory_1 = "" Then
writeXML "         <ITRForm:EmployerCategory>" & DefaultEmployerCategory_1 & "</ITRForm:EmployerCategory>"
Else
writeXML "         <ITRForm:EmployerCategory>" & UCase(EmployerCategory_1) & "</ITRForm:EmployerCategory>"
End If
 
If Gender_1 = "" Then
writeXML "         <ITRForm:Gender>" & DefaultGender_1 & "</ITRForm:Gender>"
Else
writeXML "         <ITRForm:Gender>" & UCase(Gender_1) & "</ITRForm:Gender>"
End If

If status_1 = "" Then
writeXML "         <ITRForm:Status>" & DefaultStatus_1 & "</ITRForm:Status>"
Else
writeXML "         <ITRForm:Status>" & UCase(status_1) & "</ITRForm:Status>"
End If
writeXML "    </ITRForm:PersonalInfo>"
writeXML "    <ITRForm:FilingStatus>"
 
If DesigOfficerWardorCircle_1 = "" Then
writeXML "         <ITRForm:DesigOfficerWardorCircle/>"
Else
writeXML "         <ITRForm:DesigOfficerWardorCircle>" & UCase(DesigOfficerWardorCircle_1) & "</ITRForm:DesigOfficerWardorCircle>"
End If
 
If ReturnFileSec_1 = "" Then
writeXML "         <ITRForm:ReturnFileSec>" & DefaultReturnFileSec_1 & "</ITRForm:ReturnFileSec>"
Else
writeXML "         <ITRForm:ReturnFileSec>" & UCase(ReturnFileSec_1) & "</ITRForm:ReturnFileSec>"
End If
 
If ReturnType_1 = "" Then
writeXML "         <ITRForm:ReturnType>" & DefaultReturnType_1 & "</ITRForm:ReturnType>"
Else
writeXML "         <ITRForm:ReturnType>" & UCase(ReturnType_1) & "</ITRForm:ReturnType>"
End If
 
If (ReceiptNo_1 = "" Or UCase(ReturnType_1) = "O") Then
'writeXML " <ITRForm:ReceiptNo/>"
Else

writeXML "         <ITRForm:ReceiptNo>" & UCase(ReceiptNo_1) & "</ITRForm:ReceiptNo>"
End If
 
If (OrigRetFiledDate_1 = "" Or UCase(ReturnType_1) = "O") Then
'writeXML " <ITRForm:OrigRetFiledDate/>"
Else
writeXML "         <ITRForm:OrigRetFiledDate>" & UCase(OrigRetFiledDate_1) & "</ITRForm:OrigRetFiledDate>"
End If
 
If ResidentialStatus_1 = "" Then
writeXML "         <ITRForm:ResidentialStatus>" & DefaultResidentialStatus_1 & "</ITRForm:ResidentialStatus>"
Else
writeXML "         <ITRForm:ResidentialStatus>" & UCase(ResidentialStatus_1) & "</ITRForm:ResidentialStatus>"
End If
 
 
'If AsseseeRepFlg_1 = "" Then
'writeXML " <ITRForm:AsseseeRepFlg>" & DefaultAsseseeRepFlg_1 & "</ITRForm:AsseseeRepFlg>"
'Else
'writeXML " <ITRForm:AsseseeRepFlg>" & UCase(AsseseeRepFlg_1) & "</ITRForm:AsseseeRepFlg>"
'End If
'
'If (Mid(AsseseeRepFlg_1, 1, 1) = "Y") Then
'writeXML "            <ITRForm:AssesseeRep>"
'If RepName_1 = "" Then
'writeXML " <ITRForm:RepName/>"
'Else
'writeXML " <ITRForm:RepName>" & UCase(RepName_1) & "</ITRForm:RepName>"
'End If
'
'If RepAddress_1 = "" Then
'writeXML " <ITRForm:RepAddress/>"
'Else
'writeXML " <ITRForm:RepAddress>" & UCase(RepAddress_1) & "</ITRForm:RepAddress>"
'End If
'
'If RepPAN_1 = "" Then
'writeXML " <ITRForm:RepPAN/>"
'Else
'writeXML " <ITRForm:RepPAN>" & UCase(RepPAN_1) & "</ITRForm:RepPAN>"
'End If
'
'
'writeXML "            </ITRForm:AssesseeRep>"
' End If
 
 
 
writeXML "    </ITRForm:FilingStatus>"
'writeXML "     </ITRForm:PartA_GEN1>"

End Function

Function msgbox1(strmsg As String) As String
msgValidateSheet1 = msgValidateSheet1 & strmsg & Chr(13)
End Function

Sub printerrormessage_gen1()
 If Not Validatesheet1 Then
'    If Not Validatesheet1Blanks Then
'    msgValidateSheet1 = msgValidateSheet1 & msgValidateShee12Blanks
'    Else
'    End If
    ThisComponent.CurrentController.setActiveSheet(ThisComponent.Sheets(1-1))
    MsgBox (msgValidateSheet1)
    EndProcessing
Else
    MsgBox (" Sheet is ok ")
End If
End Sub


Function ConstructITR1()
'writeXML  "<?xml version="1.0" encoding="UTF-8"?>
 'writeXML  "<ITR1FORM:ITR1xsi:schemaLocation="http://incometaxindiaefiling.gov.in/main ITRMain.xsd" xmlns:ITR1FORM="http://incometaxindiaefiling.gov.in/ITR4" xmlns:ITRForm="http://incometaxindiaefiling.gov.in/master" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">"
        writeXML "    <ITRForm:ITR1_IncomeDeductions>"
             If IncomeFromSal_IncD <> "" Then
            writeXML "           <ITRForm:IncomeFromSal>" & UCase(IncomeFromSal_IncD) & "</ITRForm:IncomeFromSal>"
             Else
            writeXML "           <ITRForm:IncomeFromSal>" & UCase(DefaultIncomeFromSal_IncD) & "</ITRForm:IncomeFromSal>"
             End If
             
             If IncomeFromHP_IncD <> "" Then
            writeXML "           <ITRForm:TotalIncomeOfHP>" & UCase(IncomeFromHP_IncD) & "</ITRForm:TotalIncomeOfHP>"
             Else
            writeXML "           <ITRForm:TotalIncomeOfHP>" & UCase(DefaultIncomeFromHP_IncD) & "</ITRForm:TotalIncomeOfHP>"
             End If
             
             
            
                 If IncomeFromOS_IncD <> "" Then
                writeXML "               <ITRForm:IncomeOthSrc>" & UCase(IncomeFromOS_IncD) & "</ITRForm:IncomeOthSrc>"
                 Else
                writeXML "               <ITRForm:IncomeOthSrc>" & UCase(DefaultIncomeFromOS_IncD) & "</ITRForm:IncomeOthSrc>"
                 End If
            
   
             If GrossTotIncome_IncD <> "" Then
            writeXML "           <ITRForm:GrossTotIncome>" & UCase(GrossTotIncome_IncD) & "</ITRForm:GrossTotIncome>"
             Else
            writeXML "           <ITRForm:GrossTotIncome>" & UCase(DefaultGrossTotIncome_IncD) & "</ITRForm:GrossTotIncome>"
             End If
            writeXML "          <ITRForm:DeductUndChapVIA>"
                 If Section80C_IncD <> "" Then
                writeXML "               <ITRForm:Section80C>" & UCase(Section80C_IncD) & "</ITRForm:Section80C>"
                 Else
                writeXML "               <ITRForm:Section80C>" & UCase(DefaultSection80C_IncD) & "</ITRForm:Section80C>"
                 End If
                 If Section80CCC_IncD <> "" Then
                writeXML "               <ITRForm:Section80CCC>" & UCase(Section80CCC_IncD) & "</ITRForm:Section80CCC>"
                 Else
                writeXML "               <ITRForm:Section80CCC>" & UCase(DefaultSection80CCC_IncD) & "</ITRForm:Section80CCC>"
                 End If
                 If Section80CCD_IncD <> "" Then
                writeXML "               <ITRForm:Section80CCD>" & UCase(Section80CCD_IncD) & "</ITRForm:Section80CCD>"
                 Else
                writeXML "               <ITRForm:Section80CCD>" & UCase(DefaultSection80CCD_IncD) & "</ITRForm:Section80CCD>"
                 End If
                 If Section80CCF_IncD <> "" Then
                writeXML "               <ITRForm:Section80CCF>" & UCase(Section80CCF_IncD) & "</ITRForm:Section80CCF>"
                 Else
                writeXML "               <ITRForm:Section80CCF>" & UCase(DefaultSection80CCF_IncD) & "</ITRForm:Section80CCF>"
                 End If
                 
                 If Section80D_IncD <> "" Then
                writeXML "               <ITRForm:Section80D>" & UCase(Section80D_IncD) & "</ITRForm:Section80D>"
                 Else
                writeXML "               <ITRForm:Section80D>" & UCase(DefaultSection80D_IncD) & "</ITRForm:Section80D>"
                 End If
                 If Section80DD_IncD <> "" Then
                writeXML "               <ITRForm:Section80DD>" & UCase(Section80DD_IncD) & "</ITRForm:Section80DD>"
                 Else
                writeXML "               <ITRForm:Section80DD>" & UCase(DefaultSection80DD_IncD) & "</ITRForm:Section80DD>"
                 End If
                 If Section80DDB_IncD <> "" Then
                writeXML "               <ITRForm:Section80DDB>" & UCase(Section80DDB_IncD) & "</ITRForm:Section80DDB>"
                 Else
                writeXML "               <ITRForm:Section80DDB>" & UCase(DefaultSection80DDB_IncD) & "</ITRForm:Section80DDB>"
                 End If
                 If Section80E_IncD <> "" Then
                writeXML "               <ITRForm:Section80E>" & UCase(Section80E_IncD) & "</ITRForm:Section80E>"
                 Else
                writeXML "               <ITRForm:Section80E>" & UCase(DefaultSection80E_IncD) & "</ITRForm:Section80E>"
                 End If
                 If Section80G_IncD <> "" Then
                writeXML "               <ITRForm:Section80G>" & UCase(Section80G_IncD) & "</ITRForm:Section80G>"
                 Else
                writeXML "               <ITRForm:Section80G>" & UCase(DefaultSection80G_IncD) & "</ITRForm:Section80G>"
                 End If
                 If Section80GG_IncD <> "" Then
                writeXML "               <ITRForm:Section80GG>" & UCase(Section80GG_IncD) & "</ITRForm:Section80GG>"
                 Else
                writeXML "               <ITRForm:Section80GG>" & UCase(DefaultSection80GG_IncD) & "</ITRForm:Section80GG>"
                 End If
                 If Section80GGA_IncD <> "" Then
                writeXML "               <ITRForm:Section80GGA>" & UCase(Section80GGA_IncD) & "</ITRForm:Section80GGA>"
                 Else
                writeXML "               <ITRForm:Section80GGA>" & UCase(DefaultSection80GGA_IncD) & "</ITRForm:Section80GGA>"
                 End If
                 If Section80GGC_IncD <> "" Then
                writeXML "               <ITRForm:Section80GGC>" & UCase(Section80GGC_IncD) & "</ITRForm:Section80GGC>"
                 Else
                writeXML "               <ITRForm:Section80GGC>" & UCase(DefaultSection80GGC_IncD) & "</ITRForm:Section80GGC>"
                 End If
                 If Section80U_IncD <> "" Then
                writeXML "               <ITRForm:Section80U>" & UCase(Section80U_IncD) & "</ITRForm:Section80U>"
                 Else
                writeXML "               <ITRForm:Section80U>" & UCase(DefaultSection80U_IncD) & "</ITRForm:Section80U>"
                 End If
                 If TotalChapVIADeductions_IncD <> "" Then
                writeXML "               <ITRForm:TotalChapVIADeductions>" & UCase(TotalChapVIADeductions_IncD) & "</ITRForm:TotalChapVIADeductions>"
                 Else
                writeXML "               <ITRForm:TotalChapVIADeductions>" & UCase(DefaultTotalChapVIADeductions_IncD) & "</ITRForm:TotalChapVIADeductions>"
                 End If
            writeXML "          </ITRForm:DeductUndChapVIA>"
             If TotalIncome_IncD <> "" Then
            writeXML "           <ITRForm:TotalIncome>" & UCase(TotalIncome_IncD) & "</ITRForm:TotalIncome>"
             Else
            writeXML "           <ITRForm:TotalIncome>" & UCase(DefaultTotalIncome_IncD) & "</ITRForm:TotalIncome>"
             End If
        writeXML "    </ITRForm:ITR1_IncomeDeductions>"
        writeXML "    <ITRForm:ITR1_TaxComputation>"
             If TotalTaxPayable_IncD <> "" Then
            writeXML "           <ITRForm:TotalTaxPayable>" & UCase(TotalTaxPayable_IncD) & "</ITRForm:TotalTaxPayable>"
             Else
            writeXML "           <ITRForm:TotalTaxPayable>" & UCase(DefaultTotalTaxPayable_IncD) & "</ITRForm:TotalTaxPayable>"
             End If
             
            writeXML "           <ITRForm:SurchargeOnTaxPayable>0</ITRForm:SurchargeOnTaxPayable>"
            
             
             If EducationCess_IncD <> "" Then
            writeXML "           <ITRForm:EducationCess>" & UCase(EducationCess_IncD) & "</ITRForm:EducationCess>"
             Else
            writeXML "           <ITRForm:EducationCess>" & UCase(DefaultEducationCess_IncD) & "</ITRForm:EducationCess>"
             End If
             If GrossTaxLiability_IncD <> "" Then
            writeXML "           <ITRForm:GrossTaxLiability>" & UCase(GrossTaxLiability_IncD) & "</ITRForm:GrossTaxLiability>"
             Else
            writeXML "           <ITRForm:GrossTaxLiability>" & UCase(DefaultGrossTaxLiability_IncD) & "</ITRForm:GrossTaxLiability>"
             End If
             If Section89_IncD <> "" Then
            writeXML "           <ITRForm:Section89>" & UCase(Section89_IncD) & "</ITRForm:Section89>"
             Else
            writeXML "           <ITRForm:Section89>" & UCase(DefaultSection89_IncD) & "</ITRForm:Section89>"
             End If
             If Section90and91_IncD <> "" Then
            writeXML "           <ITRForm:Section90and91>" & UCase(Section90and91_IncD) & "</ITRForm:Section90and91>"
             Else
            writeXML "           <ITRForm:Section90and91>" & UCase(DefaultSection90and91_IncD) & "</ITRForm:Section90and91>"
             End If
             If NetTaxLiability_IncD <> "" Then
            writeXML "           <ITRForm:NetTaxLiability>" & UCase(NetTaxLiability_IncD) & "</ITRForm:NetTaxLiability>"
             Else
            writeXML "           <ITRForm:NetTaxLiability>" & UCase(DefaultNetTaxLiability_IncD) & "</ITRForm:NetTaxLiability>"
             End If
                 If TotalIntrstPay_IncD <> "" Then
                writeXML "               <ITRForm:TotalIntrstPay>" & UCase(TotalIntrstPay_IncD) & "</ITRForm:TotalIntrstPay>"
                 Else
                writeXML "               <ITRForm:TotalIntrstPay>" & UCase(DefaultTotalIntrstPay_IncD) & "</ITRForm:TotalIntrstPay>"
                 End If
             If TotTaxPlusIntrstPay_IncD <> "" Then
            writeXML "           <ITRForm:TotTaxPlusIntrstPay>" & UCase(TotTaxPlusIntrstPay_IncD) & "</ITRForm:TotTaxPlusIntrstPay>"
             Else
            writeXML "           <ITRForm:TotTaxPlusIntrstPay>" & UCase(DefaultTotTaxPlusIntrstPay_IncD) & "</ITRForm:TotTaxPlusIntrstPay>"
             End If
        writeXML "    </ITRForm:ITR1_TaxComputation>"
        writeXML "    <ITRForm:TaxPaid>"
            writeXML "          <ITRForm:TaxesPaid>"
                 If AdvanceTax_IncD <> "" Then
                writeXML "               <ITRForm:AdvanceTax>" & UCase(AdvanceTax_IncD) & "</ITRForm:AdvanceTax>"
                 Else
                writeXML "               <ITRForm:AdvanceTax>" & UCase(DefaultAdvanceTax_IncD) & "</ITRForm:AdvanceTax>"
                 End If
                 If TDS_IncD <> "" Then
                writeXML "               <ITRForm:TDS>" & UCase(TDS_IncD) & "</ITRForm:TDS>"
                 Else
                writeXML "               <ITRForm:TDS>" & UCase(DefaultTDS_IncD) & "</ITRForm:TDS>"
                 End If
                 If SelfAssessmentTax_IncD <> "" Then
                writeXML "               <ITRForm:SelfAssessmentTax>" & UCase(SelfAssessmentTax_IncD) & "</ITRForm:SelfAssessmentTax>"
                 Else
                writeXML "               <ITRForm:SelfAssessmentTax>" & UCase(DefaultSelfAssessmentTax_IncD) & "</ITRForm:SelfAssessmentTax>"
                 End If
                 If TotalTaxesPaid_IncD <> "" Then
                writeXML "               <ITRForm:TotalTaxesPaid>" & UCase(TotalTaxesPaid_IncD) & "</ITRForm:TotalTaxesPaid>"
                 Else
                writeXML "               <ITRForm:TotalTaxesPaid>" & UCase(DefaultTotalTaxesPaid_IncD) & "</ITRForm:TotalTaxesPaid>"
                 End If
            writeXML "          </ITRForm:TaxesPaid>"
             If BalTaxPayable_IncD <> "" Then
            writeXML "           <ITRForm:BalTaxPayable>" & UCase(BalTaxPayable_IncD) & "</ITRForm:BalTaxPayable>"
             Else
            writeXML "           <ITRForm:BalTaxPayable>" & UCase(DefaultBalTaxPayable_IncD) & "</ITRForm:BalTaxPayable>"
             End If
        writeXML "    </ITRForm:TaxPaid>"
         If RefundDue_IncD <> "" Then
        writeXML "    <ITRForm:Refund>"
             If RefundDue_IncD <> "" Then
            writeXML "           <ITRForm:RefundDue>" & UCase(RefundDue_IncD) & "</ITRForm:RefundDue>"
             Else
            writeXML "           <ITRForm:RefundDue>" & UCase(DefaultRefundDue_IncD) & "</ITRForm:RefundDue>"
             End If
             If BankAccountNumber_IncD <> "" Then
            writeXML "           <ITRForm:BankAccountNumber>" & UCase(BankAccountNumber_IncD) & "</ITRForm:BankAccountNumber>"
             Else
            writeXML "           <ITRForm:BankAccountNumber/>"
             End If
             If EcsRequired_IncD <> "" Then
            writeXML "           <ITRForm:EcsRequired>" & UCase(EcsRequired_IncD) & "</ITRForm:EcsRequired>"
             Else
            writeXML "           <ITRForm:EcsRequired>" & UCase(DefaultEcsRequired_IncD) & "</ITRForm:EcsRequired>"
             End If
             If EcsRequired_IncD = "Y" Then
            writeXML "          <ITRForm:DepositToBankAccount>"
                 If MICRCode_IncD <> "" Then
                writeXML "               <ITRForm:MICRCode>" & UCase(MICRCode_IncD) & "</ITRForm:MICRCode>"
                 Else
                writeXML "               <ITRForm:MICRCode/>"
                 End If
                 If BankAccountType_IncD <> "" Then
                writeXML "               <ITRForm:BankAccountType>" & UCase(BankAccountType_IncD) & "</ITRForm:BankAccountType>"
                 Else
                writeXML "               <ITRForm:BankAccountType/>"
                 End If
            writeXML "          </ITRForm:DepositToBankAccount>"
          End If
        writeXML "    </ITRForm:Refund>"
         End If
         If Not IsEmpty(TAN_TDSal) And UBound(TAN_TDSal) > 0 Then
        writeXML "    <ITRForm:TDSonSalaries>"
 '              TDSonSalary_GenCnt=TDSonSalary_GenCnt+1
             For i = 1 To UBound(TAN_TDSal)
            writeXML "           <ITRForm:TDSonSalary>"
                writeXML "              <ITRForm:EmployerOrDeductorOrCollectDetl>"
                     If TAN_TDSal(i) <> "" Then
                    writeXML "                   <ITRForm:TAN>" & UCase(TAN_TDSal(i)) & "</ITRForm:TAN>"
                     Else
                    writeXML "                   <ITRForm:TAN/>"
                     End If
                    'writeXML "                    <ITRForm:UTN/>"
                     If EmployerOrDeductorOrCollecterName_TDSal(i) <> "" Then
                    writeXML "                   <ITRForm:EmployerOrDeductorOrCollecterName>" & UCase(EmployerOrDeductorOrCollecterName_TDSal(i)) & "</ITRForm:EmployerOrDeductorOrCollecterName>"
                     Else
                    writeXML "                   <ITRForm:EmployerOrDeductorOrCollecterName/>"
                     End If
'                    writeXML "                  <ITRForm:AddressDetail>"
'                         If AddrDetail_TDSal(i) <> "" Then
'                        writeXML "                       <ITRForm:AddrDetail>" & UCase(AddrDetail_TDSal(i)) & "</ITRForm:AddrDetail>"
'                         Else
'                        writeXML "                       <ITRForm:AddrDetail/>"
'                         End If
'                         If CityOrTownOrDistrict_TDSal(i) <> "" Then
'                        writeXML "                       <ITRForm:CityOrTownOrDistrict>" & UCase(CityOrTownOrDistrict_TDSal(i)) & "</ITRForm:CityOrTownOrDistrict>"
'                         Else
'                        writeXML "                       <ITRForm:CityOrTownOrDistrict/>"
'                         End If
'                         If StateCode_TDSal(i) <> "" Then
'                        writeXML "                       <ITRForm:StateCode>" & UCase(StateCode_TDSal(i)) & "</ITRForm:StateCode>"
'                         Else
'                        writeXML "                       <ITRForm:StateCode/>"
'                         End If
'                         If PinCode_TDSal(i) <> "" Then
'                        writeXML "                       <ITRForm:PinCode>" & UCase(PinCode_TDSal(i)) & "</ITRForm:PinCode>"
'                         Else
'                        writeXML "                       <ITRForm:PinCode/>"
'                         End If
'                    writeXML "                  </ITRForm:AddressDetail>"
                writeXML "              </ITRForm:EmployerOrDeductorOrCollectDetl>"
                 If IncChrgSal_TDSal(i) <> "" Then
                writeXML "               <ITRForm:IncChrgSal>" & UCase(IncChrgSal_TDSal(i)) & "</ITRForm:IncChrgSal>"
                 Else
                writeXML "               <ITRForm:IncChrgSal>" & UCase(DefaultIncChrgSal_TDSal) & "</ITRForm:IncChrgSal>"
                 End If
'                 If DeductUnderChapVIA_TDSal(i) <> "" Then
'                writeXML "               <ITRForm:DeductUnderChapVIA>" & UCase(DeductUnderChapVIA_TDSal(i)) & "</ITRForm:DeductUnderChapVIA>"
'                 Else
'                writeXML "               <ITRForm:DeductUnderChapVIA>" & UCase(DefaultDeductUnderChapVIA_TDSal) & "</ITRForm:DeductUnderChapVIA>"
'                 End If
'                 If TaxPayIncluSurchEdnCes_TDSal(i) <> "" Then
'                writeXML "               <ITRForm:TaxPayIncluSurchEdnCes>" & UCase(TaxPayIncluSurchEdnCes_TDSal(i)) & "</ITRForm:TaxPayIncluSurchEdnCes>"
'                 Else
'                writeXML "               <ITRForm:TaxPayIncluSurchEdnCes>" & UCase(DefaultTaxPayIncluSurchEdnCes_TDSal) & "</ITRForm:TaxPayIncluSurchEdnCes>"
'                 End If
                 If TotalTDSSal_TDSal(i) <> "" Then
                writeXML "               <ITRForm:TotalTDSSal>" & UCase(TotalTDSSal_TDSal(i)) & "</ITRForm:TotalTDSSal>"
                 Else
                writeXML "               <ITRForm:TotalTDSSal>" & UCase(DefaultTotalTDSSal_TDSal) & "</ITRForm:TotalTDSSal>"
                 End If
'                 If TaxPayRefund_TDSal(i) <> "" Then
'                writeXML "               <ITRForm:TaxPayRefund>" & UCase(TaxPayRefund_TDSal(i)) & "</ITRForm:TaxPayRefund>"
'                 Else
'                writeXML "               <ITRForm:TaxPayRefund>" & UCase(DefaultTaxPayRefund_TDSal) & "</ITRForm:TaxPayRefund>"
'                 End If
            writeXML "           </ITRForm:TDSonSalary>"
            Next
        writeXML "    </ITRForm:TDSonSalaries>"
        End If
        If Not IsEmpty(TAN_TDSoth) And UBound(TAN_TDSoth) > 0 Then
        writeXML "    <ITRForm:TDSonOthThanSals>"
 '              TDSonOthThanSal_GenCnt=TDSonOthThanSal_GenCnt+1
             For i = 1 To UBound(TAN_TDSoth)
            writeXML "           <ITRForm:TDSonOthThanSal>"
                writeXML "              <ITRForm:EmployerOrDeductorOrCollectDetl>"
                     If TAN_TDSoth(i) <> "" Then
                    writeXML "                   <ITRForm:TAN>" & UCase(TAN_TDSoth(i)) & "</ITRForm:TAN>"
                     Else
                    writeXML "                   <ITRForm:TAN/>"
                     End If
                    'writeXML "                    <ITRForm:UTN/>"
                    
                     If EmployerOrDeductorOrCollecterName_TDSoth(i) <> "" Then
                    writeXML "                   <ITRForm:EmployerOrDeductorOrCollecterName>" & UCase(EmployerOrDeductorOrCollecterName_TDSoth(i)) & "</ITRForm:EmployerOrDeductorOrCollecterName>"
                     Else
                    writeXML "                   <ITRForm:EmployerOrDeductorOrCollecterName/>"
                     End If
'                    writeXML "                  <ITRForm:AddressDetail>"
'                         If AddrDetail_TDSoth(i) <> "" Then
'                        writeXML "                       <ITRForm:AddrDetail>" & UCase(AddrDetail_TDSoth(i)) & "</ITRForm:AddrDetail>"
'                         Else
'                        writeXML "                       <ITRForm:AddrDetail/>"
'                         End If
'                         If CityOrTownOrDistrict_TDSoth(i) <> "" Then
'                        writeXML "                       <ITRForm:CityOrTownOrDistrict>" & UCase(CityOrTownOrDistrict_TDSoth(i)) & "</ITRForm:CityOrTownOrDistrict>"
'                         Else
'                        writeXML "                       <ITRForm:CityOrTownOrDistrict/>"
'                         End If
'                         If StateCode_TDSoth(i) <> "" Then
'                        writeXML "                       <ITRForm:StateCode>" & UCase(StateCode_TDSoth(i)) & "</ITRForm:StateCode>"
'                         Else
'                        writeXML "                       <ITRForm:StateCode/>"
'                         End If
'                         If PinCode_TDSoth(i) <> "" Then
'                        writeXML "                       <ITRForm:PinCode>" & UCase(PinCode_TDSoth(i)) & "</ITRForm:PinCode>"
'                         Else
'                        writeXML "                       <ITRForm:PinCode/>"
'                         End If
'                    writeXML "                  </ITRForm:AddressDetail>"
                writeXML "              </ITRForm:EmployerOrDeductorOrCollectDetl>"
'                 If AmtPaid_TDSoth(i) <> "" Then
'                writeXML "               <ITRForm:AmtPaid>" & UCase(AmtPaid_TDSoth(i)) & "</ITRForm:AmtPaid>"
'                 Else
'                writeXML "               <ITRForm:AmtPaid>" & UCase(DefaultAmtPaid_TDSoth) & "</ITRForm:AmtPaid>"
'                 End If
'                 If DatePayCred_TDSoth(i) <> "" Then
'                writeXML "               <ITRForm:DatePayCred>" & UCase(DatePayCred_TDSoth(i)) & "</ITRForm:DatePayCred>"
'                 Else
'                writeXML "               <ITRForm:DatePayCred/>"
'                 End If
                 If TotTDSOnAmtPaid_TDSoth(i) <> "" Then
                writeXML "               <ITRForm:TotTDSOnAmtPaid>" & UCase(TotTDSOnAmtPaid_TDSoth(i)) & "</ITRForm:TotTDSOnAmtPaid>"
                 Else
                writeXML "               <ITRForm:TotTDSOnAmtPaid>" & UCase(DefaultTotTDSOnAmtPaid_TDSoth) & "</ITRForm:TotTDSOnAmtPaid>"
                 End If
                 If ClaimOutOfTotTDSOnAmtPaid_TDSoth(i) <> "" Then
                writeXML "               <ITRForm:ClaimOutOfTotTDSOnAmtPaid>" & UCase(ClaimOutOfTotTDSOnAmtPaid_TDSoth(i)) & "</ITRForm:ClaimOutOfTotTDSOnAmtPaid>"
                 Else
                writeXML "               <ITRForm:ClaimOutOfTotTDSOnAmtPaid>" & UCase(DefaultClaimOutOfTotTDSOnAmtPaid_TDSoth) & "</ITRForm:ClaimOutOfTotTDSOnAmtPaid>"
                 End If
            writeXML "           </ITRForm:TDSonOthThanSal>"
            Next
        writeXML "    </ITRForm:TDSonOthThanSals>"
        End If
        If Not IsEmpty(BSRCode_TaxP) And UBound(BSRCode_TaxP) > 0 Then
        writeXML "    <ITRForm:TaxPayments>"
 '              TaxPayment_GenCnt=TaxPayment_GenCnt+1
             For i = 1 To UBound(BSRCode_TaxP)
            writeXML "           <ITRForm:TaxPayment>"
'                writeXML "              <ITRForm:NameOfBankAndBranch>"
'                     If NameOfBank_TaxP(i) <> "" Then
'                    writeXML "                   <ITRForm:NameOfBank>" & UCase(NameOfBank_TaxP(i)) & "</ITRForm:NameOfBank>"
'                     End If
'                     If NameOfBranch_TaxP(i) <> "" Then
'                    writeXML "                   <ITRForm:NameOfBranch>" & UCase(NameOfBranch_TaxP(i)) & "</ITRForm:NameOfBranch>"
'                     End If
'                writeXML "              </ITRForm:NameOfBankAndBranch>"
                 If BSRCode_TaxP(i) <> "" Then
                writeXML "               <ITRForm:BSRCode>" & UCase(BSRCode_TaxP(i)) & "</ITRForm:BSRCode>"
                 Else
                writeXML "               <ITRForm:BSRCode/>"
                 End If
                 If DateDep_TaxP(i) <> "" Then
                writeXML "               <ITRForm:DateDep>" & UCase(DateDep_TaxP(i)) & "</ITRForm:DateDep>"
                 Else
                writeXML "               <ITRForm:DateDep/>"
                 End If
                 If SrlNoOfChaln_TaxP(i) <> "" Then
                writeXML "               <ITRForm:SrlNoOfChaln>" & UCase(SrlNoOfChaln_TaxP(i)) & "</ITRForm:SrlNoOfChaln>"
                 Else
                writeXML "               <ITRForm:SrlNoOfChaln/>"
                 End If
                 If Amt_TaxP(i) <> "" Then
                writeXML "               <ITRForm:Amt>" & UCase(Amt_TaxP(i)) & "</ITRForm:Amt>"
                 Else
                writeXML "               <ITRForm:Amt>" & UCase(DefaultAmt_TaxP) & "</ITRForm:Amt>"
                 End If
            writeXML "           </ITRForm:TaxPayment>"
            Next
        writeXML "    </ITRForm:TaxPayments>"
        End If
         If TaxExmpIntInc_AIR <> "" Then
        writeXML "    <ITR1FORM:TaxExmpIntInc>" & UCase(TaxExmpIntInc_AIR) & "</ITR1FORM:TaxExmpIntInc>"
         Else
        writeXML "    <ITR1FORM:TaxExmpIntInc>" & UCase(DefaultTaxExmpIntInc_AIR) & "</ITR1FORM:TaxExmpIntInc>"
         End If
        writeXML "    <ITRForm:Verification>"
            writeXML "          <ITRForm:Declaration>"
                 If AssesseeVerName_Ver <> "" Then
                writeXML "               <ITRForm:AssesseeVerName>" & UCase(AssesseeVerName_Ver) & "</ITRForm:AssesseeVerName>"
                 Else
                writeXML "               <ITRForm:AssesseeVerName/>"
                 End If
                 If FatherName_Ver <> "" Then
                writeXML "               <ITRForm:FatherName>" & UCase(FatherName_Ver) & "</ITRForm:FatherName>"
                 End If
                 If verPAN <> "" Then
                writeXML "               <ITRForm:AssesseeVerPAN>" & UCase(verPAN) & "</ITRForm:AssesseeVerPAN>"
                 End If
            writeXML "          </ITRForm:Declaration>"
             If Place_Ver <> "" Then
            writeXML "           <ITRForm:Place>" & UCase(Place_Ver) & "</ITRForm:Place>"
             Else
            writeXML "           <ITRForm:Place/>"
             End If
             If Date_Ver <> "" Then
            writeXML "           <ITRForm:Date>" & UCase(Date_Ver) & "</ITRForm:Date>"
             Else
            writeXML "           <ITRForm:Date/>"
             End If
        writeXML "    </ITRForm:Verification>"
         If IdentificationNoOfTRP_Ver <> "" Then
        writeXML "    <ITRForm:TaxReturnPreparer>"
             If IdentificationNoOfTRP_Ver <> "" Then
            writeXML "           <ITRForm:IdentificationNoOfTRP>" & UCase(IdentificationNoOfTRP_Ver) & "</ITRForm:IdentificationNoOfTRP>"
             Else
            writeXML "           <ITRForm:IdentificationNoOfTRP/>"
             End If
             If NameOfTRP_Ver <> "" Then
            writeXML "           <ITRForm:NameOfTRP>" & UCase(NameOfTRP_Ver) & "</ITRForm:NameOfTRP>"
             Else
            writeXML "           <ITRForm:NameOfTRP/>"
             End If
             If ReImbFrmGov_Ver <> "" Then
            writeXML "           <ITRForm:ReImbFrmGov>" & UCase(ReImbFrmGov_Ver) & "</ITRForm:ReImbFrmGov>"
             Else
            writeXML "           <ITRForm:ReImbFrmGov>" & UCase(DefaultReImbFrmGov_Ver) & "</ITRForm:ReImbFrmGov>"
             End If
        writeXML "    </ITRForm:TaxReturnPreparer>"
         End If
End Function

Sub trrrrrr()
Application.EnableEvents = True
End Sub
Sub SelectLastRow(rngname As String)
    Dim i As Integer
    rangeaddress = Range(rngname).AddressLocal
    rangeaddress = Replace(rangeaddress, "$", "")
    If InStr(1, rangeaddress, ":") > 0 Then
        rangeaddress = Mid(rangeaddress, InStr(1, rangeaddress, ":") + 1, Len(rangeaddress))
        Range(rangeaddress).Select
    Else
        Range(rangeaddress).Select
    End If
End Sub


Sub filingdate()
Dim todaysdate As String
Dim newfilingdate As String

todaysdate = ThisComponent.Sheets(5-1).getCellRangeByName("DateOfProcessing").String

If Not ValidateDate_Ver() Then
   MsgBox "Please enter Verification Date"
Else
newfilingdate = Date_Ver

If Not ChkMinInclusiveDate(Date_Ver, todaysdate) Then
  newfilingdate = todaysdate
End If
newfilingdate = Mid(newfilingdate, 9, 2) + "/" + Mid(newfilingdate, 6, 2) + "/" + Mid(newfilingdate, 1, 4)

ThisComponent.Sheets(5-1).getCellRangeByName("dateoffiling").String = newfilingdate

End If

End Sub


