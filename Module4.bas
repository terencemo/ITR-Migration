Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
'interest
Dim calcInterestPayable234A As Double
Dim taxbase234A As Double
Dim decVal As Double
Dim delayedInMonths As Double
Dim calcInterestPayable234AR As Integer
Dim taxbase234AR As Double
Dim decValR As Integer
Dim calcIntrstPayable234B As Double
Dim grossAssessedTax As Double
Dim GATR As Double
Dim shortFall As Double
Dim shortFallR As Double
Dim balancePrincipal As Double
Dim balanceInterest As Double
Dim adjustedPrincipal As Double
Dim adjustedInterest As Double
Dim SATPaidAtPeriod As Double
Dim carryForwardPrinicipal As Double
Dim carryForwardInterest As Double
Dim SYSCALCGrossTaxLiability As Double
Dim calcIntrst234BOnPeriod As Double
Dim balancePrincipalR As Double
Dim balanceInterestR As Double
Dim calcIntrst234BOnPeriodR As Double
Dim calcIntrst234BUptoPeriod As Double
Dim assYear As Integer
Dim i As Integer
Dim section90 As Double
Dim totalMatchedAmount As Double
Dim section91 As Double
Dim noOfMonths As Double
Dim endDateTotDayMonths As Integer
Dim startDateTotDayMonths As Integer
Dim dateOfProcessing As String
Dim calcIntrst234C As Double
Dim qtrwiseLTCGProvNotExerTotal As Double
Dim qtrwiseLTCGProvExerTotal As Double
Dim qtrwiseSTCGIncUs111ATotal As Double
Dim qtrwiseLotteryIncUs115BBTotal As Double
Dim partBTILTCGProvNotExerTotal As Double
Dim partBTILTCGProvExerTotal As Double
Dim partBTISTCGUs111ATotal As Double
Dim partBTILotteryIncUs115BBTotal As Double
Dim schCgOsLTCGProvNotExerTotal As Double
Dim schCgOsLTCGProvExerTotal As Double
Dim schCgOsSTCGUs111ATotal As Double
Dim schCgOsLotteryIncUs115BBTotal As Double
Dim schSiLTCGProvNotExerTotal As Double
Dim schSiLTCGProvExerTotal As Double
Dim schSiSTCGUs111ATotal As Double
Dim schSiLotteryIncUs115BBTotal As Double
Dim calcLTCGProvisoNotExer As Double
Dim calcLTCGProvisoExer As Double
Dim calcSTCGIncUs111A As Double
Dim calcLotteryIncUs115BB As Double
Dim LTCGProvNotExerQ1 As Double
Dim LTCGProvExerQ1  As Double
Dim STCGIncUs111AQ1  As Double
Dim LotteryIncUs115BBQ1  As Double
Dim LTCGProvNotExerQ2 As Double
Dim LTCGProvExerQ2  As Double
Dim STCGIncUs111AQ2 As Double
Dim LotteryIncUs115BBQ2 As Double
Dim LTCGProvNotExerQ3 As Double
Dim LTCGProvExerQ3 As Double
Dim STCGIncUs111AQ3 As Double
Dim LotteryIncUs115BBQ3 As Double
Dim LTCGProvNotExerQ4 As Double
Dim LTCGProvExerQ4 As Double
Dim STCGIncUs111AQ4  As Double
Dim LotteryIncUs115BBQ4  As Double
Dim LTCGProvNotExerQ5  As Double
Dim LTCGProvExerQ5 As Double
Dim STCGIncUs111AQ5  As Double
Dim LotteryIncUs115BBQ5 As Double
Dim H1 As Double
Dim H2 As Double
Dim H3 As Double
Dim H4 As Double
Dim H5 As Double
Dim G1 As Double
Dim G2 As Double
Dim G3 As Double
Dim G4 As Double
Dim G5 As Double
Dim section89 As Double
Dim G1Perc As Double
Dim G2Perc As Double
Dim G3Perc As Double
Dim G4Perc As Double
Dim G5Perc As Double
Dim Y1 As Double
Dim Y2 As Double
Dim Y3 As Double
Dim Y4 As Double
Dim Y5 As Double
Dim totalValidTANTDSClaim As Double
Dim totalValidTANTDSClaimForYear As Double
'Dim partBTISTCGUs111ATotal As Double
'Dim partBTILotteryIncUs115BBTotal As Double
'Dim schCgOsLTCGProvNotExerTotal As Double
'Dim schCgOsLTCGProvExerTotal As Double
'Dim schCgOsSTCGUs111ATotal As Double
'Dim schCgOsLotteryIncUs115BBTotal As Double
'Dim schSiLTCGProvNotExerTotal As Double
'Dim schSiLTCGProvExerTotal As Double
'Dim schSiSTCGUs111ATotal As Double
'Dim schSiLotteryIncUs115BBTotal As Double
'Dim calcLTCGProvisoNotExer As Double
'Dim calcLTCGProvisoExer As Double
'Dim calcSTCGIncUs111A As Double
'Dim calcLotteryIncUs115BB As Double
'Dim LTCGProvNotExerQ1 As Double
'Dim LTCGProvExerQ1 As Double
'Dim STCGIncUs111AQ1 As Double
'Dim LotteryIncUs115BBQ1 As Double
'Dim LTCGProvNotExerQ2 As Double
'Dim LTCGProvExerQ2 As Double
'Dim STCGIncUs111AQ2 As Double
'Dim LotteryIncUs115BBQ2 As Double
'Dim LTCGProvNotExerQ3 As Double
'Dim LTCGProvExerQ3 As Double
'Dim STCGIncUs111AQ3 As Double
'Dim LotteryIncUs115BBQ3 As Double
'Dim LTCGProvNotExerQ4 As Double
'Dim LTCGProvExerQ4 As Double
'Dim STCGIncUs111AQ4 As Double
'Dim LotteryIncUs115BBQ4 As Double
'Dim LTCGProvNotExerQ5 As Double
'Dim LTCGProvExerQ5 As Double
'Dim STCGIncUs111AQ5 As Double
'Dim LotteryIncUs115BBQ5 As Double
Dim calcRefundDue As Double
Dim totSATPaid As Double
Dim totalTDSClaimAmount As Double
Dim TDSAmtToBeUsed As Double
Dim InterestComputationType As Object
Dim intrstPayUs234A As Double
Dim intrstPayUs234B As Double
Dim intrstPayUs234C As Double
Dim totalInrstPay As Double
Dim aggTaxInterestLiability As Double
Dim totTaxesPaid As Double
Dim balTaxPayable As Double
Dim refundDue As Double
Dim intrstPayUs234AR As Double
Dim intrstPayUs234BR As Double
Dim intrstPayUs234CR As Double
Const CONST_IntrstPay234A_Percentage As Double = 1
Const CONST_GAT_Limit As Double = 10000
Const CONST_ATP_Limit As Double = 90
Const CONST_Intrst234B_Percentage As Double = 1
Const CONST_Corporate_TaxRate1 As Double = 15
Const CONST_Corporate_TaxRate2 As Double = 45
Const CONST_Corporate_TaxRate3 As Double = 75
Const CONST_Intrst234C_Percentage  As Double = 1
Const CONST_NonCorporate_TaxRate1 As Double = 30
Const CONST_NonCorporate_TaxRate2 As Double = 60
Const CONST_Corporate_Percentage1 As Double = 12
Const CONST_Corporate_Percentage2 As Double = 36
Const CONST_Intrst234C_Fix_Months As Double = 3
Const CONST_Intrst234C_Corporate_Q5_Fix_Months As Double = 0
Const CONST_Intrst234C_NonCorporate_Q5_Fix_Months As Double = 0
Const CONST_Intrst234C_Corporate_Q4_Fix_Months As Double = 1
Const CONST_Intrst234C_NonCorporate_Q4_Fix_Months As Double = 1
Const CONST_Total_TDS_Claim_Amount As Double = 500000
Dim TDSAmtUsed As Double
Dim Total_Valid_TAN_TDS_Claim_Flag As String
Dim Total_Valid_TAN_TDS_Claim_For_Year_Flag As String
Dim Total_Matched_Amount_Flag As String
Dim ReCalc_IntrstRefund_Flag As Boolean
Dim IntrstRef_Counter As Integer
Const CONST_Refund_Recalc_Limit As Double = 25000
Const CONST_WinningFromLotteryIncome_TaxRate As Double = 30
Const CONST_LTCGProviso_TaxRate As Double = 10
Const CONST_LTCGNoProviso_TaxRate As Double = 20
Const CONST_STCG_TaxRate As Double = 10
Dim balanceTaxPayable As Double

Sub intvariable()
TDSAmtUsed = 0
Total_Valid_TAN_TDS_Claim_Flag = "N"
Total_Valid_TAN_TDS_Claim_For_Year_Flag = "N"
Total_Matched_Amount_Flag = "N"
ReCalc_IntrstRefund_Flag = True
IntrstRef_Counter = 1
End Sub

Function monthdiff(startDate As String, endDate As String) As Integer
noofdaysinmonth = 31
monthdiff = 0
startDate = CStr(Dformat1(startDate, "yyyy-mm-dd"))
endDate = CStr(Dformat1(endDate, "yyyy-mm-dd"))
startyear = Mid(startDate, 1, 4)
endyear = Mid(endDate, 1, 4)
startmonth = Mid(startDate, 6, 2)
If (startmonth = 2) Then
 noofdaysinmonth = 28
 
End If
startday = Mid(startDate, 9, 2)
If ((startmonth = 4) Or (startmonth = 6) Or (startmonth = 9) Or (startmonth = 11)) Then
 noofdaysinmonth = 30
 
End If


endmonth = Mid(endDate, 6, 2)
monthdiff = (CInt(endyear) - CInt(startyear)) * 12
monthdiff = monthdiff + (CInt(endmonth) - CInt(startmonth))

If (startDate <> endDate) Then
        
If (CInt(startday) <> noofdaysinmonth) Then
        monthdiff = monthdiff + 1
End If
End If
End Function

Function CalculateDelayedInMonths(startDate As Date, endDate As Date) As Double
      noOfMonths = 0
      If (year(endDate) >= year(startDate)) Then
          noOfMonths = Math.Abs(month(endDate) - month(startDate)) + (year(endDate) - year(startDate)) * 12
          endDateTotDayMonths = month(endDate) + day(endDate)
          startDateTotDayMonths = month(startDate) + day(startDate)

          If ((endDateTotDayMonths > startDateTotDayMonths) And ((endDateTotDayMonths - startDateTotDayMonths) >= 2) And day(endDate) <> day(startDate)) Then
                noOfMonths = noOfMonths + 1
          End If
      End If
   CalculateDelayedInMonths = noOfMonths
End Function

Function Calculate_BalanceTaxPayable(calcTaxInterestLaibility As Double, calcTotTaxesPaid As Double) As Double
        balanceTaxPayable = 0
        If (calcTaxInterestLaibility > calcTotTaxesPaid) Then
                balanceTaxPayable = WorksheetFunction.Max(0, calcTaxInterestLaibility - calcTotTaxesPaid)
        End If
        Calculate_BalanceTaxPayable = balanceTaxPayable
End Function

Function Calculate_InterestPayable234A(SYSCALCGrossTaxLiability As Double, matchedAdvanceTax As Double, TDSAmtUsed As Double, section89 As Double, section90 As Double, section91 As Double) As Double
      calcInterestPayable234A = 0
      taxbase234A = WorksheetFunction.Max(0, SYSCALCGrossTaxLiability - (matchedAdvanceTax + TDSAmtUsed + section89 + section90 + section91))
      taxbase234A = WorksheetFunction.RoundDown(taxbase234A / 100, 0) * 100
      'taxbase234A = taxbase234A / 100
      'decVal = CLng(taxbase234A)
      'taxbase234A = decVal * 100
      'delayedInMonths = monthdiff2(Sheet5.Range("duedate").value, Sheet5.Range("dateoffiling").value)
      delayedInMonths = monthdiff2("01/08/2011", ThisComponent.Sheets(5-1).getCellRangeByName("dateoffiling").value)
      If (delayedInMonths > 0) Then
        calcInterestPayable234A = CONST_IntrstPay234A_Percentage / 100 * taxbase234A * delayedInMonths
      End If
      Calculate_InterestPayable234A = calcInterestPayable234A
End Function

Function Calculate_InterestPayable234AR(intrstcomp As Object) As Double
        calcInterestPayable234AR = 0
        taxbase234AR = WorksheetFunction.Max(0, ITRGrossTaxLiability - (matchedAdvanceTax + TDSAmtUsed + TCSPaid + section89 + section90 + section91 + creditUs115JAA))
        taxbase234AR = taxbase234AR / 100
        decVal = CInt(cintaxbase234AR)
         taxbase234AR = decVal * 100
         delayedInMonths = CalculateDelayedInMonths(ThisComponent.Sheets(5-1).getCellRangeByName("dateoffiling").value, _
            ThisComponent.Sheets(1-1).getCellRangeByName("sheet1.OrigRetFiledDate").value)
         calcInterestPayable234AR = CONST_IntrstPay234A_Percentage / 100 * taxbase234AR * delayedInMonths
         Calculate_InterestPayable234AR = calcInterestPayable234AR
End Function

Function Calculate_InterestPayable234B(intrstPayable234A As Double, intrstPayable234C As Double) As Double
      calcIntrstPayable234B = 0
      Dim matchedAdvanceTax As Double
      
      matchedAdvanceTax = ThisComponent.Sheets(2-1).getCellRangeByName("IncD.AdvanceTax").value
      GAT = ThisComponent.Sheets(5-1).getCellRangeByName("GAT").value
      'GAT - grossAssessedTax = WorksheetFunction.Max(0, SYSCALCGrossTaxLiability - (TDSAmtUsed + intrstPayable234B.interestSupplements.TCSPaid + intrstPayable234B.interestSupplements.section89 + intrstPayable234B.interestSupplements.section90 + intrstPayable234B.interestSupplements.section91 + intrstPayable234B.interestSupplements.creditUs115JAA))
      'GATR = WorksheetFunction.Max(0, intrstPayable234B.ITRInterest.ITRGrossTaxLiability - (TDSAmtUsed + intrstPayable234B.interestSupplements.TCSPaid + intrstPayable234B.interestSupplements.section89 + intrstPayable234B.interestSupplements.section90 + intrstPayable234B.interestSupplements.section91 + intrstPayable234B.interestSupplements.creditUs115JAA))
      GAT = ThisComponent.Sheets(5-1).getCellRangeByName("GATS").value
      If (GAT > CONST_GAT_Limit And matchedAdvanceTax < CONST_ATP_Limit / 100 * GAT) Then
            'shortFall = WorksheetFunction.Max(0, grossAssessedTax - intrstPayable234B.interestSupplements.matchedAdvanceTax)
            'shortFallR = WorksheetFunction.Max(0, GATR - intrstPayable234B.interestSupplements.matchedAdvanceTax)
            shortFall = WorksheetFunction.Max(0, GAT - matchedAdvanceTax)
            shortFall = WorksheetFunction.RoundDown(shortFall / 100, 0) * 100
            'shortFall = shortFall / 100
            'decVal = CLng(shortFall)
            'shortFall = decVal * 100
        
            
          
             If (shortFall = 0) Then
                 Calculate_InterestPayable234B = 0
              Else
                 balancePrincipal = 0
                 balanceInterest = 0
                 adjustedPrincipal = 0
                 adjustedInterest = 0
                 SATPaidAtPeriod = 0
                 carryForwardPrinicipal = 0
                 carryForwardInterest = 0
                 calcIntrst234BOnPeriod = 0
                 balancePrincipalR = 0
                 balanceInterestR = 0
                 calcIntrst234BOnPeriodR = 0
                 calcIntrst234BUptoPeriod = 0
            End If
       dateOfProcessing = ThisComponent.Sheets(5-1).getCellRangeByName("dateoffiling").value
       assYear = 2011
        Dim yrdop As Integer
        Dim mthdop As Integer
        dateOfProcessing = Dformat1(dateOfProcessing, "yyyy-mm-dd")
        
       yrdop = CInt(Mid(dateOfProcessing, 1, 4))
        mthdop = CInt(Mid(dateOfProcessing, 6, 2))
       If (yrdop >= assYear) Then
            calcIntrst234BUptoPeriod = mthdop - 4 + (yrdop - assYear) * 12
       End If

       calcIntrst234BUptoPeriod = calcIntrst234BUptoPeriod + 1
              
       For i = 1 To calcIntrst234BUptoPeriod
             If (i = 1) Then
                    balancePrincipal = shortFall
                    
                    calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = intrstPayable234A + calcIntrst234BOnPeriod + intrstPayable234C
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat1").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
                ElseIf i = 2 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
                        calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat2").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
                ElseIf i = 3 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat3").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               ElseIf i = 4 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat4").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               ElseIf i = 5 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat5").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               ElseIf i = 6 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat6").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               ElseIf i = 7 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat7").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               ElseIf i = 8 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat8").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               ElseIf i = 9 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat9").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               ElseIf i = 10 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat10").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               ElseIf i = 11 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat11").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               
               ElseIf i = 12 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat12").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               ElseIf i = 13 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat13").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               ElseIf i = 14 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat14").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               
               ElseIf i = 15 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat15").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               
               ElseIf i = 16 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat16").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               ElseIf i = 17 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat17").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               ElseIf i = 18 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat18").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               ElseIf i = 19 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat19").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               ElseIf i = 20 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat20").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               ElseIf i = 21 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat21").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               ElseIf i = 22 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat22").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               ElseIf i = 23 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat23").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
               ElseIf i = 24 Then
                        SATPaidAtPeriod = 0
                       balancePrincipal = carryForwardPrinicipal
                                      
               calcIntrst234BOnPeriod = Math.Round(CONST_Intrst234B_Percentage / 100 * balancePrincipal)
                    
                    calcIntrstPayable234B = calcIntrstPayable234B + calcIntrst234BOnPeriod
                    balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                  
                     SATPaidAtPeriod = SATPaidAtPeriod + ThisComponent.Sheets(3-1).getCellRangeByName("Sat24").value
                
                     adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                     adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                     carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                     carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
              
               End If
          
Next
End If
      Calculate_InterestPayable234B = calcIntrstPayable234B
End Function

Function Calculate_InterestPayable234C(SYSCALCGrossTaxLiability As Double, matchedAdvanceTax As Double, TDSAmtUsed As Double, section89 As Double, section90 As Double, section91 As Double) As Double
          calcIntrst234C = 0
          qtrwiseLTCGProvNotExerTotal = 0
          qtrwiseLTCGProvExerTotal = 0
          qtrwiseSTCGIncUs111ATotal = 0
          qtrwiseLotteryIncUs115BBTotal = 0
          partBTILTCGProvNotExerTotal = 0
          partBTILTCGProvExerTotal = 0
          partBTISTCGUs111ATotal = 0
          partBTILotteryIncUs115BBTotal = 0
          schCgOsLTCGProvNotExerTotal = 0
          schCgOsLTCGProvExerTotal = 0
          schCgOsSTCGUs111ATotal = 0
          schCgOsLotteryIncUs115BBTotal = 0
          schSiLTCGProvNotExerTotal = 0
          schSiLTCGProvExerTotal = 0
          schSiSTCGUs111ATotal = 0
          schSiLotteryIncUs115BBTotal = 0
          Dim LTCG_Q1 As Double
          Dim LTCG_Q2 As Double
          Dim LTCG_Q3 As Double
          Dim LTCG_Q4 As Double
          Dim STCG_Q1 As Double
          Dim STCG_Q2 As Double
          Dim STCG_Q3 As Double
          Dim STCG_Q4 As Double
          Dim advanceTaxPaidQ1 As Double
          Dim advanceTaxPaidQ2 As Double
          Dim advanceTaxPaidQ3 As Double
          Dim advanceTaxPaidQ4 As Double
          Dim advanceTaxPaidQ5 As Double
          
          ' ------------------Declaration of Quaterwise Incomes-----------------------

          'If (income <> Null And LTCGIncProvisoExercised <> Null And quaterAmounts <> Null) Then
          ' LTCG Proviso Exercised Quater wise Total
              qtrwiseLTCGProvExerTotal = 0 'WorksheetFunction.Max(0, Math.Round(LTCG_Q1)) + WorksheetFunction.Max(0, Math.Round(LTCG_Q2)) + WorksheetFunction.Max(0, Math.Round(LTCG_Q3)) + WorksheetFunction.Max(0, Math.Round(LTCG_Q4))
          'End If

          'If (income <> Null And STCGUndSec111A <> Null And STCGUndSec111A.quaterAmounts <> Null) Then
          ' STCG Under Section 111A Quater wise Total
              qtrwiseSTCGIncUs111ATotal = WorksheetFunction.Max(0, Math.Round(STCG_Q1)) + WorksheetFunction.Max(0, Math.Round(STCG_Q2)) + WorksheetFunction.Max(0, Math.Round(STCG_Q3)) + WorksheetFunction.Max(0, Math.Round(STCG_Q4))
          'End If

          'If (income <> Null And lotteryIncUs115BB <> Null And quaterAmounts <> Null) Then
         ' Lottery Income Under Section 115BB Quater wise Total
              qtrwiseLotteryIncUs115BBTotal = ThisComponent.Sheets(12-1).getCellRangeByName("os.WinLottRacePuzz").value
              'WorksheetFunction.Max(0, Math.Round(amountQuater1)) + WorksheetFunction.Max(0, Math.Round(amountQuater2)) + WorksheetFunction.Max(0, Math.Round(amountQuater3)) + WorksheetFunction.Max(0, Math.Round(amountQuater4)) + WorksheetFunction.Max(0, Math.Round(amountQuater5))
          'End If

         ' --------------------Declaration of Part B-TI Values---------------------
          'If (income <> Null And LTCGIncProvisoNotExercised <> Null And partBTIandSchCGOSSIAmount <> Null) Then
          ' LTCG Proviso Not Exercised Part B-TI Value
              'partBTILTCGProvNotExerTotal = Math.Round(partBTIAmount)
          'End If

          'If (income <> Null And LTCGIncProvisoExercised <> Null And partBTIandSchCGOSSIAmount <> Null) Then
          ' LTCG Proviso Exercised Part B-TI Value
             ' partBTILTCGProvExerTotal = Math.Round(partBTIAmount)
          'End If

          'If (income <> Null And STCGUndSec111A <> Null And partBTIandSchCGOSSIAmount <> Null) Then
         ' STCG Income under Section 111A Part B-TI Value
              'partBTISTCGUs111ATotal = Math.Round(partBTIAmount)
          'End If

          'If (income <> Null And lotteryIncUs115BB <> Null And partBTIandSchCGOSSIAmount <> Null) Then
         ' Lottery Income under Section 115BB Part B-TI Value
              'partBTILotteryIncUs115BBTotal = Math.Round(partBTIAmount)
          'End If
       
         ' ------------------- Declaration of Schedule CG-OS Values-------------
          'If (income <> Null And LTCGIncProvisoNotExercised <> Null And partBTIandSchCGOSSIAmount <> Null) Then
            ' LTCG Proviso Not Exercised Schedule CG-OS Value
                'schCgOsLTCGProvNotExerTotal = Math.Round(scheduleCGOSAmount)
          'End If

          'If (income <> Null And LTCGIncProvisoExercised <> Null And partBTIandSchCGOSSIAmount <> Null) Then
            ' LTCG Proviso Exercised Schedule CG-OS Value
               'schCgOsLTCGProvExerTotal = Math.Round(scheduleCGOSAmount)
          'End If

         'If (income <> Null And STCGUndSec111A <> Null And partBTIandSchCGOSSIAmount <> Null) Then
            ' STCG Income under Section 111A Schedule CG-OS Value
              'schCgOsSTCGUs111ATotal = Math.Round(scheduleCGOSAmount)
         'End If

'         If (income <> Null And lotteryIncUs115BB <> Null And partBTIandSchCGOSSIAmount <> Null) Then
'            ' Lottery Income under Section 115BB Schedule CG-OS Value
'             schCgOsLotteryIncUs115BBTotal = Math.Round(scheduleCGOSAmount)
'         End If

         ' ------------------- Declaration of Schedule SI Values-------------
'         If (income <> Null And LTCGIncProvisoNotExercised <> Null And partBTIandSchCGOSSIAmount <> Null) Then
'           ' LTCG Proviso Not Exercised Schedule SI Value
'            schSiLTCGProvNotExerTotal = Math.Round(scheduleSIAmount)
'         End If
'
'         If (income <> Null And LTCGIncProvisoExercised <> Null And partBTIandSchCGOSSIAmount <> Null) Then
'          ' LTCG Proviso Exercised Schedule SI Value
'            schSiLTCGProvExerTotal = Math.Round(scheduleSIAmount)
'         End If
'
'         If (income <> Null And STCGUndSec111A <> Null And partBTIandSchCGOSSIAmount <> Null) Then
'          ' STCG Income under Section 111A Schedule SI Value
'            schSiSTCGUs111ATotal = Math.Round(scheduleSIAmount)
'         End If
'
'         If (income <> Null And lotteryIncUs115BB <> Null And partBTIandSchCGOSSIAmount <> Null) Then
'          ' Lottery Income under Section 115BB Schedule SI Value
'            schSiLotteryIncUs115BBTotal = Math.Round(scheduleSIAmount)
'          End If

         ' ---------- Declaration of variable used in the Interest 234C Computation ----------
         
         ' LTCG Proviso Not Exercised variable
'         calcLTCGProvisoNotExer = 0
'         calcLTCGProvisoNotExer = WorksheetFunction.Max(qtrwiseLTCGProvNotExerTotal, partBTILTCGProvNotExerTotal)
'         calcLTCGProvisoNotExer = WorksheetFunction.Max(calcLTCGProvisoNotExer, schCgOsLTCGProvNotExerTotal)
'         calcLTCGProvisoNotExer = WorksheetFunction.Max(calcLTCGProvisoNotExer, schSiLTCGProvNotExerTotal)
            calcLTCGProvisoNotExer = qtrwiseLTCGProvExerTotal
' LTCG Proviso Exercised variable
         
'         calcLTCGProvisoExer = 0
'         calcLTCGProvisoExer = WorksheetFunction.Max(qtrwiseLTCGProvExerTotal, partBTILTCGProvExerTotal)
'         calcLTCGProvisoExer = WorksheetFunction.Max(calcLTCGProvisoExer, schCgOsLTCGProvExerTotal)
'         calcLTCGProvisoExer = WorksheetFunction.Max(calcLTCGProvisoExer, schSiLTCGProvExerTotal)

'         ' STCG Income Under Section 111A Variable
'         calcSTCGIncUs111A = 0
'         calcSTCGIncUs111A = WorksheetFunction.Max(qtrwiseSTCGIncUs111ATotal, partBTISTCGUs111ATotal)
'         calcSTCGIncUs111A = WorksheetFunction.Max(calcSTCGIncUs111A, schCgOsSTCGUs111ATotal)
'         calcSTCGIncUs111A = WorksheetFunction.Max(calcSTCGIncUs111A, schSiSTCGUs111ATotal)
        calcSTCGIncUs111A = qtrwiseSTCGIncUs111ATotal
'         ' Lottery Income Under Section 115BB Variable
'         calcLotteryIncUs115BB = 0
'         calcLotteryIncUs115BB = WorksheetFunction.Max(qtrwiseLotteryIncUs115BBTotal, partBTILotteryIncUs115BBTotal)
'         calcLotteryIncUs115BB = WorksheetFunction.Max(calcLotteryIncUs115BB, schCgOsLotteryIncUs115BBTotal)
'         calcLotteryIncUs115BB = WorksheetFunction.Max(calcLotteryIncUs115BB, schSiLotteryIncUs115BBTotal)

        calcLotteryIncUs115BB = qtrwiseLotteryIncUs115BBTotal
         ' ------ If quater wise breakup is not given then do the following -----------
         ' Corporate First Quater Variable Declarations
         LTCGProvNotExerQ1 = 0
         LTCGProvExerQ1 = 0
         STCGIncUs111AQ1 = 0
         LotteryIncUs115BBQ1 = 0
         LTCGProvNotExerQ2 = 0
         LTCGProvExerQ2 = 0
         STCGIncUs111AQ2 = 0
         LotteryIncUs115BBQ2 = 0
         LTCGProvNotExerQ3 = 0
         LTCGProvExerQ3 = 0
         STCGIncUs111AQ3 = 0
         LotteryIncUs115BBQ3 = 0
         LTCGProvNotExerQ4 = 0
         LTCGProvExerQ4 = 0
         STCGIncUs111AQ4 = 0
         LotteryIncUs115BBQ4 = 0
         LTCGProvNotExerQ5 = 0
         LTCGProvExerQ5 = 0
         STCGIncUs111AQ5 = 0
         LotteryIncUs115BBQ5 = 0
            
          ' Non-Corporate Second Quater Variable Declarations
              LTCGProvNotExerQ2 = Math.Round(ThisComponent.Sheets(12-1).getCellRangeByName("LTCGAssNo.BalLTCGNo112").value)

              LTCGProvExerQ2 = Math.Round(ThisComponent.Sheets(12-1).getCellRangeByName("LTCG.BalLTCG112").value)

              STCGIncUs111AQ2 = Math.Round(ThisComponent.Sheets.getByName("Sheet6").getCellRangeByName("Sheet8b.TotalShortTerm").value)

              LotteryIncUs115BBQ2 = Math.Round(ThisComponent.Sheets(12-1).getCellRangeByName(" os.WinLottRacePuzz").value)

'              LTCGProvNotExerQ3 = Math.Round(WorksheetFunction.Max(0, amountQuater3))
'              LTCGProvNotExerQ4 = Math.Round(WorksheetFunction.Max(0, amountQuater4))
'              LTCGProvNotExerQ5 = Math.Round(WorksheetFunction.Max(0, amountQuater5))
'
'             LTCGProvExerQ3 = Math.Round(WorksheetFunction.Max(0, amountQuater3))
'             LTCGProvExerQ4 = Math.Round(WorksheetFunction.Max(0, amountQuater4))
'             LTCGProvExerQ5 = Math.Round(WorksheetFunction.Max(0, amountQuater5))
'
             STCGIncUs111AQ3 = Math.Round(WorksheetFunction.Max(0, amountQuater3))
             STCGIncUs111AQ4 = Math.Round(WorksheetFunction.Max(0, amountQuater4))
             'STCGIncUs111AQ5 = Math.Round(WorksheetFunction.Max(0, amountQuater5))
'
'
'            LotteryIncUs115BBQ3 = Math.Round(WorksheetFunction.Max(0, amountQuater3))
'            LotteryIncUs115BBQ4 = Math.Round(WorksheetFunction.Max(0, amountQuater4))
'            LotteryIncUs115BBQ5 = Math.Round(WorksheetFunction.Max(0, amountQuater5))

            LTCGProvNotExerQ1 = calcLTCGProvisoNotExer
         
            LTCGProvExerQ1 = calcLTCGProvisoExer

              STCGIncUs111AQ1 = calcSTCGIncUs111A

                LotteryIncUs115BBQ1 = calcLotteryIncUs115BB

         ' Checking for Corporate or Non-Corporate
         ' ---- If quater wise breakup does not tally with overall total and is less than overall total then the residual value will be added to First Quater Total --------
                LTCGProvNotExerQ1 = LTCGProvNotExerQ1 + calcLTCGProvisoNotExer - qtrwiseLTCGProvNotExerTotal

                LTCGProvExerQ1 = LTCGProvExerQ1 + calcLTCGProvisoExer - qtrwiseLTCGProvExerTotal

                STCGIncUs111AQ1 = STCGIncUs111AQ1 + calcSTCGIncUs111A - qtrwiseSTCGIncUs111ATotal

                LotteryIncUs115BBQ1 = LotteryIncUs115BBQ1 + calcLotteryIncUs115BB - qtrwiseLotteryIncUs115BBTotal

           ' -------------- If the quater wise breakup does not tally with overall total and is greater than overall total then the quater wise  breakup as reported is taken for computation ----------------------
               'LTCGProvNotExerQ1 = Math.Round(amountQuater1)

                'LTCGProvExerQ1 = Math.Round(amountQuater1)

                STCGIncUs111AQ1 = Math.Round(ThisComponent.Sheets(12-1).getCellRangeByName("AccSTCG.Upto15Of9").value)

                'LotteryIncUs115BBQ1 = Math.Round(amountQuater1)
          
          ' ---- If quater wise breakup does not tally with overall total and is less than overall total then the residual value will be added to First Quater Total --------
                LTCGProvNotExerQ2 = LTCGProvNotExerQ2 + calcLTCGProvisoNotExer - qtrwiseLTCGProvNotExerTotal

                LTCGProvExerQ2 = LTCGProvExerQ2 + calcLTCGProvisoExer - qtrwiseLTCGProvExerTotal

                STCGIncUs111AQ2 = STCGIncUs111AQ2 + calcSTCGIncUs111A - qtrwiseSTCGIncUs111ATotal

                LotteryIncUs115BBQ2 = LotteryIncUs115BBQ2 + calcLotteryIncUs115BB - qtrwiseLotteryIncUs115BBTotal

            advanceTaxPaidQ1 = ThisComponent.Sheets(3-1).getCellRangeByName("Qtr1").value
            advanceTaxPaidQ2 = ThisComponent.Sheets(3-1).getCellRangeByName("Qtr2").value
            advanceTaxPaidQ3 = ThisComponent.Sheets(3-1).getCellRangeByName("Qtr3").value
            advanceTaxPaidQ4 = ThisComponent.Sheets(3-1).getCellRangeByName("Qtr4").value
            advanceTaxPaidQ5 = ThisComponent.Sheets(3-1).getCellRangeByName("Qtr5").value

          ' Calculation Gross Assessed Tax (GAT).
            grossAssessedTax = WorksheetFunction.Max(0, SYSCALCGrossTaxLiability - (TDSAmtUsed + section89 + section90 + section91 + Math.Round(calcLTCGProvisoNotExer * CONST_LTCGNoProviso_TaxRate / 100) + Math.Round(calcLTCGProvisoExer * CONST_LTCGProviso_TaxRate / 100) + Math.Round(calcSTCGIncUs111A * CONST_STCG_TaxRate / 100) + Math.Round(calcLotteryIncUs115BB * CONST_WinningFromLotteryIncome_TaxRate / 100)))
          ' Rounding down to the nearest hundred rupees.
         grossAssessedTax = grossAssessedTax / 100
         decVal = CLng(grossAssessedTax)
         grossAssessedTax = decVal * 100
         ' Checking GAT limit for calculating interest payable 234C.
         If (grossAssessedTax > CONST_GAT_Limit) Then
               ' Assigining Advance tAx paid in quater 1 to 5 to the variables.
               H1 = advanceTaxPaidQ1
               H2 = advanceTaxPaidQ2
               H3 = advanceTaxPaidQ3
               H4 = advanceTaxPaidQ4
               H5 = advanceTaxPaidQ5
               G1 = grossAssessedTax
               G2 = G1 + Math.Round(LTCGProvNotExerQ1 * CONST_LTCGNoProviso_TaxRate / 100) + Math.Round(LTCGProvExerQ1 * CONST_LTCGProviso_TaxRate / 100) + Math.Round(STCGIncUs111AQ1 * CONST_STCG_TaxRate / 100) + Math.Round(LotteryIncUs115BBQ1 * CONST_WinningFromLotteryIncome_TaxRate / 100)
               G3 = G2 + Math.Round(LTCGProvNotExerQ2 * CONST_LTCGNoProviso_TaxRate / 100) + Math.Round(LTCGProvExerQ2 * CONST_LTCGProviso_TaxRate / 100) + Math.Round(STCGIncUs111AQ2 * CONST_STCG_TaxRate / 100) + Math.Round(LotteryIncUs115BBQ2 * CONST_WinningFromLotteryIncome_TaxRate / 100)
               G4 = G3 + Math.Round(LTCGProvNotExerQ3 * CONST_LTCGNoProviso_TaxRate / 100) + Math.Round(LTCGProvExerQ3 * CONST_LTCGProviso_TaxRate / 100) + Math.Round(STCGIncUs111AQ3 * CONST_STCG_TaxRate / 100) + Math.Round(LotteryIncUs115BBQ3 * CONST_WinningFromLotteryIncome_TaxRate / 100)
               G5 = grossAssessedTax + Math.Round(calcLTCGProvisoNotExer * CONST_LTCGNoProviso_TaxRate / 100) + Math.Round(calcLTCGProvisoExer * CONST_LTCGProviso_TaxRate / 100) + Math.Round(calcSTCGIncUs111A * CONST_STCG_TaxRate / 100) + Math.Round(calcLotteryIncUs115BB * CONST_WinningFromLotteryIncome_TaxRate / 100)
               G1Perc = 1
               G2Perc = 1
               G3Perc = 1
               G4Perc = 1
               G5Perc = 1

                      G1Perc = 0
                      G2Perc = CONST_NonCorporate_TaxRate1 / 100
                      G3Perc = CONST_NonCorporate_TaxRate2 / 100

                  G1 = G1Perc * G1
                  G2 = G2Perc * G2
                  G3 = G3Perc * G3
                  G4 = G4Perc * G4
                  G5 = G5Perc * G5

                  Y1 = 0
                  Y2 = 0
                  Y3 = 0
                  Y4 = 0
                  Y5 = 0

                   Y1 = WorksheetFunction.Max(0, Math.Round(CONST_Intrst234C_Fix_Months * CONST_Intrst234C_Percentage / 100 * (G1 - H1)))
                   Y2 = WorksheetFunction.Max(0, Math.Round(CONST_Intrst234C_Fix_Months * CONST_Intrst234C_Percentage / 100 * (G2 - (H1 + H2))))
                   Y3 = WorksheetFunction.Max(0, Math.Round(CONST_Intrst234C_Fix_Months * CONST_Intrst234C_Percentage / 100 * (G3 - (H1 + H2 + H3))))
                   Y4 = WorksheetFunction.Max(0, Math.Round(CONST_Intrst234C_NonCorporate_Q4_Fix_Months * CONST_Intrst234C_Percentage / 100 * (G4 - (H1 + H2 + H3 + H4))))
                   Y5 = WorksheetFunction.Max(0, Math.Round(CONST_Intrst234C_NonCorporate_Q5_Fix_Months * CONST_Intrst234C_Percentage / 100 * (G5 - (H1 + H2 + H3 + H4 + H5))))

                     calcIntrst234C = Y1 + Y2 + Y3 + Y4 + Y5
End If
       
        Calculate_InterestPayable234C = calcIntrst234C
        
End Function

Function Calculate_InterestPayable234CR(intrstPayable234C As Variant) As Double
          calcIntrst234C = 0
          qtrwiseLTCGProvNotExerTotal = 0
          qtrwiseLTCGProvExerTotal = 0
          qtrwiseSTCGIncUs111ATotal = 0
          qtrwiseLotteryIncUs115BBTotal = 0
          partBTILTCGProvNotExerTotal = 0
          partBTILTCGProvExerTotal = 0
          partBTISTCGUs111ATotal = 0
          partBTILotteryIncUs115BBTotal = 0
          schCgOsLTCGProvNotExerTotal = 0
          schCgOsLTCGProvExerTotal = 0
          schCgOsSTCGUs111ATotal = 0
          schCgOsLotteryIncUs115BBTotal = 0
          schSiLTCGProvNotExerTotal = 0
          schSiLTCGProvExerTotal = 0
          schSiSTCGUs111ATotal = 0
          schSiLotteryIncUs115BBTotal = 0

          ' ------------------Declaration of Quaterwise Incomes-----------------------
          If (income <> Null And LTCGIncProvisoNotExercised <> Null And quaterAmounts <> Null) Then
              ' LTCG Proviso Not Exercised Quater wise Total
              qtrwiseLTCGProvNotExerTotal = WorksheetFunction.Max(0, Math.Round(amountQuater1)) + WorksheetFunction.Max(0, Math.Round(amountQuater2)) + WorksheetFunction.Max(0, Math.Round(amountQuater3)) + WorksheetFunction.Max(0, Math.Round(amountQuater4)) + WorksheetFunction.Max(0, Math.Round(amountQuater5))
          End If

          If (income <> Null And LTCGIncProvisoExercised <> Null And quaterAmounts <> Null) Then
              ' LTCG Proviso Exercised Quater wise Total
              qtrwiseLTCGProvExerTotal = WorksheetFunction.Max(0, Math.Round(amountQuater1)) + WorksheetFunction.Max(0, Math.Round(amountQuater2)) + WorksheetFunction.Max(0, Math.Round(amountQuater3)) + WorksheetFunction.Max(0, Math.Round(amountQuater4)) + WorksheetFunction.Max(0, Math.Round(amountQuater5))
          End If

          If (income <> Null And STCGUndSec111A <> Null And quaterAmounts <> Null) Then
            ' STCG Under Section 111A Quater wise Total
              qtrwiseSTCGIncUs111ATotal = WorksheetFunction.Max(0, Math.Round(amountQuater1)) + WorksheetFunction.Max(0, Math.Round(amountQuater2)) + WorksheetFunction.Max(0, Math.Round(amountQuater3)) + WorksheetFunction.Max(0, Math.Round(amountQuater4)) + WorksheetFunction.Max(0, Math.Round(amountQuater5))
          End If

         If (income <> Null And lotteryIncUs115BB <> Null And quaterAmounts <> Null) Then
            ' Lottery Income Under Section 115BB Quater wise Total
              qtrwiseLotteryIncUs115BBTotal = WorksheetFunction.Max(0, Math.Round(amountQuater1)) + WorksheetFunction.Max(0, Math.Round(amountQuater2)) + WorksheetFunction.Max(0, Math.Round(amountQuater3)) + WorksheetFunction.Max(0, Math.Round(amountQuater4)) + WorksheetFunction.Max(0, Math.Round(amountQuater5))
         End If

         ' --------------------Declaration of Part B-TI Values---------------------
         If (income <> Null And LTCGIncProvisoNotExercised <> Null And partBTIandSchCGOSSIAmount <> Null) Then
           ' LTCG Proviso Not Exercised Part B-TI Value
            partBTILTCGProvNotExerTotal = Math.Round(partBTIAmount)
         End If

         If (income <> Null And LTCGIncProvisoExercised <> Null And partBTIandSchCGOSSIAmount <> Null) Then
            ' LTCG Proviso Exercised Part B-TI Value
            partBTILTCGProvExerTotal = Math.Round(partBTIAmount)
         End If

         If (income <> Null And STCGUndSec111A <> Null And partBTIandSchCGOSSIAmount <> Null) Then
            ' STCG Income under Section 111A Part B-TI Value
            partBTISTCGUs111ATotal = Math.Round(partBTIAmount)
         End If

         If (income <> Null And lotteryIncUs115BB <> Null And partBTIandSchCGOSSIAmount <> Null) Then
            ' Lottery Income under Section 115BB Part B-TI Value
            partBTILotteryIncUs115BBTotal = Math.Round(partBTIAmount)
         End If

         ' ------------------- Declaration of Schedule CG-OS Values-------------
         If (income <> Null And LTCGIncProvisoNotExercised <> Null And partBTIandSchCGOSSIAmount <> Null) Then
            ' LTCG Proviso Not Exercised Schedule CG-OS Value
            schCgOsLTCGProvNotExerTotal = Math.Round(scheduleCGOSAmount)
         End If

         If (income <> Null And LTCGIncProvisoExercised <> Null And partBTIandSchCGOSSIAmount <> Null) Then
            ' LTCG Proviso Exercised Schedule CG-OS Value
            schCgOsLTCGProvExerTotal = Math.Round(scheduleCGOSAmount)
         End If

         If (income <> Null And STCGUndSec111A <> Null And partBTIandSchCGOSSIAmount <> Null) Then
            ' STCG Income under Section 111A Schedule CG-OS Value
            schCgOsSTCGUs111ATotal = Math.Round(scheduleCGOSAmount)
         End If

         If (income <> Null And lotteryIncUs115BB <> Null And partBTIandSchCGOSSIAmount <> Null) Then
            ' Lottery Income under Section 115BB Schedule CG-OS Value
            schCgOsLotteryIncUs115BBTotal = Math.Round(scheduleCGOSAmount)
         End If

         ' ------------------- Declaration of Schedule SI Values-------------
         If (income <> Null And LTCGIncProvisoNotExercised <> Null And partBTIandSchCGOSSIAmount <> Null) Then
            ' LTCG Proviso Not Exercised Schedule SI Value
            schSiLTCGProvNotExerTotal = Math.Round(scheduleSIAmount)
         End If

         If (income <> Null And LTCGIncProvisoExercised <> Null And partBTIandSchCGOSSIAmount <> Null) Then
            ' LTCG Proviso Exercised Schedule SI Value
            schSiLTCGProvExerTotal = Math.Round(scheduleSIAmount)
         End If

         If (income <> Null And STCGUndSec111A <> Null And partBTIandSchCGOSSIAmount <> Null) Then
            ' STCG Income under Section 111A Schedule SI Value
            schSiSTCGUs111ATotal = Math.Round(scheduleSIAmount)
         End If

         If (income <> Null And lotteryIncUs115BB <> Null And partBTIandSchCGOSSIAmount <> Null) Then
            ' Lottery Income under Section 115BB Schedule SI Value
            schSiLotteryIncUs115BBTotal = Math.Round(scheduleSIAmount)
         End If
       
         ' ---------- Declaration of variable used in the Interest 234C Computation ----------
         ' LTCG Proviso Not Exercised variable
         calcLTCGProvisoNotExer = 0
         calcLTCGProvisoNotExer = WorksheetFunction.Max(qtrwiseLTCGProvNotExerTotal, partBTILTCGProvNotExerTotal)
         calcLTCGProvisoNotExer = WorksheetFunction.Max(calcLTCGProvisoNotExer, schCgOsLTCGProvNotExerTotal)
         calcLTCGProvisoNotExer = WorksheetFunction.Max(calcLTCGProvisoNotExer, schSiLTCGProvNotExerTotal)
         ' LTCG Proviso Exercised variable
          calcLTCGProvisoExer = 0
          calcLTCGProvisoExer = WorksheetFunction.Max(qtrwiseLTCGProvExerTotal, partBTILTCGProvExerTotal)
          calcLTCGProvisoExer = WorksheetFunction.Max(calcLTCGProvisoExer, schCgOsLTCGProvExerTotal)
          calcLTCGProvisoExer = WorksheetFunction.Max(calcLTCGProvisoExer, schSiLTCGProvExerTotal)
         ' STCG Income Under Section 111A Variable
          calcSTCGIncUs111A = 0
          calcSTCGIncUs111A = WorksheetFunction.Max(qtrwiseSTCGIncUs111ATotal, partBTISTCGUs111ATotal)
          calcSTCGIncUs111A = WorksheetFunction.Max(calcSTCGIncUs111A, schCgOsSTCGUs111ATotal)
          calcSTCGIncUs111A = WorksheetFunction.Max(calcSTCGIncUs111A, schSiSTCGUs111ATotal)
         ' Lottery Income Under Section 115BB Variable
          calcLotteryIncUs115BB = 0
          calcLotteryIncUs115BB = WorksheetFunction.Max(qtrwiseLotteryIncUs115BBTotal, partBTILotteryIncUs115BBTotal)
          calcLotteryIncUs115BB = WorksheetFunction.Max(calcLotteryIncUs115BB, schCgOsLotteryIncUs115BBTotal)
          calcLotteryIncUs115BB = WorksheetFunction.Max(calcLotteryIncUs115BB, schSiLotteryIncUs115BBTotal)
         ' ------ If quater wise breakup is not given then do the following -----------
         ' Corporate First Quater Variable Declarations
          LTCGProvNotExerQ1 = 0
          LTCGProvExerQ1 = 0
          STCGIncUs111AQ1 = 0
          LotteryIncUs115BBQ1 = 0
          LTCGProvNotExerQ2 = 0
          LTCGProvExerQ2 = 0
          STCGIncUs111AQ2 = 0
          LotteryIncUs115BBQ2 = 0
          LTCGProvNotExerQ3 = 0
          LTCGProvExerQ3 = 0
          STCGIncUs111AQ3 = 0
          LotteryIncUs115BBQ3 = 0
          LTCGProvNotExerQ4 = 0
          LTCGProvExerQ4 = 0
          STCGIncUs111AQ4 = 0
          LotteryIncUs115BBQ4 = 0
          LTCGProvNotExerQ5 = 0
          LTCGProvExerQ5 = 0
          STCGIncUs111AQ5 = 0
          LotteryIncUs115BBQ5 = 0

          If (income <> Null And LTCGIncProvisoNotExercised <> Null And quaterAmounts <> Null) Then
              ' Non-Corporate Second Quater Variable Declarations
              LTCGProvNotExerQ2 = Math.Round(amountQuater2)
         End If

         If (income <> Null And LTCGIncProvisoExercised <> Null And quaterAmounts <> Null) Then
              LTCGProvExerQ2 = Math.Round(amountQuater2)
         End If

         If (income <> Null And STCGUndSec111A <> Null And quaterAmounts <> Null) Then
              STCGIncUs111AQ2 = Math.Round(amountQuater2)
         End If

         If (income <> Null And lotteryIncUs115BB <> Null And quaterAmounts <> Null) Then
              LotteryIncUs115BBQ2 = Math.Round(amountQuater2)
         End If

         If (income <> Null And LTCGIncProvisoNotExercised <> Null And quaterAmounts <> Null) Then
              LTCGProvNotExerQ3 = Math.Round(WorksheetFunction.Max(0, amountQuater3))
              LTCGProvNotExerQ4 = Math.Round(WorksheetFunction.Max(0, amountQuater4))
              LTCGProvNotExerQ5 = Math.Round(WorksheetFunction.Max(0, amountQuater5))
         End If

         If (income <> Null And LTCGIncProvisoExercised <> Null And quaterAmounts <> Null) Then
             LTCGProvExerQ3 = Math.Round(WorksheetFunction.Max(0, amountQuater3))
             LTCGProvExerQ4 = Math.Round(WorksheetFunction.Max(0, amountQuater4))
             LTCGProvExerQ5 = Math.Round(WorksheetFunction.Max(0, amountQuater5))
         End If

         If (income <> Null And STCGUndSec111A <> Null And quaterAmounts <> Null) Then
             STCGIncUs111AQ3 = Math.Round(WorksheetFunction.Max(0, amountQuater3))
             STCGIncUs111AQ4 = Math.Round(WorksheetFunction.Max(0, amountQuater4))
             STCGIncUs111AQ5 = Math.Round(WorksheetFunction.Max(0, amountQuater5))
         End If

         If (income <> Null And lotteryIncUs115BB <> Null And quaterAmounts <> Null) Then
            LotteryIncUs115BBQ3 = Math.Round(WorksheetFunction.Max(0, amountQuater3))
            LotteryIncUs115BBQ4 = Math.Round(WorksheetFunction.Max(0, amountQuater4))
            LotteryIncUs115BBQ5 = Math.Round(WorksheetFunction.Max(0, amountQuater5))
         End If

         If (qtrwiseLTCGProvNotExerTotal = 0) Then
                LTCGProvNotExerQ1 = calcLTCGProvisoNotExer
         End If

         If (qtrwiseLTCGProvExerTotal = 0) Then
                LTCGProvExerQ1 = calcLTCGProvisoExer
         End If

         If (qtrwiseSTCGIncUs111ATotal = 0) Then
                STCGIncUs111AQ1 = calcSTCGIncUs111A
         End If

        If (qtrwiseLotteryIncUs115BBTotal = 0) Then
               LotteryIncUs115BBQ1 = calcLotteryIncUs115BB
        End If

        ' Checking for Corporate or Non-Corporate
        If (intrstPayable234C.personalInfo.formName = 6) Then
        ' ---- If quater wise breakup does not tally with overall total and is less than overall total then the residual value will be added to First Quater Total --------
           If (qtrwiseLTCGProvNotExerTotal < calcLTCGProvisoNotExer) Then
               LTCGProvNotExerQ1 = LTCGProvNotExerQ1 + calcLTCGProvisoNotExer - qtrwiseLTCGProvNotExerTotal
           End If

           If (qtrwiseLTCGProvExerTotal < calcLTCGProvisoExer) Then
                LTCGProvExerQ1 = LTCGProvExerQ1 + calcLTCGProvisoExer - qtrwiseLTCGProvExerTotal
           End If

           If (qtrwiseSTCGIncUs111ATotal < calcSTCGIncUs111A) Then
                STCGIncUs111AQ1 = STCGIncUs111AQ1 + calcSTCGIncUs111A - qtrwiseSTCGIncUs111ATotal
           End If

           If (qtrwiseLotteryIncUs115BBTotal < calcLotteryIncUs115BB) Then
                LotteryIncUs115BBQ1 = LotteryIncUs115BBQ1 + calcLotteryIncUs115BB - qtrwiseLotteryIncUs115BBTotal
           End If

           ' -------------- If the quater wise breakup does not tally with overall total and is greater than overall total then the quater wise  breakup as reported is taken for computation ----------------------

           If (qtrwiseLTCGProvNotExerTotal >= calcLTCGProvisoNotExer And income <> Null And LTCGIncProvisoNotExercised <> Null And quaterAmounts <> Null) Then
                LTCGProvNotExerQ1 = Math.Round(amountQuater1)
           End If

           If (qtrwiseLTCGProvExerTotal >= calcLTCGProvisoExer And income <> Null And LTCGIncProvisoExercised <> Null And LTCGIncProvisoExercised.quaterAmounts <> Null) Then
                LTCGProvExerQ1 = Math.Round(amountQuater1)
           End If

           If (qtrwiseSTCGIncUs111ATotal >= calcSTCGIncUs111A And income <> Null And STCGUndSec111A <> Null And STCGUndSec111A.quaterAmounts <> Null) Then
                STCGIncUs111AQ1 = Math.Round(amountQuater1)
           End If

           If (qtrwiseLotteryIncUs115BBTotal >= calcLotteryIncUs115BB And income <> Null And lotteryIncUs115BB <> Null And lotteryIncUs115BB.quaterAmounts <> Null) Then
                LotteryIncUs115BBQ1 = Math.Round(amountQuater1)
           End If
     Else
        ' ---- If quater wise breakup does not tally with overall total and is less than overall total then the residual value will be added to First Quater Total --------
           If (qtrwiseLTCGProvNotExerTotal < calcLTCGProvisoNotExer) Then
                LTCGProvNotExerQ2 = LTCGProvNotExerQ2 + calcLTCGProvisoNotExer - qtrwiseLTCGProvNotExerTotal
           End If

           If (qtrwiseLTCGProvExerTotal < calcLTCGProvisoExer) Then
                LTCGProvExerQ2 = LTCGProvExerQ2 + calcLTCGProvisoExer - qtrwiseLTCGProvExerTotal
           End If

           If (qtrwiseSTCGIncUs111ATotal < calcSTCGIncUs111A) Then
                STCGIncUs111AQ2 = STCGIncUs111AQ2 + calcSTCGIncUs111A - qtrwiseSTCGIncUs111ATotal
           End If

           If (qtrwiseLotteryIncUs115BBTotal < calcLotteryIncUs115BB) Then
                LotteryIncUs115BBQ2 = LotteryIncUs115BBQ2 + calcLotteryIncUs115BB - qtrwiseLotteryIncUs115BBTotal
           End If
      End If

         ' Calculation Gross Assessed Tax (GAT).
         grossAssessedTax = WorksheetFunction.Max(0, ITRGrossTaxLiability - (TDSAmtUsed + TCSPaid + section89 + section90 + section91 + creditUs115JAA + Math.Round(calcLTCGProvisoNotExer * CONST_LTCGNoProviso_TaxRate / 100) + Math.Round(calcLTCGProvisoExer * CONST_LTCGProviso_TaxRate / 100) + Math.Round(calcSTCGIncUs111A * CONST_STCG_TaxRate / 100) + Math.Round(calcLotteryIncUs115BB * CONST_WinningFromLotteryIncome_TaxRate / 100)))
         ' Rounding down to the nearest hundred rupees.
         grossAssessedTax = grossAssessedTax / 100
         decVal = CInt(grossAssessedTax)
         grossAssessedTax = decVal * 100

         ' Checking GAT limit for calculating interest payable 234C.
         If (grossAssessedTax > CONST_GAT_Limit) Then
              ' Assigining Advance tAx paid in quater 1 to 5 to the variables.
              H1 = advanceTaxPaidQ1
              H2 = advanceTaxPaidQ2
              H3 = advanceTaxPaidQ3
              H4 = advanceTaxPaidQ4
              H5 = advanceTaxPaidQ5
              ' CONST_Intrst234C_Others_Q1_TaxRate *
              G1 = grossAssessedTax
              G2 = G1 + Math.Round(LTCGProvNotExerQ1 * CONST_LTCGNoProviso_TaxRate / 100) + Math.Round(LTCGProvExerQ1 * CONST_LTCGProviso_TaxRate / 100) + Math.Round(STCGIncUs111AQ1 * CONST_STCG_TaxRate / 100) + Math.Round(LotteryIncUs115BBQ1 * CONST_WinningFromLotteryIncome_TaxRate / 100)
              G3 = G2 + Math.Round(LTCGProvNotExerQ2 * CONST_LTCGNoProviso_TaxRate / 100) + Math.Round(LTCGProvExerQ2 * CONST_LTCGProviso_TaxRate / 100) + Math.Round(STCGIncUs111AQ2 * CONST_STCG_TaxRate / 100) + Math.Round(LotteryIncUs115BBQ2 * CONST_WinningFromLotteryIncome_TaxRate / 100)
              G4 = G3 + Math.Round(LTCGProvNotExerQ3 * CONST_LTCGNoProviso_TaxRate / 100) + Math.Round(LTCGProvExerQ3 * CONST_LTCGProviso_TaxRate / 100) + Math.Round(STCGIncUs111AQ3 * CONST_STCG_TaxRate / 100) + Math.Round(LotteryIncUs115BBQ3 * CONST_WinningFromLotteryIncome_TaxRate / 100)
              G5 = grossAssessedTax + Math.Round(calcLTCGProvisoNotExer * CONST_LTCGNoProviso_TaxRate / 100) + Math.Round(calcLTCGProvisoExer * CONST_LTCGProviso_TaxRate / 100) + Math.Round(calcSTCGIncUs111A * CONST_STCG_TaxRate / 100) + Math.Round(calcLotteryIncUs115BB * CONST_WinningFromLotteryIncome_TaxRate / 100)
              G1Perc = 1
              G2Perc = 1
              G3Perc = 1
              G4Perc = 1
              G5Perc = 1

            If (intrstPayable234C.personalInfo.formName = 6) Then
                  G1Perc = CONST_Corporate_TaxRate1 / 100
                  G2Perc = CONST_Corporate_TaxRate2 / 100
                  G3Perc = CONST_Corporate_TaxRate3 / 100
           Else
                  G1Perc = 0
                  G2Perc = CONST_NonCorporate_TaxRate1 / 100
                  G3Perc = CONST_NonCorporate_TaxRate2 / 100
           End If
                  G1 = G1Perc * G1
                  G2 = G2Perc * G2
                  G3 = G3Perc * G3
                  G4 = G4Perc * G4
                  G5 = G5Perc * G5
                  Y1 = 0
                  Y2 = 0
                  Y3 = 0
                  Y4 = 0
                  Y5 = 0

         If (ThisComponent.Sheets(1-1).getCellRangeByName("formname").value = 6) Then
                If (H1 <= CONST_Corporate_Percentage1 * G1) Then
                       Y1 = WorksheetFunction.Max(0, Math.Round(CONST_Intrst234C_Fix_Months * CONST_Intrst234C_Percentage / 100 * (G1 - H1)))
                End If
                If (H2 <= CONST_Corporate_Percentage2 * G2) Then
                       Y2 = WorksheetFunction.Max(0, Math.Round(CONST_Intrst234C_Fix_Months * CONST_Intrst234C_Percentage / 100 * (G2 - (H1 + H2))))
                End If
                       Y3 = WorksheetFunction.Max(0, Math.Round(CONST_Intrst234C_Fix_Months * CONST_Intrst234C_Percentage / 100 * (G3 - (H1 + H2 + H3))))
                       Y4 = WorksheetFunction.Max(0, Math.Round(CONST_Intrst234C_Corporate_Q4_Fix_Months * CONST_Intrst234C_Percentage / 100 * (G4 - (H1 + H2 + H3 + H4))))
                       Y5 = WorksheetFunction.Max(0, Math.Round(CONST_Intrst234C_Corporate_Q5_Fix_Months * CONST_Intrst234C_Percentage / 100 * (G5 - (H1 + H2 + H3 + H4 + H5))))
                       calcIntrst234C = Y1 + Y2 + Y3 + Y4 + Y5
          Else
                Y1 = WorksheetFunction.Max(0, Math.Round(CONST_Intrst234C_Fix_Months * CONST_Intrst234C_Percentage / 100 * (G1 - H1)))
                Y2 = WorksheetFunction.Max(0, Math.Round(CONST_Intrst234C_Fix_Months * CONST_Intrst234C_Percentage / 100 * (G2 - (H1 + H2))))
                Y3 = WorksheetFunction.Max(0, Math.Round(CONST_Intrst234C_Fix_Months * CONST_Intrst234C_Percentage / 100 * (G3 - (H1 + H2 + H3))))
                Y4 = WorksheetFunction.Max(0, Math.Round(CONST_Intrst234C_NonCorporate_Q4_Fix_Months * CONST_Intrst234C_Percentage / 100 * (G4 - (H1 + H2 + H3 + H4))))
                Y5 = WorksheetFunction.Max(0, Math.Round(CONST_Intrst234C_NonCorporate_Q5_Fix_Months * CONST_Intrst234C_Percentage / 100 * (G5 - (H1 + H2 + H3 + H4 + H5))))
                calcIntrst234C = Y1 + Y2 + Y3 + Y4 + Y5

       End If
   End If

   ' Setting Income values to output facts
   ' Setting LTCG Proviso Not Exercised Output fact
   amountQuater1 = LTCGProvNotExerQ1
   amountQuater2 = LTCGProvNotExerQ2
   amountQuater3 = LTCGProvNotExerQ3
   amountQuater4 = LTCGProvNotExerQ4
   amountQuater5 = LTCGProvNotExerQ5
   amountTotal = calcLTCGProvisoNotExer

   ' Setting LTCG Proviso Exercised Output fact
   amountQuater1 = LTCGProvExerQ1
   amountQuater2 = LTCGProvExerQ2
   amountQuater3 = LTCGProvExerQ3
   amountQuater4 = LTCGProvExerQ4
   amountQuater5 = LTCGProvExerQ5
   amountTotal = calcLTCGProvisoExer

   ' Setting STCG Income Under Section 111A Output factamountQuater1 = STCGIncUs111AQ1
   amountQuater2 = STCGIncUs111AQ2
   amountQuater3 = STCGIncUs111AQ3
   amountQuater4 = STCGIncUs111AQ4
   amountQuater5 = STCGIncUs111AQ5
   amountTotal = calcSTCGIncUs111A

   ' Setting Lottery Income Under Section 115BB Output fact
   amountQuater1 = LotteryIncUs115BBQ1
   amountQuater2 = LotteryIncUs115BBQ2
   amountQuater3 = LotteryIncUs115BBQ3
   amountQuater4 = LotteryIncUs115BBQ4
   amountQuater5 = LotteryIncUs115BBQ5
   amountTotal = calcLotteryIncUs115BB

       Calculate_InterestPayable234CR = calcIntrst234C
End Function

Function Calculate_RefundDue(calcTaxInterestLiability As Double, calcTotTaxesPaid As Double) As Double
       calcRefundDue = 0
       If (calcTotTaxesPaid > calcTaxInterestLiability) Then
          calcRefundDue = WorksheetFunction.Max(0, calcTotTaxesPaid - calcTaxInterestLiability)
       End If
       IntrstRef_Counter = IntrstRef_Counter + 1

       Calculate_RefundDue = calcRefundDue
End Function

Function Calculate_TotalInterestPay(intrstPay234A As Double, intrstPay234B As Double, intrstPay234C As Double) As Double
       ' This rule calculates the total interest payable.
       Calculate_TotalInterestPay = (intrstPay234A + intrstPay234B + intrstPay234C)
End Function

Function Calculate_TotalSATPaid(intrstSupplements As Variant) As Double
            totSATPaid = 0
            satPaymentList = selfAssessmentTaxPaid
          If (satPaymentList.Size() > 0) Then
                java.util.listIterator satPaymentIte = satPaymentList.listIterator()
                    Do While (satPaymentIte.hasNext())
                         satPayment = satPaymentIte.Next()
                         totSATPaid = totSATPaid + Math.Round(satPayment.amount)
                    Loop
          End If
          Calculate_TotalSATPaid = totSATPaid
End Function

Function Calculate_TotalTaxPlusInterestPay(totalInterestPay As Double, grossTaxLiable As Double, section89 As Double, section90 As Double, section91 As Double) As Double
      Calculate_TotalTaxPlusInterestPay = (totalInterestPay + WorksheetFunction.Max(0, grossTaxLiable - (section89 + section90 + section91)))
End Function

Function Decide_TDSAmtToBeUse(totTDSClaimAmt As Double, totValidTANTDSClaimAmt As Double, totValidTANTDSClaimAmtForYear As Double, totMatchedAmt As Double) As Double
       TDSAmtToBeUsed = 0
   If (totTDSClaimAmt > CONST_Total_TDS_Claim_Amount And IntrstRef_Counter = 1) Then
     TDSAmtToBeUsed = WorksheetFunction.Min(totMatchedAmt, totValidTANTDSClaimAmtForYear)
     ReCalc_IntrstRefund_Flag = False
   ElseIf (totTDSClaimAmt <= CONST_Total_TDS_Claim_Amount And IntrstRef_Counter = 1) Then
     TDSAmtToBeUsed = WorksheetFunction.Min(totValidTANTDSClaimAmt, totValidTANTDSClaimAmtForYear)
   End If

   If (TDSAmtToBeUsed = totValidTANTDSClaimAmt And IntrstRef_Counter = 1) Then
      Total_Valid_TAN_TDS_Claim_Flag = "Y"
   ElseIf (TDSAmtToBeUsed = totValidTANTDSClaimAmtForYear And IntrstRef_Counter = 1) Then
      Total_Valid_TAN_TDS_Claim_For_Year_Flag = "Y"
   ElseIf (TDSAmtToBeUsed = totMatchedAmt And IntrstRef_Counter = 1) Then
      Total_Matched_Amount_Flag = "Y"
   End If

   If (IntrstRef_Counter = 2) Then
     TDSAmtToBeUsed = totMatchedAmt
     Total_Matched_Amount_Flag = "Y"
     Total_Valid_TAN_TDS_Claim_For_Year_Flag = "N"
      Total_Valid_TAN_TDS_Claim_Flag = "N"
   End If
   Decide_TDSAmtToBeUse = TDSAmtToBeUsed
End Function

Function Recalculate_IntrstRefund(intrstType As Variant, calcRefund As Double)
      If (ReCalc_IntrstRefund_Flag And IntrstRef_Counter <= 2 And calcRefund > CONST_Refund_Recalc_Limit) Then
         MsgBox (intrstType)
      End If
End Function

Function RoundDownToNearestHundredRupees(value As Double) As Double
      ' This function round down to the interest to the nearest hundered rupees.
        value = value / 100
        decVal = CInt(value)
        roundDownVal = decVal * 100
         RoundDownToNearestHundredRupees = roundDownVal
End Function

Sub COMPUTE_INTEREST()

Dim SYSCALCGrossTaxLiability As Double
Dim matchedAdvanceTax As Double
Dim TDSAmtUsed As Double
Dim section89 As Double
Dim section90 As Double
Dim section91 As Double

Call Module3.filingdate



SYSCALCGrossTaxLiability = ThisComponent.Sheets(5-1).getCellRangeByName("GTLR").value
' INSTEAD OF THIS IF U USE GTLS IT WILL BE SYS. NOW IT IS AS PER RETURN.

matchedAdvanceTax = ThisComponent.Sheets(2-1).getCellRangeByName("IncD.AdvanceTax").value
TDSAmtUsed = ThisComponent.Sheets(2-1).getCellRangeByName("IncD.TDS").value
section89 = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.Section89").value
section90 = ThisComponent.Sheets(1-1).getCellRangeByName("IncD.Section90and91").value
section91 = 0

       intrstPayUs234A = Calculate_InterestPayable234A(SYSCALCGrossTaxLiability, matchedAdvanceTax, TDSAmtUsed, section89, section90, section91)
       ThisComponent.Sheets(5-1).getCellRangeByName("Calc_234A").value = Round(intrstPayUs234A)
       
       
'       intrstPayUs234AR = Calculate_InterestPayable234AR(InterestComputationType)
               
        Dim gatt As Long
        gatt = ThisComponent.Sheets(5-1).getCellRangeByName("GAT").value
        'intrstPayUs234C = Calculate_InterestPayable234C(SYSCALCGrossTaxLiability, matchedAdvanceTax, TDSAmtUsed, section89, section90, section91)
        intrstPayUs234C = ThisComponent.Sheets(5-1).getCellRangeByName("int234C").value
        gatt = ThisComponent.Sheets(5-1).getCellRangeByName("GATS").value
        
        If (gatt < 10000) Then
        intrstPayUs234C = 0
        End If
        
        
        ThisComponent.Sheets(5-1).getCellRangeByName("Calc_234C").value = Round(intrstPayUs234C)
        
        'intrstPayUs234C = 6431
        intrstPayUs234B = Calculate_InterestPayable234B(intrstPayUs234A, intrstPayUs234C)
        ThisComponent.Sheets(5-1).getCellRangeByName("Calc_234B").value = Round(intrstPayUs234B)
'       intrstPayUs234CR = Calculate_InterestPayable234CR(InterestComputationType)
'       totalInrstPay = Calculate_TotalInterestPay(intrstPayUs234A, intrstPayUs234B, intrstPayUs234CR)
'       aggTaxInterestLiability = Calculate_TotalTaxPlusInterestPay(totalInrstPay, SYSCALCGrossTaxLiability, section89, section90, section91)
'       TDSAmtUsed = Decide_TDSAmtToBeUse(totalTDSClaimAmount, totalValidTANTDSClaim, totalValidTANTDSClaimForYear, totalMatchedAmount)
'       totTaxesPaid = matchedAdvanceTax + TDSAmtUsed + TCSPaid + Calculate_TotalSATPaid(interestSupplements)
'       balTaxPayable = Calculate_BalanceTaxPayable(aggTaxInterestLiability, totTaxesPaid)
'       refundDue = Calculate_RefundDue(aggTaxInterestLiability, totTaxesPaid)
'       itrID = itrID
'       totalMatchedAmountFlag = Total_Matched_Amount_Flag
'       totalValidTANTDSClaimFlag = Total_Valid_TAN_TDS_Claim_Flag
'       totalValidTANTDSClaimForYearFlag = Total_Valid_TAN_TDS_Claim_For_Year_Flag
'       CALCIntrstPayUs234A = intrstPayUs234A
'       CALCIntrstPayUs234B = intrstPayUs234B
'       CALCIntrstPayUs234C = intrstPayUs234CR
'       CALCTotalIntrstPay = totalInrstPay
'       CALCAggregateTaxInterestLiability = aggTaxInterestLiability
'       CALCTotalTaxesPaid = totTaxesPaid
'       CALCBalTaxPayable = balTaxPayable
'       calcRefundDue = refundDue
'       Call Recalculate_IntrstRefund(InterestComputationType, refundDue)
End Sub

Function monthdiff2(startDate As String, endDate As String) As Integer
noofdaysinmonth = 31
monthdiff2 = 0
startDate = CStr(Dformat1(startDate, "yyyy-mm-dd"))
endDate = CStr(Dformat1(endDate, "yyyy-mm-dd"))
startyear = Mid(startDate, 1, 4)
endyear = Mid(endDate, 1, 4)
startmonth = Mid(startDate, 6, 2)
If (startmonth = 2) Then
 noofdaysinmonth = 28
 
End If
startday = Mid(startDate, 9, 2)
If ((startmonth = 4) Or (startmonth = 6) Or (startmonth = 9) Or (startmonth = 11)) Then
 noofdaysinmonth = 30
 
End If


endmonth = Mid(endDate, 6, 2)
monthdiff2 = (CInt(endyear) - CInt(startyear)) * 12
monthdiff2 = monthdiff2 + (CInt(endmonth) - CInt(startmonth))

If (startDate < endDate) Then
        
If (CInt(startday) < Mid(endDate, 9, 2)) Then
        monthdiff2 = monthdiff2 + 1
End If
End If
End Function


