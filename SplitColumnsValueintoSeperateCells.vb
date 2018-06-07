Attribute VB_Name = "Module2"


Sub splitByColB()
    
    Dim sheetrangebase As String
    sheetrangebase = "B"
    
    Dim sheetrangebaseC As String
    sheetrangebaseC = "C"
    
    Dim sheetrangebaseD As String
    sheetrangebaseD = "D"
    
    Dim sheetrangebasetitle As String
    sheetrangebasetitle = "A"
    
    Dim sheetrange As String
    sheetrange = ""
    
    Dim cMax, CValue As Integer
    cMax = 800
    
    Dim Sp1, Sp2 As String
    Dim Sp3, Sp4, Sp5 As String
    
    Dim cSplits As Double
    cSplits = 0
    
    Dim rc As Integer
    rc = 1
    
    Dim PRn As Integer
    
    Dim CurrentPRr As Integer
    CurrentPRr = 0
    
    Dim sht As Worksheet
    Dim CellOrder As Range
    Dim CellType, CellPR As Range
    'Dim CellOrder, CellType As String
    Set sht = Sheets("Sheet4")
    Range("F1").Select
    Dim r As Range, i As Long, ar
    'Set r = Worksheets("Sheet1").Range("B1").End(xlDown)
    
    NumRows = Range("F1", Range("F1").End(xlDown)).Rows.Count
    Set r = Worksheets("Sheet1").Range("F1").End(xlDown)
    
    Range("F1").Select
    
    NCB = "B"
    NCC = "C"
    NCD = "D"
    NCA = "A"
                            
    sTypeofCol = "Detailed Description"
                            
      For x = 1 To NumRows
         ' Insert your code here.
         ' Selects cell down 1 row from active cell.
         'GET PR NUMBER ==================================
         ActiveCell.Offset(0, -5).Select
         'ActiveCell.Offset(1, 0).Select
         PRn = ActiveCell.value
         
         ActiveCell.Offset(0, 5).Select
         CellValueOrig = ActiveCell.value
         
'         Set CellOrder = sht.Range(sheetrangebaseC)
'         Set CellType = sht.Range(sheetrangebaseD)
         'Set CellPR = sht.Range(ActiveCell)
         'CellOrder.value = 1
         'CellPR.value = PRn
                
         cValueOrig = ActiveCell.value
         cValueOrigLength = Len(ActiveCell.value)
         cSplits = cValueOrigLength / cMax
         
                If (cValueOrigLength > cMax) Then
'                cSplits = cValueOrig / cMax
'                Set CellOrder = sht.Range(sheetrangebaseC)
'                Set CellType = sht.Range(sheetrangebaseD)
'                Set CellPR = sht.Range(sheetrangetitle)
'                CellOrder.value = 1
'                CellPR.value = PRn
                
                    If (cSplits > 14) Then
                    
                        Sp1 = Left(cValueOrig, cMax)
                        Sp2 = Mid(cValueOrig, cValueOrigLength - cMax, cMax)
                        Sp3 = Mid(cValueOrig, cValueOrigLength - (cMax * 2), cMax)
                        Sp4 = Mid(cValueOrig, cValueOrigLength - (cMax * 3), cMax)
                        Sp5 = Mid(cValueOrig, cValueOrigLength - (cMax * 4), cMax)
                        Sp6 = Mid(cValueOrig, cValueOrigLength - (cMax * 5), cMax)
                        Sp7 = Mid(cValueOrig, cValueOrigLength - (cMax * 6), cMax)
                        Sp8 = Mid(cValueOrig, cValueOrigLength - (cMax * 7), cMax)
                        Sp9 = Mid(cValueOrig, cValueOrigLength - (cMax * 8), cMax)
                        Sp10 = Mid(cValueOrig, cValueOrigLength - (cMax * 9), cMax)
                        Sp11 = Mid(cValueOrig, cValueOrigLength - (cMax * 10), cMax)
                        Sp12 = Mid(cValueOrig, cValueOrigLength - (cMax * 11), cMax)
                        Sp13 = Mid(cValueOrig, cValueOrigLength - (cMax * 12), cMax)
                        Sp14 = Mid(cValueOrig, cValueOrigLength - (cMax * 13), cMax)
                        Sp15 = Right(cValueOrig, cValueOrigLength - (cMax * 14))
                        
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "1"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp1
                            CellPR.value = PRn
                            
                            ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "2"
                            CellType.value = "Progress Notes"
                            CellValue.value = Sp2
                            CellPR.value = PRn
                            
                            ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "3"
                            CellType.value = "Progress Notes"
                            CellValue.value = Sp3
                            CellPR.value = PRn
                            
                            ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "4"
                            CellType.value = "Progress Notes"
                            CellValue.value = Sp4
                            CellPR.value = PRn
                        
                             ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "5"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp5
                            CellPR.value = PRn
                                                        ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "6"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp6
                            CellPR.value = PRn
                                                        ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "7"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp7
                            CellPR.value = PRn
                                                        ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "8"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp8
                            CellPR.value = PRn
                                                        ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "9"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp9
                            CellPR.value = PRn
                                                        ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "10"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp10
                            CellPR.value = PRn
                                                        ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "11"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp11
                            CellPR.value = PRn
                                                        ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "12"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp12
                            CellPR.value = PRn
                                                        ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "13"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp13
                            CellPR.value = PRn
                                                        ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "14"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp14
                            CellPR.value = PRn
                                                        ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "14"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp4
                            CellPR.value = PRn
                        

                    ElseIf (cSplits > 4) Then
                    
                        Sp1 = Left(cValueOrig, cMax)
                        Sp2 = Mid(cValueOrig, cValueOrigLength - cMax, cMax)
                        Sp3 = Mid(cValueOrig, cValueOrigLength - (cMax * 2), cMax)
                        Sp4 = Right(cValueOrig, cValueOrigLength - (cMax * 3))
                        Sp5 = Right(cValueOrig, cValueOrigLength - (cMax * 4))
                        
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "1"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp1
                            CellPR.value = PRn
                            
                            ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "2"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp2
                            CellPR.value = PRn
                            
                            ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "3"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp3
                            CellPR.value = PRn
                            
                            ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "4"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp4
                            CellPR.value = PRn
                            
                            ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "5"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp5
                            CellPR.value = PRn
                    
                    ElseIf (cSplits > 3) Then
                    
                        Sp1 = Left(cValueOrig, cMax)
                        Sp2 = Mid(cValueOrig, cValueOrigLength - cMax, cMax)
                        Sp3 = Mid(cValueOrig, cValueOrigLength - (cMax * 2), cMax)
                        Sp4 = Right(cValueOrig, cValueOrigLength - (cMax * 3))
                        
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "1"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp1
                            CellPR.value = PRn
                            
                            ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "2"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp2
                            CellPR.value = PRn
                            
                            ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "3"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp3
                            CellPR.value = PRn
                            
                            ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "4"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp4
                            CellPR.value = PRn
                            
                    ElseIf (cSplits > 2) Then
                    
                        Sp1 = Left(cValueOrig, cMax)
                        Sp2 = Mid(cValueOrig, cValueOrigLength - cMax, cMax)
                        Sp3 = Right(cValueOrig, cValueOrigLength - (cMax + cMax))
                                                     
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "1"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp1
                            CellPR.value = PRn
                            
                            ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "2"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp2
                            CellPR.value = PRn
                            
                            ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "3"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp3
                            CellPR.value = PRn
                            
                            ' =========================================
                    
                    ElseIf (cSplits > 1) Then
                    
                        Sp1 = Left(cValueOrig, cMax)
                        Sp2 = Right(cValueOrig, cValueOrigLength - cMax)
                    
                        ' SETUP CELLS ===========================
        
''                        Sheets("Sheet2").Select
''                        ActiveSheet.Range(sheetrange).Select
''                        ActiveSheet.Paste
                            
  
                            
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "1"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp1
                            CellPR.value = PRn
                            
                            ' =========================================
                            
                            rc = rc + 1
                            NCAtmp = NCA & rc
                            NCBtmp = NCB & rc
                            NCCtmp = NCC & rc
                            NCDtmp = NCD & rc
                        
                            Set CellValue = sht.Range(NCBtmp)
                            Set CellOrder = sht.Range(NCCtmp)
                            Set CellType = sht.Range(NCDtmp)
                            Set CellPR = sht.Range(NCAtmp)
                            
                            CellOrder.value = "2"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp2
                            CellPR.value = PRn
                        
                    Else
                        
                        ' ANY NUMBER NOW =================================
                        
                        
                    
                    End If
                
'                Set CellOrder = sht.Range(sheetrangebaseC)
'                Set CellType = sht.Range(sheetrangebaseD)
'                'CellOrder.Value = "1"
'                CellType.value = sTypeofCol
                
            Else
                            
                NCAtmp = NCA & rc
                NCBtmp = NCB & rc
                NCCtmp = NCC & rc
                NCDtmp = NCD & rc
                        
                Set CellValue = sht.Range(NCBtmp)
                Set CellOrder = sht.Range(NCCtmp)
                Set CellType = sht.Range(NCDtmp)
                Set CellPR = sht.Range(NCAtmp)
                            
                CellOrder.value = "1"
                CellType.value = sTypeofCol
                CellValue.value = cValueOrig
                CellPR.value = PRn
                
            End If
            
         rc = rc + 1
         ActiveCell.Offset(1, 0).Select
      Next
    
    'Set r = Worksheets("Sheet1").Range("B1412:" & ActiveSheet.Range("B1").End(xlDown).Address)
    'Range("A1").Select
        ' CREATE BASE STRINGS FOR PASTING INTO SHEET 2 --------

        
        ' PASTE PR NUMBER INTO A COLUMN -----------------
        'Range(sheetrangetitle).Select
        'Sheet2.Range(sheetrangetitle).Value = i
    
'    Do While r.Row > 1
'
'
'        '
'        sheetrangebase = "B"
'        sheetrangebaseC = "C"
'        sheetrangebaseD = "D"
'        sheetrangebasetitle = "A"
'        sheetrange = sheetrangebase & rc
'        sheetrangebase = sheetrangebase & rc
'        sheetrangetitle = sheetrangebasetitle & rc
'        sheetrangebaseC = sheetrangebaseC & rc
'        sheetrangebaseD = sheetrangebaseD & rc
'
''        Sheets("Sheet1").Select
''        Range("A1").Select
''        Set CellDValue = sht.Range(sheetrangebase)
''
''        CValuetmp = CellDValue.value
'
'        CValue = Len(r.value)
'        CTextValue = r.value
'
'        Sheets("Sheet1").Select
'        Range(sheetrangetitle).Select
'         Selection.Copy
'
'         Sheets("Sheet2").Select
'        ActiveSheet.Range(sheetrangetitle).Select
'        ActiveSheet.Paste
'
'        If (PRn <> 0) Then
'        Else
'            PRn = sht.Range(sheetrangetitle)
'        End If
'
'        ' PASTE PR NUMBER INTO A COLUMN -----------------
'
'            If (CValue > cMax) Then
'                cSplits = CValue / cMax
'                Set CellOrder = sht.Range(sheetrangebaseC)
'                Set CellType = sht.Range(sheetrangebaseD)
'                Set CellPR = sht.Range(sheetrangetitle)
'                CellOrder.value = 1
'                CellPR.value = PRn
'
'                    If (cSplits > 4) Then
'
'                        Sp1 = Left(r.value, cMax)
'                        Sp2 = Mid(r.value, CValue - cMax, cMax)
'                        Sp3 = Mid(r.value, CValue - (cMax * 2), cMax)
'                        Sp4 = Mid(r.value, CValue - (cMax * 3), cMax)
'                        Sp5 = Right(r.value, CValue - (cMax * 4))
'
'                    ElseIf (cSplits > 3) Then
'
'                        Sp1 = Left(r.value, cMax)
'                        Sp2 = Mid(r.value, CValue - cMax, cMax)
'                        Sp3 = Mid(r.value, CValue - (cMax * 2), cMax)
'                        Sp4 = Right(r.value, CValue - (cMax * 3))
'
'                            sheetrangebase = "B"
'                            sheetrangebaseC = "C"
'                            sheetrangebaseD = "D"
'                            sheetrangebasetitle = "A"
'                            sheetrange = sheetrangebase & rc
'                            sheetrangebase = sheetrangebase & rc
'                            sheetrangetitle = sheetrangebasetitle & rc
'                            sheetrangebaseC = sheetrangebaseC & rc
'                            sheetrangebaseD = sheetrangebaseD & rc
'
'                            Sheets("Sheet1").Select
'                            Range(sheetrangetitle).Select
'                            Selection.Copy
'
'                            Sheets("Sheet2").Select
'                            ActiveSheet.Range(sheetrangetitle).Select
'                            ActiveSheet.Paste
'
'                            Set CellValue = sht.Range(sheetrange)
'                            Set CellOrder = sht.Range(sheetrangebaseC)
'                            Set CellType = sht.Range(sheetrangebaseD)
'                            CellOrder.value = "1"
'                            CellType.value = sTypeofCol
'                            CellValue.value = Sp1
'                            CellPR.value = PRn
'
'                            ' ADD A SECOND LINE ITEM WITH ORDER 2 =====
'                            ' ADD ONE TO THE ROW COUNT ================
'
'                            'iValue = AddColumn(X, X, X, X)
'
'                            Sheets("Sheet1").Select
'                            Range(sheetrangetitle).Select
'                            Selection.Copy
'
'                            rc = rc + 1
'                            sheetrangebase = "B"
'                            sheetrangebaseC = "C"
'                            sheetrangebaseD = "D"
'                            sheetrangebasetitle = "A"
'                            sheetrange = sheetrangebase & rc
'                            sheetrangebase = sheetrangebase & rc
'                            sheetrangetitle = sheetrangebasetitle & rc
'                            sheetrangebaseC = sheetrangebaseC & rc
'                            sheetrangebaseD = sheetrangebaseD & rc
'
'
'                            ' NEED TO COPY ORIGINAL PR FOR LINE 2 BELOW =========
'                            Sheets("Sheet2").Select
'                            ActiveSheet.Range(sheetrangetitle).Select
'                            ActiveSheet.Paste
'
'                            Set CellValue = sht.Range(sheetrange)
'                            Set CellOrder = sht.Range(sheetrangebaseC)
'                            Set CellType = sht.Range(sheetrangebaseD)
'                            CellOrder.value = "2"
'                            CellType.value = sTypeofCol
'                            CellValue.value = Sp2
'                            CellPR.value = PRn
'
'                            ' ================================
'
'
''                            Sheets("Sheet1").Select
''                            Range(sheetrangetitle).Select
''                            Selection.Copy
'
'                            rc = rc + 1
'                            sheetrangebase = "B"
'                            sheetrangebaseC = "C"
'                            sheetrangebaseD = "D"
'                            sheetrangebasetitle = "A"
'                            sheetrange = sheetrangebase & rc
'                            sheetrangebase = sheetrangebase & rc
'                            sheetrangetitle = sheetrangebasetitle & rc
'                            sheetrangebaseC = sheetrangebaseC & rc
'                            sheetrangebaseD = sheetrangebaseD & rc
'
'
'                            ' NEED TO COPY ORIGINAL PR FOR LINE 3 BELOW =========
'                            Sheets("Sheet2").Select
'                            ActiveSheet.Range(sheetrangetitle).Select
'                            ActiveSheet.Paste
'
'                            Set CellValue = sht.Range(sheetrange)
'                            Set CellOrder = sht.Range(sheetrangebaseC)
'                            Set CellType = sht.Range(sheetrangebaseD)
'                            CellOrder.value = "3"
'                            CellType.value = sTypeofCol
'                            CellValue.value = Sp3
'                            CellPR.value = PRn
'
''                            Sheets("Sheet1").Select
''                            Range(sheetrangetitle).Select
''                            Selection.Copy
'
'                            rc = rc + 1
'                            sheetrangebase = "B"
'                            sheetrangebaseC = "C"
'                            sheetrangebaseD = "D"
'                            sheetrangebasetitle = "A"
'                            sheetrange = sheetrangebase & rc
'                            sheetrangebase = sheetrangebase & rc
'                            sheetrangetitle = sheetrangebasetitle & rc
'                            sheetrangebaseC = sheetrangebaseC & rc
'                            sheetrangebaseD = sheetrangebaseD & rc
'
'
'                            ' NEED TO COPY ORIGINAL PR FOR LINE 4 BELOW =========
''                            Sheets("Sheet2").Select
''                            ActiveSheet.Range(sheetrangetitle).Select
''                            ActiveSheet.Paste
'
'                            Set CellValue = sht.Range(sheetrange)
'                            Set CellOrder = sht.Range(sheetrangebaseC)
'                            Set CellType = sht.Range(sheetrangebaseD)
'                            Set CellPR = sht.Range(sheetrangetitle)
'                            CellOrder.value = "4"
'                            CellType.value = sTypeofCol
'                            CellValue.value = Sp4
'                            CellPR.value = PRn
'
'                    ElseIf (cSplits > 2) Then
'
'                        Sp1 = Left(r.value, cMax)
'                        Sp2 = Mid(r.value, cMax - CValue, cMax)
'                        Sp5 = Right(r.value, (cMax + cMax) - CValue)
'
'                            sheetrangebase = "B"
'                            sheetrangebaseC = "C"
'                            sheetrangebaseD = "D"
'                            sheetrangebasetitle = "A"
'                            sheetrange = sheetrangebase & rc
'                            sheetrangebase = sheetrangebase & rc
'                            sheetrangetitle = sheetrangebasetitle & rc
'                            sheetrangebaseC = sheetrangebaseC & rc
'                            sheetrangebaseD = sheetrangebaseD & rc
'
'                            Sheets("Sheet1").Select
'                            Range(sheetrangetitle).Select
'                            Selection.Copy
'
'                            Sheets("Sheet2").Select
'                            ActiveSheet.Range(sheetrangetitle).Select
'                            ActiveSheet.Paste
'
'                            Set CellValue = sht.Range(sheetrange)
'                            Set CellOrder = sht.Range(sheetrangebaseC)
'                            Set CellType = sht.Range(sheetrangebaseD)
'                            CellOrder.value = "1"
'                            CellType.value = sTypeofCol
'                            CellValue.value = Sp1
'
'                            ' ADD A SECOND LINE ITEM WITH ORDER 2 =====
'                            ' ADD ONE TO THE ROW COUNT ================
'
'                            'iValue = AddColumn(X, X, X, X)
'
'                            Sheets("Sheet1").Select
'                            Range(sheetrangetitle).Select
'                            Selection.Copy
'
'                            rc = rc + 1
'                            sheetrangebase = "B"
'                            sheetrangebaseC = "C"
'                            sheetrangebaseD = "D"
'                            sheetrangebasetitle = "A"
'                            sheetrange = sheetrangebase & rc
'                            sheetrangebase = sheetrangebase & rc
'                            sheetrangetitle = sheetrangebasetitle & rc
'                            sheetrangebaseC = sheetrangebaseC & rc
'                            sheetrangebaseD = sheetrangebaseD & rc
'
'
'                            ' NEED TO COPY ORIGINAL PR FOR LINE 2 BELOW =========
'                            Sheets("Sheet2").Select
'                            ActiveSheet.Range(sheetrangetitle).Select
'                            ActiveSheet.Paste
'
'                            Set CellValue = sht.Range(sheetrange)
'                            Set CellOrder = sht.Range(sheetrangebaseC)
'                            Set CellType = sht.Range(sheetrangebaseD)
'                            CellOrder.value = "2"
'                            CellType.value = sTypeofCol
'                            CellValue.value = Sp2
'
'                            ' ================================
'
'
'                            Sheets("Sheet1").Select
'                            Range(sheetrangetitle).Select
'                            Selection.Copy
'
'                            rc = rc + 1
'                            sheetrangebase = "B"
'                            sheetrangebaseC = "C"
'                            sheetrangebaseD = "D"
'                            sheetrangebasetitle = "A"
'                            sheetrange = sheetrangebase & rc
'                            sheetrangebase = sheetrangebase & rc
'                            sheetrangetitle = sheetrangebasetitle & rc
'                            sheetrangebaseC = sheetrangebaseC & rc
'                            sheetrangebaseD = sheetrangebaseD & rc
'
'
'                            ' NEED TO COPY ORIGINAL PR FOR LINE 2 BELOW =========
'                            Sheets("Sheet2").Select
'                            ActiveSheet.Range(sheetrangetitle).Select
'                            ActiveSheet.Paste
'
'                            Set CellValue = sht.Range(sheetrange)
'                            Set CellOrder = sht.Range(sheetrangebaseC)
'                            Set CellType = sht.Range(sheetrangebaseD)
'                            CellOrder.value = "3"
'                            CellType.value = sTypeofCol
'                            CellValue.value = Sp3
'
'                    ElseIf (cSplits > 1) Then
'
'                        Sp1 = Left(CTextValue, cMax)
'                        Sp2 = Right(CTextValue, CValue - cMax)
'
'                        ' SETUP CELLS ===========================
'
'''                        Sheets("Sheet2").Select
'''                        ActiveSheet.Range(sheetrange).Select
'''                        ActiveSheet.Paste
'
'                            sheetrangebase = "B"
'                            sheetrangebaseC = "C"
'                            sheetrangebaseD = "D"
'                            sheetrangebasetitle = "A"
'                            sheetrange = sheetrangebase & rc
'                            sheetrangebase = sheetrangebase & rc
'                            sheetrangetitle = sheetrangebasetitle & rc
'                            sheetrangebaseC = sheetrangebaseC & rc
'                            sheetrangebaseD = sheetrangebaseD & rc
'
'                            Sheets("Sheet1").Select
'                            Range(sheetrangetitle).Select
'                            Selection.Copy
'
'                            Sheets("Sheet2").Select
'                            ActiveSheet.Range(sheetrangetitle).Select
'                            ActiveSheet.Paste
'
'                            Set CellValue = sht.Range(sheetrange)
'                            Set CellOrder = sht.Range(sheetrangebaseC)
'                            Set CellType = sht.Range(sheetrangebaseD)
'                            CellOrder.value = "1"
'                            CellType.value = sTypeofCol
'                            CellValue.value = Sp1
'
'                            ' ADD A SECOND LINE ITEM WITH ORDER 2 =====
'                            ' ADD ONE TO THE ROW COUNT ================
'
'                            'iValue = AddColumn(X, X, X, X)
'
''                            Sheets("Sheet1").Select
''                            Range(sheetrangetitle).Select
''                            Selection.Copy
'
'                            rc = rc + 1
'                            sheetrangebase = "B"
'                            sheetrangebaseC = "C"
'                            sheetrangebaseD = "D"
'                            sheetrangebasetitle = "A"
'                            sheetrange = sheetrangebase & rc
'                            sheetrangebase = sheetrangebase & rc
'                            sheetrangetitle = sheetrangebasetitle & rc
'                            sheetrangebaseC = sheetrangebaseC & rc
'                            sheetrangebaseD = sheetrangebaseD & rc
'
'
'                            ' NEED TO COPY ORIGINAL PR FOR LINE 2 BELOW =========
'                            Sheets("Sheet2").Select
'                            ActiveSheet.Range(sheetrangetitle).Select
'                            ActiveSheet.Paste
'
'                            Set CellValue = sht.Range(sheetrange)
'                            Set CellOrder = sht.Range(sheetrangebaseC)
'                            Set CellType = sht.Range(sheetrangebaseD)
'                            CellOrder.value = "2"
'                            CellType.value = sTypeofCol
'                            CellValue.value = Sp2
'
'                    Else
'                    End If
'
'                Set CellOrder = sht.Range(sheetrangebaseC)
'                Set CellType = sht.Range(sheetrangebaseD)
'                'CellOrder.Value = "1"
'                CellType.value = sTypeofCol
'
'            Else
'                Sheets("Sheet1").Select
'                Range(sheetrange).Select
'                Selection.Copy
'
'                Sheets("Sheet2").Select
'                ActiveSheet.Range(sheetrange).Select
'                ActiveSheet.Paste
'
'                Set CellOrder = sht.Range(sheetrangebaseC)
'                Set CellType = sht.Range(sheetrangebaseD)
'                Set CellPR = sht.Range(sheetrangetitle)
'                CellOrder.value = "1"
'                CellType.value = sTypeofCol
'                CellPR.value = PRn
'            End If
'
''        ar = Split(r.Value, ",")
''
''        If UBound(ar) >= 0 Then r.Value = ar(0)
''        For i = UBound(ar) To 1 Step -1
''            r.EntireRow.Copy
''            r.Offset(1).EntireRow.Insert
''            r.Offset(1).Value = ar(i)
''        Next
'
'        PRn = PRn + 1
'        CurrentPRr = CurrentPRr + 1
'        ' RESET TO CURRENT ROW ================================
'        Set r = r.Offset(-1)
'        rc = rc + 1
'    Loop
    
End Sub

Sub AddColumn(rc As Integer, value As String, order As Integer, sptype As String)

                            Sheets("Sheet1").Select
                            Range(sheetrangetitle).Select
                            Selection.Copy
                            
                            rc = rc + 1
                            sheetrangebase = "B"
                            sheetrangebaseC = "C"
                            sheetrangebaseD = "D"
                            sheetrangebasetitle = "A"
                            sheetrange = sheetrangebase & rc
                            sheetrangebase = sheetrangebase & rc
                            sheetrangetitle = sheetrangebasetitle & rc
                            sheetrangebaseC = sheetrangebaseC & rc
                            sheetrangebaseD = sheetrangebaseD & rc

                            
                            ' NEED TO COPY ORIGINAL PR FOR LINE 2 BELOW =========
                            Sheets("Sheet2").Select
                            ActiveSheet.Range(sheetrangetitle).Select
                            ActiveSheet.Paste
                        
                            Set CellValue = sht.Range(sheetrange)
                            Set CellOrder = sht.Range(sheetrangebaseC)
                            Set CellType = sht.Range(sheetrangebaseD)
                            CellOrder.value = "2"
                            CellType.value = sTypeofCol
                            CellValue.value = Sp2

End Sub
