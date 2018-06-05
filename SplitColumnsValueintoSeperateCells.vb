

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

    Dim sht As Worksheet
    Dim CellOrder As Range
    Dim CellType As Range
    'Dim CellOrder, CellType As String
    Set sht = Sheets("Sheet2")
    Range("A1").Select
    Dim r As Range, i As Long, ar
    Set r = Worksheets("Sheet1").Range("B1").End(xlDown)
    Range("A1").Select
        ' CREATE BASE STRINGS FOR PASTING INTO SHEET 2 --------

        
        ' PASTE PR NUMBER INTO A COLUMN -----------------
        'Range(sheetrangetitle).Select
        'Sheet2.Range(sheetrangetitle).Value = i
    
    Do While r.Row > 1
        CValue = Len(r.value)
        
        sheetrangebase = "B"
        sheetrangebaseC = "C"
        sheetrangebaseD = "D"
        sheetrangebasetitle = "A"
        sheetrange = sheetrangebase & rc
        sheetrangebase = sheetrangebase & rc
        sheetrangetitle = sheetrangebasetitle & rc
        sheetrangebaseC = sheetrangebaseC & rc
        sheetrangebaseD = sheetrangebaseD & rc
    
        Sheets("Sheet1").Select
        Range(sheetrangetitle).Select
         Selection.Copy
        
         Sheets("Sheet2").Select
        ActiveSheet.Range(sheetrangetitle).Select
        ActiveSheet.Paste
        
        ' PASTE PR NUMBER INTO A COLUMN -----------------
        
            If (CValue > cMax) Then
                cSplits = CValue / cMax
                Set CellOrder = sht.Range(sheetrangebaseC)
                Set CellType = sht.Range(sheetrangebaseD)
                CellOrder.value = 1
                
                    If (cSplits > 4) Then
                    
                        Sp1 = Left(r.value, cMax)
                        Sp2 = Mid(r.value, (cMax * 2) - CValue, cMax)
                        Sp3 = Mid(r.value, (cMax * 3) - CValue, cMax)
                        Sp4 = Mid(r.value, (cMax * 4) - CValue, cMax)
                        Sp5 = Right(r.value, (cMax * 5) - CValue)
                    
                    ElseIf (cSplits > 3) Then
                    
                        Sp1 = Left(r.value, cMax)
                        Sp2 = Mid(r.value, (cMax * 2) - CValue, cMax)
                        Sp3 = Mid(r.value, (cMax * 3) - CValue, cMax)
                        Sp4 = Mid(r.value, (cMax * 4) - CValue, cMax)
                    
                    ElseIf (cSplits > 2) Then
                    ElseIf (cSplits > 1) Then
                    
                        Sp1 = Left(r.value, cMax)
                        Sp2 = Right(r.value, (cMax * 2) - CValue)
                    
                        ' SETUP CELLS ===========================
        
''                        Sheets("Sheet2").Select
''                        ActiveSheet.Range(sheetrange).Select
''                        ActiveSheet.Paste
                            
                            sheetrangebase = "B"
                            sheetrangebaseC = "C"
                            sheetrangebaseD = "D"
                            sheetrangebasetitle = "A"
                            sheetrange = sheetrangebase & rc
                            sheetrangebase = sheetrangebase & rc
                            sheetrangetitle = sheetrangebasetitle & rc
                            sheetrangebaseC = sheetrangebaseC & rc
                            sheetrangebaseD = sheetrangebaseD & rc
        
                            Sheets("Sheet1").Select
                            Range(sheetrangetitle).Select
                            Selection.Copy
                            
                            Sheets("Sheet2").Select
                            ActiveSheet.Range(sheetrangetitle).Select
                            ActiveSheet.Paste
                        
                            Set CellValue = sht.Range(sheetrange)
                            Set CellOrder = sht.Range(sheetrangebaseC)
                            Set CellType = sht.Range(sheetrangebaseD)
                            CellOrder.value = "1"
                            CellType.value = "Progress Notes"
                            CellValue.value = Sp1
                        
                            ' ADD A SECOND LINE ITEM WITH ORDER 2 =====
                            ' ADD ONE TO THE ROW COUNT ================
                                                                                                         
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
                            CellType.value = "Progress Notes"
                            CellValue.value = Sp2
                        
                    Else
                    End If
                
                Set CellOrder = sht.Range(sheetrangebaseC)
                Set CellType = sht.Range(sheetrangebaseD)
                'CellOrder.Value = "1"
                CellType.value = "Progress Notes"
                
            Else
                Sheets("Sheet1").Select
                Range(sheetrange).Select
                Selection.Copy
                
                Sheets("Sheet2").Select
                ActiveSheet.Range(sheetrange).Select
                ActiveSheet.Paste
                
                Set CellOrder = sht.Range(sheetrangebaseC)
                Set CellType = sht.Range(sheetrangebaseD)
                CellOrder.value = "1"
                CellType.value = "Progress Notes"
            End If
        
'        ar = Split(r.Value, ",")
'
'        If UBound(ar) >= 0 Then r.Value = ar(0)
'        For i = UBound(ar) To 1 Step -1
'            r.EntireRow.Copy
'            r.Offset(1).EntireRow.Insert
'            r.Offset(1).Value = ar(i)
'        Next
        
        
        Set r = r.Offset(1)
            rc = rc + 1
    Loop
    
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
                            CellType.value = "Progress Notes"
                            CellValue.value = Sp2

End Sub
