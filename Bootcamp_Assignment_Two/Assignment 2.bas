Attribute VB_Name = "Module1"
Sub ticker()

    Sheets.Add.Name = "Combined_Data"
    Sheets("Combined_Data").Move Before:=Sheets(1)
    Set combined_sheet = Worksheets("Combined_Data")
    
   For Each ws In Worksheets
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Change"
    Range("L1").Value = "Total Stock Volumn"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volumn"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Percentage"
    
    Dim last As Long
    last = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).row
    
    Dim lastRowYear As Double
    lastRowYear = ws.Cells(Rows.Count, "A").End(xlUp).row - 1
    combined_sheet.Range("A" & last & ":Q" & ((lastRowYear - 1) + last)).Value = ws.Range("A2:Q" & (lastRowYear + 1)).Value
    
    Columns("J:J").NumberFormat = "General"
    
    Set colRangeJ = ws.Range("J:J")
    colRangeJ.Columns.AutoFit
    
    Set colRangeK = ws.Range("K:K")
    colRangeK.NumberFormat = "0.00%"
    colRangeK.Columns.AutoFit
    
    Set colRangeL = ws.Range("L:L")
    colRangeL.Columns.AutoFit
    
    Range("K1").Value = "Percentage Change"
    
    Dim row As Integer
    row = 2
    
    Dim year As Double
    Dim day As Double
    year = 0
    day = 0
    
    Dim first_day As Long
    Dim Total_Volumn As Double
    Total_Volumn = 0
    
    
    For i = 2 To last
    
        Dim morning As Double
        morning = Cells(i, 3).Value
        Dim night As Double
        night = Cells(i, 6).Value
        Dim Volumn As Double
        
        If (Cells(i, 2).Value = 20180102 Or 20190102 Or 20200102) Then
            first_day = Cells(i, 3).Value
        End If
    
        If (Cells(i, 1).Value <> Cells(i + 1, 1).Value) Then
            Cells(row, 9).Value = Cells(i, 1).Value
            
            day = night - morning
            Range("J" & row).Value = year + day
            Dim Total_Change As Double
            Total_Change = Range("J" & row).Value
            
            If (Total_Change >= 0) Then
                Range("J" & row).Interior.ColorIndex = 4
            ElseIf (Total_Change < 0) Then
                Range("J" & row).Interior.ColorIndex = 3
            End If
            
            Dim percent As Double
            Dim change As Double
            percent = Range("J" & row).Value
            change = first_day + percent
            Range("K" & row).Value = ((change - first_day) / first_day)
            
            Volumn = Cells(i, 7).Value
            Total_Volumn = Total_Volumn + Volumn
            Range("L" & row).Value = Total_Volumn

            row = row + 1
            year = 0
            first_day = 0
            Total_Volumn = 0

        Else
        
           
            Volumn = Cells(i, 7).Value
            
            day = night - morning
            year = year + day
            Total_Volumn = Total_Volumn + Volumn
            
        End If
        
        Volumn = 0
        day = 0
        
    Next i
    
    
    Next ws
    
    combined_sheet.Range("A1:K1").Value = Sheets(2).Range("A1:K1").Value
    combined_sheet.Columns("A:K").AutoFit
    
End Sub

Sub percent()
    Dim greatest As Double
    Dim worst As Double
    Dim Max_Volumn As Double
    Dim last As Long
    last = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).row
    
    greatest = WorksheetFunction.Max(Range("K2:K" & last))
    worst = WorksheetFunction.Min(Range("K2:K" & last))
    Max_Volumn = WorksheetFunction.Max(Range("L2:L" & last))
    
    greatest = Range("Q2").Value
    worst = Range("Q3").Value
    Max_Volumn = Range("Q4").Value
    
    For i = 2 To last
        If (Cells(i, 10).Value = greatest) Then
            Cells(i, 1).Value = Range("P2").Value
            Cells(i, 10).Value = Range("Q2").Value
        ElseIf (Cells(i, 10).Value = worst) Then
            Cells(i, 1).Value = Range("P3").Value
            Cells(i, 10).Value = Range("Q3").Value
        ElseIf (Cells(i, 11).Value = worst) Then
            Cells(i, 1).Value = Range("P4").Value
            Cells(i, 11).Value = Range("Q4").Value
        End If
    
    Next i
        
End Sub


