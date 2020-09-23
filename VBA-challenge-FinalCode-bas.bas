Attribute VB_Name = "Module1"

Sub VBAHomework()
    For Each ws In Worksheets
        ws.Activate
        Call CalculateSummary
    Next ws
End Sub

'----------------------------------------------------
Sub CalculateSummary()
    ' Start writing your code here
    Dim ticker As String
    Dim yearly_change As Double
    Dim yearly_change1 As Double
    Dim percent_change As Double
    Dim total_volume As Long
    Dim summary_row As Long
    Dim O1 As Double
    Dim O As Double
    Dim C1 As Double
    Dim C As Double
    Dim r As Range
    Dim rO As Range
    Dim rC As Range
    
   
    
    

        
        
    RowsNum = Cells(Rows.Count, 9).End(xlUp).Row
    summary_row = 2
    NumRows = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To NumRows
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            ticker = Cells(i, 1).Value
            total_volume = total_volume + Cells(i, 7).Value
            Range("I" & summary_row).Value = ticker
            Range("J" & summary_row).Value = O
            Range("K" & summary_row).Value = C
            Range("L" & summary_row).Value = yearly_change
            'Range("M" & summary_row).Value = percent_change
            Range("N" & summary_row).Value = total_volume
            summary_row = summary_row + 1
            
            
            Set rO = Range(Cells(2, 3), Cells(i, 3))
            O = rO.Rows(i).Value
            O1 = rO.Rows(1).Value
            Range("J" & 2).Value = O1
            
            Set r = Range(Cells(2, 11), Cells(i, 11))
            C1 = r.Rows(1).Value
            yearly_change1 = C1 - O1
            Range("L" & 2).Value = yearly_change1
            
        Else
            
            Set rC = Range(Cells(2, 6), Cells(i, 6))
            C = rC.Rows(i).Value
            total_volume = 0
            yearly_change = C - O
            
        End If
    Next i

    Debug.Print ActiveSheet.Name
Call ColorIndex
End Sub
Sub ColorIndex()


NumRows = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To NumRows
    If Cells(i, 12).Value > 0 Then
        Cells(i, 12).Interior.ColorIndex = 4
    ElseIf Cells(i, 12).Value < 0 Then
        Cells(i, 12).Interior.ColorIndex = 3
    Else
        Cells(i, 12).Interior.ColorIndex = 0
    End If
Next i
    
Call Percentage
End Sub
    
Sub Percentage()


Dim rO As Range




NumRows = Cells(Rows.Count, 12).End(xlUp).Row
For i = 2 To NumRows
    If Cells(i, 10).Value <> 0 Then
        Cells(i, 13).Value = (Cells(i, 12).Value / Cells(i, 10).Value) * 100
    Else
        Cells(i, 13).Value = Null
    End If
Next i

Call SetTitle
End Sub
'-------------------------------------------------------

Sub SetTitle()
    'Range("I:S").Value = ""
    'Range("I:S").Interior.ColorIndex = 0
' Set title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Open Value"
    Range("K1").Value = "Close Value"
    Range("L1").Value = "Yearly Change"
    Range("M1").Value = "Percent Change"
    Range("N1").Value = "Total Stock Volume"
    Range("R1").Value = "Ticker"
    Range("S1").Value = "Value"
    'this is for challenge only
    Range("Q2").Value = "Greatest % Increase"
    Range("Q3").Value = "Greatest % Decrease"
    Range("Q4").Value = "Greatest Total Volume"
    Range("I:Q").Columns.AutoFit
End Sub


