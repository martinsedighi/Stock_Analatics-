Attribute VB_Name = "RibbonX_Code"
Sub stock()

  Dim total As Double
  Dim i As Long
  Dim change As Double
  Dim j As Integer
  Dim start As Long
  Dim rowCount As Long
  Dim days As Integer
  Dim averageChange As Double
  
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume "
    
    j = 0
    total = 0
    change = 0
    start = 2
    
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To rowCount
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    total = total + Cells(i, 7).Value
    If total = 0 Then
    
      Range("I" & 2 + j).Value = Cells(i, 1).Value
      Range("J" & 2 + j).Value = 0
      Range("K" & 2 + j).Value = "%" & 0
      Range("L" & 2 + j).Value = 0
    
   Else
   
   If Cells(start, 3) = 0 Then
     For find_Value = start To i
        If Cells(find_Value, 3) <> 0 Then
            start = find_Value
            Exit For
        End If
     Next find_Value
   End If
    
    change = (Cells(i, 6) - Cells(start, 3))
    percentChange = Round((change / Cells(start, 3) * 100), 2)
    
    start = i + 1
    
      Range("I" & 2 + j).Value = Cells(i, 1).Value
      Range("J" & 2 + j).Value = Round(change, 2)
      Range("K" & 2 + j).Value = "%" & percentChange
      Range("L" & 2 + j).Value = total
    
    If change > 0 Then
        Range("J" & 2 + j).Interior.ColorIndex = 4
      ElseIf change < 0 Then
        Range("J" & 2 + j).Interior.ColorIndex = 3
      Else
        Range("J" & 2 + j).Interior.ColorIndex = 0
    End If
    
  End If
        
       total = 0
       change = 0
       j = j + 1
       days = 0
       
       Else
          total = total + Cells(i, 7).Value
   End If
   
   Next i
End Sub

