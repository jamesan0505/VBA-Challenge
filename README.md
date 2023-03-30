# VBA-Challenge
Sub VBAchallenge()

Dim Ticker_Name As String

Dim Yearly_Change As Double
Yearly_Change = 0

Dim Percent_Change As Double
Percent_Change = 0

Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim Open_Value As Double
Open_Value = Cells(2, 3).Value


For i = 2 To 759001
  If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Ticker_Name = Cells(i, 1).Value
    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
    Yearly_Change = Cells(i, 6).Value - Open_Value
    Percent_Change = (Cells(i, 6).Value - Open_Value) / Open_Value
    Range("I" & Summary_Table_Row).Value = Ticker_Name
    Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
    Range("J" & Summary_Table_Row).Value = Yearly_Change
    Range("K" & Summary_Table_Row).Value = Percent_Change
    Summary_Table_Row = Summary_Table_Row + 1
    Total_Stock_Volume = 0
    Open_Value = Cells(i + 1, 3).Value
    
 Else
    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
    
 End If
 

 Next i
 
For i = 2 To 3001
  If Cells(i, 10).Value >= 0 Then
    Cells(i, 10).Interior.ColorIndex = 4
  Else
    Cells(i, 10).Interior.ColorIndex = 3
    
  End If
 
Next i
 

  

 
 

End Sub
