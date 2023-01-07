Attribute VB_Name = "Module1"
Sub TAKE_TWO()
'Establish variable for worksheet
Dim headers() As Variant
Dim Sheet As Worksheet
Dim Workbook As Workbook
Set Workbook = ActiveWorkbook

'Identify headers
headers() = Array("Ticker ", "Date ", "Open ", "High ", "Low", "Close ", "Volume ", " ", _
    "Ticker", "Yearly_Change", "Percent_Change", "Stock Volume", " ", " ", " ", "Ticker", "Value")
For Each Sheet In Workbook.Worksheets

    With Sheet
    .Rows(1).Value = " "
    For i = LBound(headers()) To UBound(headers())
    .Cells(1, 1 + i).Value = headers(i)
    
    Next i
    End With
   Next Sheet
   'Loop through entire workbook
For Each Sheet In Worksheets

' Set variables
Dim ticker_name As String
ticker_name = " "
Dim total_ticker_Volume As Double
total_ticker_Volume = 0
Dim Beg_Price As Double
Beg_Price = 0
Dim Yearly_Price_change As Double
Yearly_Price_change = 0
Dim yearly_percent_change As Double
yearly_percent_change = 0
Dim Max_Ticker As String
Max_Ticker = " "
Dim Min_Ticker As String
Min_Ticker = " "
Dim Max_Percent As Double
Max_Percent = 0
Dim Min_Percent As Double
Min_Percet = 0
Dim Max_Volume_Ticker As String
Max_Volume_Ticker = " "
Dim Max_Volume As Double
Max_Volume = 0
'Establish location of variable
Dim Summary_Table As Long
Summary_Table = 2
' Set Row count
Dim lastrow As Long
' Loop through all sheets
lastrow = Sheet.Cells(Rows.Count, 1).End(xlUp).Row


' set value of  beg stock for ticker
Beg_Price = Sheet.Cells(2, 3).Value
'loop from beginning
For i = 2 To lastrow
'check for same ticker name
If Sheet.Cells(i + 1, 1).Value <> Sheet.Cells(i, 1).Value Then
ticker_name = Sheet.Cells(i, 1).Value
'calc
end_price = Sheet.Cells(i, 6).Value
Yearly_Price_change = end_price - Beg_Price
'establish conditions
If Beg_Price <> 0 Then
Yearly_Price_Change_Percent = (Yearly_Price_change / Beg_Price) * 100
End If

'add ticker name total volume
total_ticker_Volume = total_ticker_Volume + Sheet.Cells(i, 7).Value
' push ticker name in the summary table column I
Sheet.Range("I" & Summary_Table).Value = ticker_name
'Print the year over year price change in summary table column J
Sheet.Range("J" & Summary_Table).Value = Yearly_Price_change
'color fill yearly price change; red for negative green for positive
If (Yearly_Price_change > 0) Then
Sheet.Range("J" & Summary_Table).Interior.ColorIndex = 4
ElseIf (Yearly_Price_change <= 0) Then
Sheet.Range("J" & Summary_Table).Interior.ColorIndex = 3
End If

'print the yearly change as percent
Sheet.Range("K" & Summary_Table).Value = (CStr(Yearly_Price_Change_Percent) & "%")
'print the total total volume summary table
Sheet.Range("L" & Summary_Table).Value = total_ticker_Volume
'Add 1 to summary table row
Summary_Table = Summary_Table + 1
'get next beg price
Beg_Price = Sheet.Cells(i + 1, 3).Value
' Calculations
If (Yearly_Price_Change_Percent > Max_Percent) Then
Max_Percent = Yearly_Price_Change_Percent
Max_Ticker_Name = ticker_name
End If

ElseIf (Yearly_Price_Change_Percent < Min_Percent) Then
Min_Percent = Yearly_Price_Change_Percent
Min_Ticker_Name = ticker_name

End If

If (total_ticker_Volume > Max_Volume) Then
Max_Volume = total_ticker_Volume
Max_Volume_Ticker_Name = ticker_name

' Reset Value
Yearly_Price_Change_Percent = 0
total_ticker_Volume = 0

'next ticker name enter enter new volume
total_ticker_Volume = total_ticker_Volume + Sheet.Cells(i, 7).Value
End If

Next i


'Print values in cells
Sheet.Range("Q2").Value = (CStr(Max_Percent) & "%")
Sheet.Range("Q3").Value = (CStr(Min_Percent) & "%")
Sheet.Range("P2").Value = Max_Ticker_Name
Sheet.Range("P3").Value = Min_Ticker_Name
Sheet.Range("Q4").Value = Max_Volume
Sheet.Range("O2").Value = "Greatest % Increase"
Sheet.Range("O3").Value = "Greatest % Decrease"
Sheet.Range("O4").Value = "Greatest Total Volume"

Next Sheet
End Sub


