# Unit-2-HW-Submission
Unit 2 Assignment - Easy + Challenge

----------------------------------------------

Sub TotalStockVolume():

'Set Variables

Dim ticker As String
Dim total As Double
Dim lastrow As Long
Dim i, j As Integer


'Loop Through All Worksheets
For Each ws In Worksheets


'Determine the last Row
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Add Header Names
 ws.Cells(1, 9).Value = "Ticker"
 ws.Cells(1, 10).Value = "Total Stock Value"


'Set Initial Numbers
 total = 0
 j = 2


'Loop Through Each Year
 For i = 2 To lastrow
 
     If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
         total = total + ws.Cells(i + 1, 7).Value
         
     Else
         ticker = ws.Cells(i, 1).Value
         ws.Cells(j, 9).Value = ticker
         ws.Cells(j, 10).Value = total
         
         'Add a New Row Next Ticker and Reset Total
         j = j + 1
         total = 0
         
     End If

 Next i
 
Next ws

End Sub

