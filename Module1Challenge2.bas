Attribute VB_Name = "Module1"
'worked on this with a learning assistant over zoom. Figured it out together.
Sub Stock()

   Dim i As Long '(i,j)
   Dim Total_Rows As Long
   Dim Ticker_Name As String
   Dim Total_Volume As Double
   Dim Start_Ticker As Long
   Dim Start_Closing As Double
   Dim Summary_Table_Row As Integer
   j = 0
   Summary_Table_Row = 2
   Total_Rows = Cells(Rows.Count, "A").End(xlUp).Row
   
   'Sheets("2018").Activate
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    
    Start_Ticker = 2
    
   For i = 2 To Total_Rows 'full rows breaks excel

       If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          'Start of a new ticker
          Start_Ticker = i + 1
        
          'Get the last volume
          Total_Volume = Total_Volume + Cells(i, 7).Value
          
          Start_Closing = Cells(i, 6).Value
          'MsgBox "New Ticker " & Cells(i + 1, 1).Value & " " & Start_Closing
          Range("I" & 2 + j).Value = Cells(i, 1).Value
          Range("J" & 2 + j).Value = Cells(i, 6).Value - Cells(Start_Ticker - 1, 3).Value
          Range("K" & 2 + j).Value = FormatPercent((Cells(i, 6).Value - Cells(Start_Ticker - 1, 3).Value) / Cells(Start_Ticker - 1, 3).Value)
          j = j + 1
          
         

      Else
         'Tally up the volume
          Total_Volume = Total_Volume + Cells(i, 7).Value
       End If
        
   Next i
   'MsgBox "Done"
   
End Sub


