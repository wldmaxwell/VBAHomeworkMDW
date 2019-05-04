 Sub easy():

    Dim ticker As String
    Dim volume As Double
    volume = 0

    Dim Summary_Table_Row As Integer
 

    Cells(1, 8).Value = "ticker"
    Cells(1, 9).Value = "Total Stock Volume"
  
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    Summary_Table_Row = 2

    For i = 2 To lastrow

      
      If Cells(i - 1, 1) = Cells(i, 1) And Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          

          ticker = Cells(i, 1).Value


          vol = vol + Cells(i, 7).Value


          Range("H" & Summary_Table_Row).Value = ticker

          Range("I" & Summary_Table_Row).Value = volume

          Summary_Table_Row = Summary_Table_Row + 1

          volume = 0


      Else

          volume = volume + Cells(i, 7).Value


      End If


    Next i

End Sub



