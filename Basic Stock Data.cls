VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockAnalysis()
    Dim WS As Worksheet
    Dim Ticker As String
    Dim TotaV As Double
    Dim Summary_Table_Row As Integer

    For Each WS In Worksheets
    WS.Activate

    TotalV = 0

    Summary_Table_Row = 2
    
    WS.Range("I1").Value = "Ticker"
    WS.Range("J1").Value = "TotalV"
    
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To LastRow

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                Ticker = Cells(i, 1).Value

                TotalV = TotalV + Cells(i, 7).Value

                WS.Range("I" & Summary_Table_Row).Value = Ticker

                WS.Range("J" & Summary_Table_Row).Value = TotalV

                Summary_Table_Row = Summary_Table_Row + 1
                
                TotalV = 0

            Else

                TotalV = TotalV + Cells(i, 7).Value

            End If
              
        Next i

    Next WS

End Sub

