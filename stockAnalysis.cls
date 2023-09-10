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

For Each ws In Worksheets
'Variable Declarations

Dim Ticker As String
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStockVolume As Double
Dim LastRow As Long
Dim OutputRow As Long
Dim i As Long
Dim Rng As Range
Dim LastRow_OutputTable As Long

' Number of rows in each sheet

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Assign Headers to Output columns

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "YearlyChange"
ws.Cells(1, 11).Value = "PercentChange"
ws.Cells(1, 12).Value = "TotalStockVolume"

'Assign initial values to TotalStockVolume and  OpeningPrice

TotalStockVolume = 0
OpeningPrice = ws.Cells(2, 3).Value

'Assign initial value to output row

OutputRow = 2

'For loop through each year to calulate YearlyChange and Totla Stock Volume

For i = 2 To LastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                    Ticker = ws.Cells(i, 1).Value
                    ClosingPrice = ws.Cells(i, 6).Value
                    
                    YearlyChange = ClosingPrice - OpeningPrice
                    
                    
                    ' Calculate Total Stock Volume per Ticker
                    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                    
                    
                    'Assign values to Output Fileds
                    
                    ws.Cells(OutputRow, 9).Value = Ticker
                    ws.Cells(OutputRow, 10).Value = YearlyChange
                    
                    If YearlyChange >= 0 Then
                            ws.Cells(OutputRow, 10).Interior.Color = vbGreen
                            ElseIf YearlyChange < 0 Then
                            ws.Cells(OutputRow, 10).Interior.Color = vbRed
                    End If
                    
                    ws.Cells(OutputRow, 12).Value = TotalStockVolume
                    
                    'Check for OpeningPrice = 0
                    
                    If OpeningPrice = 0 Then
                                PercentChange = 0
                                Else
                                PercentChange = YearlyChange / OpeningPrice
                    End If
                    
                    ws.Cells(OutputRow, 11).Value = PercentChange
                    ws.Cells(OutputRow, 11).NumberFormat = "0.00%"
                    
                    'Increase Output Field Row by 1 to move to next row
                    
                    OutputRow = OutputRow + 1
                    
                    'Reset Total Stock Volume and Opening Price Values for next iteration
                    TotalStockVolume = 0
                    OpeningPrice = ws.Cells(i + 1, 3).Value
                    
                    Else
                    ' Keep adding total stock volume
                    
                    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
            End If

Next i

'Assign Header Values to colums for this set of output

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volumn"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

'Find number of rows of output field
LastRow_OutputTable = ws.Cells(Rows.Count, 9).End(xlUp).Row


For i = 2 To LastRow_OutputTable
                        
                    'Find Greatest % Increase
                    
            If ws.Cells(i, 11).Value = WorksheetFunction.Max(Range("K2:K" & LastRow_OutputTable)) Then
            
                        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                        ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                        ws.Cells(2, 17).NumberFormat = "0.00%"
                        
                         'Find Greatest % Decrease
                ElseIf ws.Cells(i, 11).Value = WorksheetFunction.Min(Range("K2:K" & LastRow_OutputTable)) Then
            
                        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                        ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                        ws.Cells(3, 17).NumberFormat = "0.00%"
                        
                          'Find Gratest Total Volume
                          
                ElseIf ws.Cells(i, 12).Value = WorksheetFunction.Max(Range("L2:L" & LastRow_OutputTable)) Then
            
                        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                        ws.Cells(4, 17).Value = ws.Cells(i, 12).Value

                    
            End If

Next i


Next ws

End Sub



