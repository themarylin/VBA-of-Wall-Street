Attribute VB_Name = "main"
Sub main()
    'Declare variables
    Dim ws As Worksheet
    
    'start of the VBA looping procedure
    For Each ws In Sheets
        Call setHeaders(ws)
        Call findTotal(ws)
    Next ws
        
    
End Sub

Public Sub setHeaders(ws)
    ws.Activate
    
    Cells.Range("I1").Value = "Ticker"
    Cells.Range("J1").Value = "Yearly Change"
    Cells.Range("K1").Value = "Percent Change"
    Cells.Range("L1").Value = "Total Stock Volume"
    Cells.Range("O2").Value = "Greatest % Increase"
    Cells.Range("O3").Value = "Greatest % Decrease"
    Cells.Range("O4").Value = "Greatest Total Volume"
    Cells.Range("P1").Value = "Ticker"
    Cells.Range("Q1").Value = "Value"
    
End Sub

Public Sub findTotal(ws)
    ws.Activate
  
    'Declare variables
    Dim lastRow As Long
    Dim totalVolume As Double
    Dim ticker As String
    Dim counter As Integer
    Dim initial As Double
    Dim final As Double
    Dim change As Double
    Dim perchange As Double
    Dim maxPChange As Double
    Dim minPChange As Double
    Dim maxVolume As Double
    Dim pMaxTicker As String
    Dim pMinTicker As String
    Dim vMaxTicker As String
    
    'find lastRow
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    'sets printer variable row number
    printer = 2
    
    'set ticker and initial total
    ticker = Cells(2, 1).Value
    totalVolume = Cells(2, 7).Value
    initial = Cells(2, 3).Value
    maxPChange = 0
    minPChange = 0
    maxVolume = 0

    'Use while function to continue until last row
    For Row = 2 To lastRow
        If (Cells(Row + 1, 1).Value = ticker) Then
            totalVolume = totalVolume + Cells(Row + 1, 7).Value
        Else
            'set final close value and calculate change
            final = Cells(Row, 6).Value
            change = final - initial
            perchange = change / initial
            
            'print total values to excelsheet
            Cells(printer, 9).Value = ticker
            Cells(printer, 10).Value = change
            Cells(printer, 11).Value = perchange
            Cells(printer, 12).Value = totalVolume
            
            'format cells
            If Cells(printer, 10).Value <= 0 Then
                Cells(printer, 10).Interior.ColorIndex = 3
            Else
                Cells(printer, 10).Interior.ColorIndex = 4
            End If
            
            Cells(printer, 11).NumberFormat = "0.00%"
            
            'find largest and smallest changes and volume
            If Cells(printer, 11).Value > maxPChange Then
                maxPChange = Cells(printer, 11).Value
                pMaxTicker = Cells(printer, 9).Value
            ElseIf Cells(printer, 11).Value < minPChange Then
                minPChange = Cells(printer, 11).Value
                pMinTicker = Cells(printer, 9).Value
            End If
            
            If Cells(printer, 12).Value > maxVolume Then
                maxVolume = Cells(printer, 12).Value
                vMaxTicker = Cells(printer, 9).Value
            End If
            
            'update variables
            ticker = Cells(Row + 1, 1).Value
            totalVolume = 0
            printer = printer + 1
        End If
        
    Next Row
    
    'print max and min values
    Range("P2").Value = pMaxTicker
    Range("P3").Value = pMinTicker
    Range("P4").Value = vMaxTicker
    Range("Q2").Value = maxPChange
    Range("Q3").Value = minPChange
    Range("Q4").Value = maxVolume
    
    'format cells
    Range("Q2", "Q3").NumberFormat = "0.00%"
        
End Sub

