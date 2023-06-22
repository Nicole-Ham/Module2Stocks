VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock()

    Dim readName As String
    Dim nextName As String
    Dim totalSV As Double
    Dim openPrice As Double
    Dim closePrice As Double
    
    Dim i As Long
    Dim lastRow As Long
    Dim curSheet As Worksheet
    
    For Each curSheet In ActiveWorkbook.Worksheets
    
        curSheet.Range("I1").Value = "Ticket"
        curSheet.Range("J1").Value = "Yearly Change"
        curSheet.Range("K1").Value = "Percent Change"
        curSheet.Range("L1").Value = "Total Stock Volume"
        
        maxIncP = 0
        maxDecP = 0
        maxTV = 0
        groupNo = 1
        totalSV = 0
        openPrice = curSheet.Cells(2, 3).Value
        lastRow = curSheet.Cells(curSheet.Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastRow
        
            readName = curSheet.Cells(i, 1).Value
            nextName = curSheet.Cells(i + 1, 1).Value
            
            If nextName = readName Then
                totalSV = totalSV + curSheet.Cells(i, 7).Value
            Else
                totalSV = totalSV + curSheet.Cells(i, 7).Value
                closePrice = Cells(i, 6).Value
                curSheet.Cells(groupNo + 1, 9).Value = readName
                curSheet.Cells(groupNo + 1, 12).Value = totalSV
                If totalSV > maxTV Then
                    maxTV = totalSV
                    maxTVT = readName
                End If
                
                curSheet.Cells(groupNo + 1, 10).Value = closePrice - openPrice
                curSheet.Cells(groupNo + 1, 11).Value = (closePrice - openPrice) / openPrice
                
                yrChange = closePrice - openPrice
                PctChange = yrChange / openPrice
                If PctChange > maxIncP Then
                   maxIncP = PctChange
                   maxIncT = readName
                End If
                
                If PctChange < maxDecP Then
                   maxDecP = PctChange
                  maxDecT = readName
                End If
                    
                If yrChange > 0 Then
                    curSheet.Cells(groupNo + 1, 10).Interior.ColorIndex = 14          'green
                    
                Else
                    curSheet.Cells(groupNo + 1, 10).Interior.ColorIndex = 9           'red
                   
                End If
                
                curSheet.Cells(groupNo + 1, 10).Value = yrChange
                curSheet.Cells(groupNo + 1, 11).Value = PctChange
                curSheet.Cells(groupNo + 1, 12).Value = totalSV
                
                totalSV = 0
                openPrice = curSheet.Cells(i + 1, 3).Value
                groupNo = groupNo + 1
                
            
            End If
            
        
        Next i
        
        curSheet.Range("O2").Value = "Greatest % Increase"
        curSheet.Range("O3").Value = "Greatest % Decrease"
        curSheet.Range("O4").Value = "Greatest Total Volume"
        curSheet.Range("P1").Value = "Ticker"
        curSheet.Range("Q1").Value = "Value"
        curSheet.Range("P2").Value = maxIncT
        curSheet.Range("P3").Value = maxDecT
        curSheet.Range("P4").Value = maxTVT
        curSheet.Range("Q2").Value = maxIncP
        curSheet.Range("Q3").Value = maxDecP
        curSheet.Range("Q4").Value = maxTV
        
        curSheet.Range("K1:K" & lastRow).NumberFormat = "0.00%"
        curSheet.Cells(2, 17).NumberFormat = "0.00%"
        curSheet.Cells(3, 17).NumberFormat = "0.00%"
        
    Next curSheet
    
End Sub

