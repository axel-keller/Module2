' Module2Challenge


Sub AllSheets()

Dim ws As Worksheet
    
For Each ws In Worksheets
        
ws.Activate
       
Call Module2

Next ws

End Sub

Sub Module2()

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    Dim T As Long
    Dim Ticker As String
    Dim lastRow As Long
    Dim tsv As Double
    Dim tablerow As Integer
    Dim firstValue As Double
    Dim endValue As Double
    Dim qc As Double
    Dim pc As Double
    Dim maxpercentchange As Double
    Dim maxTicker As String
    Dim minpercentchange As Double
    Dim minTicker As String
    Dim maxVolume As Double
    Dim maxVolumeTicker As String
    
    tsv = 0
    tablerow = 2
    'Used Learning Assistant to get code for finding the last row within the dataset
    lastRow = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    'Used Learning Assistant to set the following variables 
    maxpercentchange = -999999
    minpercentchange = 999999
    maxVolume = -999999
    
    For T = 2 To lastRow
        tsv = tsv + Cells(T, 7).Value
        
        If Cells(T, 1).Value <> Cells(T - 1, 1).Value Then
            firstValue = Cells(T, 3).Value
        End If
        
        If T = lastRow Or Cells(T + 1, 1).Value <> Cells(T, 1).Value Then
            Ticker = Cells(T, 1).Value
            endValue = Cells(T, 6).Value
            
            qc = endValue - firstValue
            
            If firstValue <> 0 Then
                pc = (endValue - firstValue) / firstValue
            Else
                pc = 0
            End If
            
            Range("I" & tablerow).Value = Ticker
            Range("L" & tablerow).Value = tsv
            Range("J" & tablerow).Value = qc
            Range("K" & tablerow).Value = pc
            
            If pc > maxpercentchange Then
                maxpercentchange = pc
                maxTicker = Ticker
            End If
            
            If pc < minpercentchange Then
                minpercentchange = pc
                minTicker = Ticker
            End If
            
            If tsv > maxVolume Then
                maxVolume = tsv
                maxVolumeTicker = Ticker
            End If
            
            If pc > 0 Then
                Range("J" & tablerow).Interior.ColorIndex = 4
            ElseIf pc = 0 Then
                Range("J" & tablerow).Interior.ColorIndex = 2
            Else
                Range("J" & tablerow).Interior.ColorIndex = 3
            End If
            
            tablerow = tablerow + 1
            tsv = 0
        End If
        
        Range("K2:K" & tablerow).NumberFormat = "0.00%"
        Range("Q2:Q3").NumberFormat = "0.00%"
    Next T
    
    Range("P2").Value = maxTicker
    Range("Q2").Value = maxpercentchange
    
    Range("P3").Value = minTicker
    Range("Q3").Value = minpercentchange
    
    Range("P4").Value = maxVolumeTicker
    Range("Q4").Value = maxVolume
    
End Sub

