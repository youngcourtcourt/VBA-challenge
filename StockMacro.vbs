Attribute VB_Name = "StockMacro"
Sub StockMacro()

Dim ticker As String

Dim openingprice As Double
openingprice = 0

Dim closingprice As Double
closingprice = 0

Dim yearlychange As Double
yearlychange = 0

Dim totalvolume As Double
totalvolume = 0

Dim column As Long
column = 1

Dim n As Long
n = 0

Dim m As Long
m = 0

lastws = ThisWorkbook.Worksheets.Count

For j = 1 To lastws

Worksheets(j).Select

Dim summarytablerow As Integer
summarytablerow = 2

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Range("J1") = "Ticker"

Range("K1") = "Yearly Change"

Range("L1") = "Percent Change"

Range("M1") = "Total Stock Volume"

    For i = 2 To lastrow
        If Cells(i + 1, column).Value <> Cells(i, column).Value Then
        
            m = i - n
            
            ticker = Cells(i, 1).Value
            
            openingprice = Cells(m, 3).Value
            
            closingprice = Cells(i, 6).Value
            
            yearlychange = closingprice - openingprice
            
            totalvolume = totalvolume + Cells(i, 7).Value
            
            Range("J" & summarytablerow) = Cells(i, 1).Value
            
            Range("K" & summarytablerow) = yearlychange
    
            If yearlychange > 0 Then
                Range("K" & summarytablerow).Interior.ColorIndex = 4
                
                ElseIf yearlychange < 0 Then
                Range("K" & summarytablerow).Interior.ColorIndex = 3
            
            End If
            
            If openingprice = 0 Then
            
            Range("L" & summarytablerow) = Format(openingprice, "Percent")
            
            Else
            
            Range("L" & summarytablerow) = Format(((closingprice - openingprice) / openingprice), "Percent")
            
            End If
            
            Range("M" & summarytablerow) = totalvolume
            
            summarytablerow = summarytablerow + 1
            
            totalvolume = 0
            
            n = 0
            
            m = 0
            
        Else
        
            totalvolume = totalvolume + Cells(i, 7).Value
            
            n = n + 1
        End If
    Next i
 Next j
End Sub
