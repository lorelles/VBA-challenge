Attribute VB_Name = "Module1"
'' Create a script that will loop through all the stocks for one year and output the following information:
'' 1. Ticker symbol
'' 2. Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'' 3. The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'' 4. The total stock volume of the stock
'' Use conditional formatting to highlight positive change in green and negative change in red.
'' I received assistance in from a Learning Assistant


Sub stock_data()

    ' Loop through all worksheets
    'Dim ws As Worksheet
    For Each ws In Worksheets
        
        ' Label new cells and autofit to display data
         ws.Cells(1, 9).Value = "Ticker"
         ws.Cells(1, 10).Value = "Yearly Change"
         ws.Cells(1, 11).Value = "Percent Change"
         ws.Cells(1, 12).Value = "Total Stock Volume"
         ws.Columns("I:L").AutoFit
       
        Dim rowcount As LongLong
    
        ' Set variable for Ticker
        Dim Ticker As String
    
        ' Set variable for holding Yearly Change
        Dim Change As Double
        Change = 0
        
        'Set variable for holding Percent Change
        Dim PercentChange As Double
        
        'Set variable for holding total stock volume
        Dim Total As LongLong
    
        ' Keep track of the location for each ticker symbol in the summary table
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
    
        ' Determine the last column number
        rowcount = Cells(Rows.Count, 1).End(xlUp).Row
                    
        ' For Loop to iterate through whole worksheet
        For i = 2 To rowcount
                                    
            ' Check if still in same Ticker symbol, if ticker changes then print results
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                ' Store Total as variable
                Total = Total + Cells(i, 7).Value
            
                ' Set total volume to 0
                If Total = 0 Then
                
                   ' Print the results
                   Range("I" & 2 + j).Value = Cells(i, 1).Value
                   Range("J" & 2 + j).Value = 0
                   Range("K" & 2 + j).Value = "%" & 0
                   Range("L" & 2 + j).Value = 0
               
               Else
                   
                   ' Find First non zero starting value (I received assistance from Learning Assistant)
                       If Cells(i, 3) = 0 Then
                       For find_value = Start To i
                           If Cells(find_value, 3).Value <> 0 Then
                               Start = find_value
                               Exit For
                           End If
                        Next find_value
                        
                   End If
                   
                   ' Calculate Percent Change (I received assistance from Learning Assistant)
                   Change = (Cells(i, 6) - Cells(i, 3))
                   PercentChange = Round((Change / Cells(i, 3) * 100), 2)
                   
                   ' Now start the next stock ticker
                   Start = i + 1
                   
                   ' Then print the results
                   Range("I" & 2 + j).Value = Cells(i, 1).Value
                   Range("J" & 2 + j).Value = Round(Change, 2)
                   Range("K" & 2 + j).Value = "%" & PercentChange
                   Range("L" & 2 + j).Value = Total
                   
                   ' Set the colors so positive is in green and negative is in red
                   Select Case Change
                       Case Is > 0
                           Range("J" & 2 + j).Interior.ColorIndex = 4
                       Case Is < 0
                           Range("J" & 2 + j).Interior.ColorIndex = 3
                       Case Else
                           Range("J" & 2 + j).Interior.ColorIndex = 0
                   End Select
                   
               End If
               
               ' Now reset the variables for each new stock ticker
               Total = 0
               Change = 0
               j = j + 1
               Days = 0
               
           ' If ticker is still the same add results
           Else
           
               Total = Total + Cells(i, 7).Value
               
           End If
        
        Next i
        
        Next ws
    
    
 End Sub
 
    




