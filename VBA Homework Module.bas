Attribute VB_Name = "Module1"
Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call ticker
    Next
    Application.ScreenUpdating = True
End Sub



Sub ticker()

Dim label As String

Dim OpenPrice As Double
Dim ClosePrice As Double
Dim FinalPrice As Double
Dim PercentChange As Double

Dim MaxPercentCounter As Double
Dim MinPercentCounter As Double
Dim MaxVolume As Double
Dim MaxTickerSymbol As String
Dim MinTickerSymbol As String
Dim MaxVolumeTicker As String

'I am manually putting the Opening Value into OpenPrice
OpenPrice = Range("C2").Value

Dim Counter As Double
Counter = 0

Dim LabelLocation As Integer

LabelLocation = 2

'I am using this function to automatically detect the last row of data, instead of putting manually
last = Cells(Rows.Count, 1).End(xlUp).Row

'Putting all the Headings for the data output
Range("K1").Value = "Ticker"
Range("L1").Value = "Yearly Change"
Range("M1").Value = "Percent Change"
Range("N1").Value = "Total Stock Volume"

'Looping all the data rows
For i = 2 To last

    'If the row below current row is NOT the same value
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'Store the Ticker Value before the change into Label
        label = Cells(i, 1).Value
        
        'Print the Ticker Value into a specified location
        Cells(LabelLocation, 11).Value = label
              
        'Adding the last value to Counter before printing
        Counter = Counter + Cells(i, 7).Value
        
        'Print the Total Volume into the specified location
        Cells(LabelLocation, 14).Value = Counter
             
        'Putting the Closing value into ClosePrice
        ClosePrice = Cells(i, 6).Value
        
        'Calculating final price
        FinalPrice = ClosePrice - OpenPrice
        
        
        'Print the FinalPrice into the specified location(Yearly Change heading)
        Cells(LabelLocation, 12).Value = FinalPrice
        
        'Change the cell color based on the FinalPrice Value
        If FinalPrice < 0 Then
            Cells(LabelLocation, 12).Interior.ColorIndex = 3
            Else
            Cells(LabelLocation, 12).Interior.ColorIndex = 4
            
        End If
        
        'This fix division by 0 error.
            If OpenPrice = 0 Then
            PercentChange = 0
            Else
               
            'Calculating the Percent Change
             PercentChange = (ClosePrice - OpenPrice) / OpenPrice
        
            End If
        
        
        'Print the Percent Change into the specified location
        Cells(LabelLocation, 13).Value = PercentChange
        
                    
        'Convert into a percentage format
        Cells(LabelLocation, 13).NumberFormat = "0.00%"
        
         'Change the cell color based on the PercentChange Value.
         'This is ommited because interior color index change has been performed on Yearly Change.
        'If Cells(LabelLocation, 11).Value < 0 Then
            'Cells(LabelLocation, 11).Interior.ColorIndex = 3
            'Else
            'Cells(LabelLocation, 11).Interior.ColorIndex = 4
            
        'End If
        
        'Increase the specified location into the next row
        LabelLocation = LabelLocation + 1
        
        'Reset the volume counter
        Counter = 0
        
        'reset the OpenPrice
        OpenPrice = Cells(i + 1, 3).Value
        
        'formula to calculate percent change is (close price - open price)/open price
        
    
        
    Else
    
        'Add the volume on that row to the Counter
        Counter = Counter + Cells(i, 7).Value
        
       

    End If
    

Next i


'Putting "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
'Output data are in Column K to Column N

'USE the following variables for storing values.
'Dim MaxPercentCounter As Double
'Dim MinPercentCounter As Double
'Dim MaxVolume As Double

'Dim MaxTickerSymbol As String
'Dim MinTickerSymbol As String
'Dim MaxVolumeTicker As String

'I am manually putting the first value as PercentCounter
MaxPercentCounter = Range("M2").Value
MinPercentCounter = Range("M2").Value
MaxVolume = Range("N2").Value

'Loop through all the result
For h = 2 To last
    
    'Loop through to get Greatest Percent Change
    If MaxPercentCounter < Cells(h + 1, 13).Value Then
        MaxPercentCounter = Cells(h + 1, 13).Value
        MaxTickerSymbol = Cells(h + 1, 11).Value
    End If
    
    'Loop through to get the Lowest Percent Change
    If MinPercentCounter > Cells(h + 1, 13).Value Then
        MinPercentCounter = Cells(h + 1, 13).Value
        MinTickerSymbol = Cells(h + 1, 11).Value
    End If
    
    'Loop through to get the Greatest Volume
    If MaxVolume < Cells(h + 1, 14).Value Then
        MaxVolume = Cells(h + 1, 14).Value
        MaxVolumeTicker = Cells(h + 1, 11).Value
    End If
    
    
    
Next h

    Range("P2").Value = "Greatest % Increase"
    Range("P3").Value = "Greatest % Decrease"
    Range("P4").Value = "Greatest Total Volume"
    Range("Q1").Value = "Ticker"
    Range("R1").Value = "Value"
    
    
    'Printing the Result of Greatest Percent Change, Ticker Symbol, and Greatest Volume
    Range("R2").Value = MaxPercentCounter
    Range("Q2").Value = MaxTickerSymbol
    
    Range("R3").Value = MinPercentCounter
    Range("Q3").Value = MinTickerSymbol
    
    Range("R4").Value = MaxVolume
    Range("Q4").Value = MaxVolumeTicker
    
    'Formating the Column as
    Range("R2:R3").NumberFormat = "0.00%"
    
    'Formating bold font for headings
    Range("K1:N1").Font.Bold = True
    Range("K1:K5000").Font.Bold = True
    Range("Q1:R1").Font.Bold = True
    Range("P2:P4").Font.Bold = True
    

End Sub

Sub ClearAll()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Clear
    Next
    Application.ScreenUpdating = True
End Sub




'This is to clear
Sub Clear()

'I am using this function to automatically detect the last row of data, instead of putting manually
last = Cells(Rows.Count, 1).End(xlUp).Row


Range("K1:N5000").ClearContents
Range("P1:R4").ClearContents
Range("L1:N5000").Interior.ColorIndex = 0


End Sub


