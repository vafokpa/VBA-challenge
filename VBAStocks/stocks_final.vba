Sub stonksFinal()

'creating a variable for the ticker
    Dim ticker As String

'creating a variable for the opening stock price
    Dim open_stock As Double
    open_stock = 0

'creating a variable for the closing stock price
    Dim close_stock As Double
    close_stock = 0

'creating a variable for the volume
    Dim volume As Double
    volume = 0

' creating a variable for the summary rows
    Dim summary_row As Integer
    summary_row = 2

'Loop through tickers
For i = 2 To Range("A2").End(xlDown).Row

    volume = volume + Cells(i, 7).Value

    'capturing the opening stock value
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
    open_stock = Cells(i, 3).Value
    
    End If

    'Checking to see if we are changing ticker (i.e. at a border)
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'Set ticker name
        ticker = Cells(i, 1).Value
        
        'Capture the closing stock value
        close_stock = Cells(i, 6).Value
       
        'If the ticker is changing I want copy over my new volume to the summary page
        Cells(summary_row, 12).Value = volume

        'Print the ticker to the summary
        Cells(summary_row, 9).Value = ticker
        
        'Print the change in stocks
        Cells(summary_row, 10).Value = close_stock - open_stock
        
        'make sure we don't divide by zero
        If open_stock = 0 Then
        Cells(summary_row, 11).Value = 0

        End If
        
        'make sure we don't divide by 0 again
        If open_stock <> 0 Then
        'Print the percent change in stocks
        Cells(summary_row, 11).Value = (close_stock - open_stock) / open_stock

        End If

        'Add one to the summary_row
        summary_row = summary_row + 1

        'reset volume
        volume = 0
        
        'reset close_stock
        close_stock = 0
        'reset open_stock
        open_stock = 0

    Else
        'Add to the volume
        Cells(summary_row, 12).Value = volume
    
    End If
    
Next i


End Sub

