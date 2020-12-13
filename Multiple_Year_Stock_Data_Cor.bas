Attribute VB_Name = "Module1"
Public Sub Wksht_Loops()

'Creating Worksheet Loops Through Entire Workbook
       
'Setting index for loop through Worksheets
    Dim i As Integer

'Setting up loop sequence
    For i = 1 To Worksheets.Count
    Worksheets(i).Select

'List of Subroutines that run through Worksheets loop
    Multi_Yr_Stk
    Summary_Table_Calc
    Formating
    
'Calling for next sheet
    Next i

End Sub

'Setting up Subroutine for creating sorting table for Ticker, Yearly Change, Percent Change and Total Stock Volume
Public Sub Multi_Yr_Stk()

'Defining column headers for results
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
'Defining column width for designated columns
    Columns("J:K").ColumnWidth = 12.5
    Columns("L:L").ColumnWidth = 14
    Columns("O:O").ColumnWidth = 18
    Columns("P:Q").ColumnWidth = 11
    
'Dimensioning parameters and variables for memory space creation
    Dim Ticker As String                'Defining Ticker header as String type
    Dim Yearly_Change As Double         'Defining Yearly Change as Double type
    Yearly_Change = 0
    
    Dim StartRow As Double           'Defining Starting Row in the sequence as Double type
    StartRow = 2                     'Starter data row in data set
    
    Dim Percent_Change As Double        'Defining Percent Change as Double type
    Percent_Change = 0
    
    Dim Tot_Stk_Vol As Double             'Defining Total Stock Volume as Long type
    Tot_Stk_Vol = 0
    
    Dim Open_Price As Double
    Open_Price = 0
    
    Dim Close_Price As Double
    Close_Price = 0
    
    'Dim LastRow As Long                 'Defining Individual last row as Long type
    Dim Summary_Row As Integer             'Defining Summary Row for iteration as Long type
    Summary_Row = 2                     'Initializing iterated row
    
    'Dim i As Long                       'Defining iterator as Long type
    Dim MaxValue As Double              'Defining Greatest Percent Increase as Double type
    Dim MinValue As Double              'Defining Greatest Percent Decrease as Double type
    Dim MaxTotVol As Double             'Defining Maximum Total Volume as Double type
            
'Initializing key rows for iteration
     
'Defining last row in the data set
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
                       
'Initiating For loop for iterating Columns "I" to "L"
    For i = 2 To LastRow                'Defining the rows for iteration
            
        ' Ticker identification
        Ticker = Cells(i, 1).Value
        
        ' Entering values for Opening price for each ticker
        Open_Price = Cells(StartRow, 3).Value
        
        ' Entering values for Closing price for each ticker
         Close_Price = Cells(i, 6).Value
        
        ' Calculating yearly change from opening and closing prices
        Yearly_Change = Close_Price - Open_Price
    
    ' Setting conditional statement for computing non-zero denominator percentage change
        If Open_Price = 0 Then
      
            Percent_Change = 0
            
        Else
            
            Percent_Change = Yearly_Change / Open_Price
        
      
        End If
        
     ' Setting conditions for color changes for each respective cell in the Yearly change column, red and green respectively
        If (Yearly_Change > 0) Then
            Cells(Summary_Row, 10).Interior.ColorIndex = 4                      'Filling negative value cells in Yearly Change column with red color
        Else
            Cells(Summary_Row, 10).Interior.ColorIndex = 3                      'Filling positive value cells in Yearly Change column with green color
        End If
            
'Defining conditions for iteration
        If Cells(i + 1, 1).Value <> Ticker Then
           
        ' Set the Brand name
        StartRow = i + 1
            
            'Ticker = Cells(i, 1).Value                                              'Creating the values for the "I" Column with iteration results
            
        Tot_Stk_Vol = Tot_Stk_Vol + Cells(i, 7).Value
            
                
            ' Print the Ticker entries in the Summary Table
            Range("I" & Summary_Row).Value = Ticker

            ' Print the Ticker Volume to the Summary Table
            Range("L" & Summary_Row).Value = Tot_Stk_Vol
      
            ' Print the Yearly Change to the Summary Table
            Range("J" & Summary_Row).Value = Yearly_Change
      
            ' Print the Percentage Change to the Summary Table
            Range("K" & Summary_Row).Value = Percent_Change
                
            'Reinitializing iterative variables
            Summary_Row = Summary_Row + 1                                       'Step increase in iterated row
            Yearly_Change = 0
            Percent_Change = 0                                                  'Reinitializing Percent Change iterator
            Tot_Stk_Vol = 0                                                     'Reinitializing Total Stock Volumer iterator
       
       ' If earlier condition is not met, then the condition below holds. This provides for the entries in
       ' each summary cells
        Else

      ' Add to the Ticker Total
            Tot_Stk_Vol = Tot_Stk_Vol + Cells(i, 7).Value
        
        End If
    
    Next i

End Sub

'Computing values for Summary table; creating new subroutine
Public Sub Summary_Table_Calc()

'Calculating maximum percent change value and corresponding Ticker
    MaxValue = Application.WorksheetFunction.Max(Range("K:K"))                              'Calculating Greatest Percent increase
    Range("P2") = Application.WorksheetFunction.XLookup(MaxValue, [K:K], [I:I], "None")     'Conducting lookup for Ticker that matches Greatest Percent Increase
    Range("Q2").Value = MaxValue                                                            'Outputing maximum or greatest percent increase to appropriate cell
                    
'Calculating minimum percent change value and corresponding Ticker
    MinValue = Application.WorksheetFunction.Min(Range("K:K"))                              'Calculating Greatest Percent decrease
    Range("P3") = Application.WorksheetFunction.XLookup(MinValue, [K:K], [I:I], "None")     'Conducting lookup for Ticker that matches Greatest Percent Decrease
    Range("Q3").Value = MinValue                                                            'Outputing minimum or greatest percent decrease to appropriate cell
     
'Calculating Greatest Total Volume
    MaxTotVol = Application.WorksheetFunction.Max(Range("L:L"))                             'Calculating Greatest Total Volume
    Range("Q4").Value = MaxTotVol                                                           'Outputing maximum or greatest Total Volume
    Range("P4") = Application.WorksheetFunction.XLookup(MaxTotVol, [L:L], [I:I], "None")    'Conducting lookup for Greatest Total Volume
 
End Sub

'Creating formating for final presentation
Public Sub Formating()

'Formating Summary Table
    Cells(2, 17).Select
    Selection.NumberFormat = "#,##0.00%"                   'Comma and percentage formating for values in column Q for Greatest Percent Increase
    Cells(3, 17).Select
    Selection.NumberFormat = "#,##0.00%"                   'Comma and percentage formating for values in column Q for Greatest Percent Decrease
    Range("Q4").NumberFormat = "#,###,##0"                 'Comma formating for values in column Q for Greatest Total Volume
    Columns("L:L").Select
    Selection.NumberFormat = "#,###,##0"                   'Comma formating for values in column L
    Columns("Q:Q").ColumnWidth = 14
    
    

End Sub
