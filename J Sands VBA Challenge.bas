Attribute VB_Name = "VBA_Challenge"
Sub stock_analysis()
Application.ScreenUpdating = False

MsgBox ("The data analsyis will commence after closing this box and will take some time. A new message box will be displayed once complete")

For Each Ws In Worksheets

        'Find the lastrow
        lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'convert date column to number rather than text
        [B:B].Select
        With Selection
            .NumberFormat = "General"
            .Value = .Value
        End With

    'setup headers for each new column required
        Ws.Range("I1").Value = "Year"
        Ws.Range("J1").Value = "Ticker Symbol"
        Ws.Range("K1").Value = "Yearly Change"
        Ws.Range("L1").Value = "Percent Change"
        Ws.Range("M1").Value = "Total Stock Volume"

      'Set the unique ticker symbol (in order to recognise multiples of the same ticker code)
      Dim Ticker_Symbol As String
         
      'Set an initial variable for holding the Yearly Change
      Dim Yearly_Change As Double
    
      'Set an initial variable for holding the Yearly Percent Change
      Dim Percent_Change As Double
      
      'Set an initial variable for holding the Total Stock Volume
      Dim Total_SV As Double
      Total_SV = 0
  
      ' Keep track of the location for each ticker symbol in the summary table
      Dim Summary_Table_Row As Integer
      Summary_Table_Row = 2
    
      ' Loop through all ticker data
      For i = 2 To lastrow
    
        ' Check if we are still within the same ticker symbol, if it is not...
        If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
    
          ' Set the Ticker symbol
          Ticker_Symbol = Ws.Cells(i, 1).Value
          ' Print the Ticker Symbol in the Summary Table
          Ws.Range("J" & Summary_Table_Row).Value = Ticker_Symbol
          
          ' Add to the Total Stock Volume
          Total_SV = Total_SV + Ws.Cells(i, 7).Value
          
          ' Print the Total Stock Volume to the Summary Table
          Ws.Range("M" & Summary_Table_Row).Value = Total_SV
                
        'Calculate the opening and closing value of stock for the year
            Dim LTR As Range
            Dim lr As Long

            Ws.Range("A1:A" & lastrow).AutoFilter field:=1, Criteria1:=Ws.Range("J" & Summary_Table_Row)
            Set LTR = Ws.Range("A2:A" & lastrow).SpecialCells(xlCellTypeVisible)

            lr = Ws.Cells(Rows.Count, 1).SpecialCells(xlCellTypeVisible).End(xlDown).Row

            openvalue = LTR.Cells(1, 3).Value
            closevalue = Ws.Range("B" & lr).Offset(0, 4).Value
            
            'calculate the annual change from the opening price to the closing price
            Yearly_Change = closevalue - openvalue
            
            'calculate the percentage change from the opening price a the beginning of a given year to the closing price at the year end
            Percent_Change = Yearly_Change / openvalue
     
          ' Print the Yearly Change to the Summary Table
          Ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
            
          ' Print the Percent Change to the Summary Table
          Ws.Range("L" & Summary_Table_Row).Value = Percent_Change
    
          ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
                                 
          ' Reset the Stock Total
          Total_SV = 0
          Ws.ShowAllData
    
        ' If the cell immediately following a row is the same ticker symbol...
        Else
    
         'Add to the Total Stock Volume
          Total_SV = Total_SV + Cells(i, 7).Value
    
        End If
    
      Next i
    
    'determine greatest values
     
        'Find the lastrow for the grouped data
        lastrowGD = Ws.Cells(Rows.Count, 1).End(xlUp).Row
        'setup headers for each new column required
        Ws.Range("P2").Value = "Greatest % Increase"
        Ws.Range("P3").Value = "Greatest % Decrease"
        Ws.Range("P4").Value = "Greatest Total Volume"
        Ws.Range("Q1").Value = "Year"
        Ws.Range("R1").Value = "Ticker"
        Ws.Range("S1").Value = "Value"
      
      'declare variables
      Dim Rng As Range
      Dim RngVolume As Range
      Dim GreatestIncrease As Double
      Dim GreatestDecrease As Double
      Dim GreatestVolume As Double
      Dim Iindex As Long
      Dim Dindex As Long
      Dim Vindex As Long
          
          'set range from which to determine beginning and end
          Set Rng = Ws.Range("L1:L" & lastrowGD)
          Set RngVolume = Ws.Range("M1:M" & lastrowGD)
          
           'determine the greatest % decrease
          GreatestDecrease = Application.WorksheetFunction.Min(Rng)
          Dindex = Application.WorksheetFunction.Match(GreatestDecrease, Rng, 0)
          DGetAddr = Rng.Cells(Dindex).Address
          
          'determine the greatest % increase
          GreatestIncrease = Application.WorksheetFunction.Max(Rng)
          Iindex = Application.WorksheetFunction.Match(GreatestIncrease, Rng, 0)
          IGetAddr = Rng.Cells(Iindex).Address
          'determine the greatest total volume
          GreatestVolume = Application.WorksheetFunction.Max(RngVolume)
          Vindex = Application.WorksheetFunction.Match(GreatestVolume, RngVolume, 0)
          VGetAddr = RngVolume.Cells(Vindex).Address
          
        Ws.Range("S2").Value = GreatestIncrease
        Ws.Range("S3").Value = GreatestDecrease
        Ws.Range("S4").Value = GreatestVolume
        
        Ws.Range("Q2").Value = Range(IGetAddr).Offset(0, -3).Value
        Ws.Range("Q3").Value = Range(DGetAddr).Offset(0, -3).Value
        Ws.Range("Q4").Value = Range(VGetAddr).Offset(0, -4).Value
        
        Ws.Range("R2").Value = Range(IGetAddr).Offset(0, -2).Value
        Ws.Range("R3").Value = Range(DGetAddr).Offset(0, -2).Value
        Ws.Range("R4").Value = Range(VGetAddr).Offset(0, -3).Value

    'Adjust column widths and format cells
    Ws.Columns("J:S").AutoFit
    Ws.Columns("k:k").NumberFormat = "0.00_);(0.00)"
    Ws.Columns("l:l").NumberFormat = "0.00%"
    Ws.Columns("m:m").NumberFormat = "#,##0"
    Ws.Columns("Q:Q").NumberFormat = "@"
    Ws.Range("S1").HorizontalAlignment = xlRight
    Ws.Range("S2:S3").NumberFormat = "0.00%"
    Ws.Range("S4").NumberFormat = "#,##0"
   
   
        'Find the lastrow
        lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
   
    'conditionally format the positive change to green (4) and negative change to red (3)
    For CI = 2 To lastrow
    
            If Ws.Cells(CI, 11) > 0 Then
            Ws.Cells(CI, 11).Interior.ColorIndex = 4 'green
            ElseIf Ws.Cells(CI, 11) < 0 Then
            Ws.Cells(CI, 11).Interior.ColorIndex = 3 'red
            Else
            
            End If
            
    Next CI
    
Next Ws
    
Application.ScreenUpdating = True
MsgBox ("congratulations, data summary is complete")

End Sub

