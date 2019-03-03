Attribute VB_Name = "Module1"
' create subrtines for tasks that will be repeated
Sub color()
        
        'if value of cell <=0 then make it red
        If Selection.Value <= 0 Then
            Selection.Interior.color = RGB(255, 0, 0)
        'if value of cell <=0 then make it green
        Else
            Selection.Interior.color = RGB(0, 255, 0)
        End If
 
End Sub


'Needed before we start looping to grab values.
Sub SortData()
        Dim LastCell As String
            'Find the range starting from A1
            Range("A1").Select
            Selection.End(xlDown).Select
            Selection.End(xlToRight).Select
            ' put the end cell into a variable
            LastCell = Selection.Address
            'MsgBox (LastCell)
            ActiveSheet.Sort.SortFields.Clear
            'sort
        With ActiveSheet.Sort
             .SortFields.Add Key:=Range("A1", Range("A1").End(xlDown)), Order:=xlAscending
             .SortFields.Add Key:=Range("B1", Range("B1").End(xlDown)), Order:=xlAscending
             .SetRange Range("A1:" & LastCell)
             .Header = xlYes
             .Apply
        End With
End Sub


'This subrutine will create the summary after the data is sorted.
Sub CreateSummaryTable():
        '1. define variables
        Dim year As Integer
        Dim WorkingYear As Integer
        Dim ticker As String
        Dim workingTicker As String
        Dim MasterRow As Integer
        Dim WorkingVolume As Double
        Dim volume As Double
        Dim OpenPrice As Double
        Dim WorkingOpenPrice As Double
        Dim closePrice As Double
        Dim workingClosePrice As Double
        Dim InserRow As Integer
        Dim MaxPercentageDecrease As Double
        Dim MaxPercentageIncrease As Double
        Dim MaxVolume As Variant
        Dim IncreaseTicker As String
        Dim DecreaseTicker As String
        Dim VolumeTicker As String
        
        
        'set initial variables. This set of variables is designed to store the last commited set of variables.
        
        
        year = 0
        ticker = ""
        InserRow = 2
        OpenPrice = 0
        volume = 0
        closePrice = 0
        MaxPercentageDecrease = 0
        MaxPercentageIncrease = 0
        MaxVolume = 0
        
        
        ' loop through the rows and save the contents of each cel in different variable. Those will be the working variables.
        ' We will determine what to do, depending on whether the row contains a new set of year/Ticker or not.
        
        For year_loop = 2 To 1000000
        
        ' Make sure we do not process data if there is not data to process
                    If IsEmpty(Cells(year_loop, 2).Value) Then
                    Exit For
                    
                    Else
        
        
        'If there is data to process
        ' First, save the data of each cell into a variable. The current rows will be stored into WorkingVariables
                        WorkingYear = Left(Cells(year_loop, 2).Value, 4)
                        workingTicker = Cells(year_loop, 1).Value
                        WorkingOpenPrice = Cells(year_loop, 3).Value
                        workingClosePrice = Cells(year_loop, 6).Value
                        WorkingVolume = Cells(year_loop, 7).Value + volume
                    End If
                
        ' Evaluate if the year and the ticker are new or if they resamble the rows above
                    If (year <> WorkingYear Or ticker <> workingTicker) Then
                    
        'If the Year/Ticker is new, we have to record the data for the ended Year/Ticker into the summary table
                        'Cells(InserRow, 9) = year
                        Cells(InserRow, 10) = ticker
                        Cells(InserRow, 11) = closePrice - OpenPrice
                        
                        If OpenPrice = 0 Then
                            Cells(InserRow, 12) = 0
                            Else
                            Cells(InserRow, 12) = (closePrice - OpenPrice) / OpenPrice
                        End If
                        Cells(InserRow, 13) = WorkingVolume
                        'Cells(InserRow, 14) = OpenPrice
                        'Cells(InserRow, 15) = closePrice
        'format Price change calling sub color.
                        Cells(InserRow, 12).Select
                        Call color
     
        'Build Top Tikers table
        'once we have the summary for the ticker inserted into the summary table,  we have the data to compare to past tikers summaries.
        ' we need to save max % change, Min%change and max volume into their own variables to build the TopTikers table
                                If Cells(InserRow, 12) > MaxPercentageIncrease Then
                                    MaxPercentageIncrease = Cells(InserRow, 12)
                                    IncreaseTicker = ticker
                                End If
                                If Cells(InserRow, 12) < MaxPercentageDecrease Then
                                    MaxPercentageDecrease = Cells(InserRow, 12)
                                    DecreaseTicker = ticker
                                End If
                                If Cells(InserRow, 13) > MaxVolume Then
                                    MaxVolume = Cells(InserRow, 13)
                                    VolumeTicker = ticker
                                End If
                
        'To keep track of where to insert next, save the row number into a variable
                        InserRow = InserRow + 1
        'Then we have to reset volume and closing price as we need to re-start the count for the new year/ticker.
        'we also have to Save the openPrice as that will be needed for the summary table
        
                        volume = 0
                        closePrice = 0
                        OpenPrice = WorkingOpenPrice
                        
                        
                        
        'If the year/Ticker is the same than the row above, sum volume and record row closing price. Do not reset openPrice as we need the very fist one of the set.
                    Else
                        volume = WorkingVolume
                        closePrice = workingClosePrice
                
                    End If
                    
                    
                    
                    
            
        'Update the year/ ticker variables with the values we just processed, so that we can compare next row workingVariables to determine if the new row has new data
                year = WorkingYear
                ticker = workingTicker
                
        
                
            
        Next year_loop
    
        Cells(2, 18).Value = IncreaseTicker
        Cells(3, 18).Value = DecreaseTicker
        Cells(4, 18).Value = VolumeTicker
        
        Cells(2, 19).Value = MaxPercentageIncrease
        Cells(3, 19).Value = MaxPercentageDecrease
        Cells(4, 19).Value = MaxVolume
            
                
End Sub
'Once the summary table has been created, we need to select top by volume and change


Sub format()

   'Format the summary tables: add headers to the tables, delete extra row, format numbers
   
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "% change"
    Range("M1").Value = "Volume"
    Range("I2:O2").Select
    Selection.Delete Shift:=xlUp
    
    Columns("L:L").NumberFormat = "0.00%"
    Range("S2:S3").NumberFormat = "0.00%"
    
    
    'format the top tickers table
    Cells(2, 17).Value = "Greatest % increase"
        Cells(3, 17).Value = "Greatest % Decrease"
        Cells(4, 17).Value = "Greatest total volume"
        Cells(1, 18).Value = "Ticker"
        Cells(1, 19).Value = "Value"
        
   
    
End Sub

'This subroutine loops thorugh every sheet and executes the other subrutines
Sub ExecuteCreateTablesInEverySheet()
' loop from sheet to sheet and execute each macro
    For SheetCounter = 1 To ThisWorkbook.Sheets.Count
        ThisWorkbook.Sheets(SheetCounter).Activate
        ' macro1
        Call SortData
        'macro2
        Call CreateSummaryTable
        
        'macro 4
        Call format
    Range("a1").Select
    Next SheetCounter
 
End Sub

