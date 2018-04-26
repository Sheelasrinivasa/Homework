Attribute VB_Name = "Module3"
Function UniqueTickers()

'For WorkSheetNames
    Dim SheetArray(3) As String
'For Sheetcounter
    Dim Sheetcounter As Integer
'For TargetVolume
    Dim TargetVolume As Double
'For SourceVolume
    Dim SourceVolume As Long
'For SourceTicker cell reference
    Dim SourceTicker As String
'For TargetTicker cell reference
    Dim TargetTicker As String
'For last row count of Ticker column
    Dim AlastRow As Long
'For last row count of Target Ticker column
    Dim IlastRow As Long
    
    SheetArray(1) = "2015"
    SheetArray(2) = "2016"
    SheetArray(3) = "2014"

' Counter to iterate between sheets
    For Sheetcounter = 1 To 3

'Get Source Ticker last row count
        AlastRow = Sheets(SheetArray(Sheetcounter)).Range("A1048576").End(xlUp).Row
        Debug.Print (SheetArray(Sheetcounter))
'Get Target Ticker last row count
        TargetCounter = Sheets(SheetArray(Sheetcounter)).Range("I1048576").End(xlUp).Row
        Debug.Print (TargetCounter)
'Counter to iterate between source ticker column
        For SourceTickerCounter = 2 To AlastRow
        Debug.Print (SourceTickerCounter)

'To set Source ticker cell referene
            SourceTicker = Sheets(SheetArray(Sheetcounter)).Cells(SourceTickerCounter, 1)

'To set Target ticker cell reference
            TargetTicker = Sheets(SheetArray(Sheetcounter)).Cells(TargetCounter + 1, 9)

'To set Source Volume cell reference
            SourceVolume = Sheets(SheetArray(Sheetcounter)).Cells(SourceTickerCounter, 7)

'To set Target volume cell reference and reset counter
            TargetVolume = Sheets(SheetArray(Sheetcounter)).Cells(TargetCounter + 1, 10)
            
'Condition to check if two following tickers are not equal to
            If SourceTicker <> Cells(SourceTickerCounter + 1, 1) Then

'Then add existing source volume to existing target volume and pass value to reference cell
'Reset Targetcounter +1 to navigate to below cell
                TargetTicker = SourceTicker
                Sheets(SheetArray(Sheetcounter)).Cells(TargetCounter + 1, 9) = TargetTicker
                TargetVolume = TargetVolume + SourceVolume
                Sheets(SheetArray(Sheetcounter)).Cells(TargetCounter + 1, 10) = TargetVolume
                TargetCounter = TargetCounter + 1

'Condition to check if two following tickers are equal
            ElseIf Cells(SourceTickerCounter + 1, 1) = SourceTicker Then

'Update target ticker with source ticker and add new source volume to existing total target volume
'Pass new value to total column
                'TargetTicker = SourceTicker
                TargetVolume = TargetVolume + SourceVolume
                Sheets(SheetArray(Sheetcounter)).Cells(TargetCounter + 1, 10) = TargetVolume
                'Sheets(SheetArray(Sheetcounter)).Cells(TargetCounter + 1, 9) = TargetTicker
            
            End If
                
         Next SourceTickerCounter
                        
    Next Sheetcounter

End Function


       
