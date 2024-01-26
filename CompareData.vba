Sub CompareData()

    ' Varibels
    Dim wsSource As Worksheet, wsResult As Worksheet
    Dim firstRow As Long
    Dim lastRow As Long, lastCol As Long
    Dim firstPersonIndex As Long, secondPersonIndex As Long
    Dim dataIndex As Long
    Dim sameDataCount As Long, numCellsChecked As Long
    Dim printIndex As Long
    Dim match As Boolean
    
    Dim firstPersonName As String, secondPersonName As String
    Dim firstPersonTele As String, secondPersonTele As String
    
    ' Set manually
    firstRow = 2
    printIndex = 1
    lastCol = Columns("AP").Column
   
    ' Set source and result sheets
    Set wsSource = ThisWorkbook.Sheets("Sheet1")
    
    ' Set Known Cells
    Dim male As Range
    Set male = wsSource.Range("AU3")
    Dim female As Range
    Set female = wsSource.Range("AU4")
    Dim lookForFemale As Range
    Set lookForFemale = wsSource.Range("AU5")
    Dim lookForMale As Range
    Set lookForMale = wsSource.Range("AU6")
    Dim lookForBoth As Range
    Set lookForBoth = wsSource.Range("AU7")
   
    ' Add a new sheet for results
    Set wsResult = Sheets.Add(After:=Sheets(Sheets.Count))
    wsResult.Name = "Results" & (Sheets.Count)
   
    ' Find the last row with data in the source sheet
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
   
    ' Loop through every 2 persons for each one
    For firstPersonIndex = firstRow To lastRow
    
        ' Print the first perindex
        printIndex = printIndex + 1
        wsResult.Cells(printIndex, 1).Value = firstPersonIndex - 2
        wsResult.Cells(printIndex, 2).Value = "---------"
        wsResult.Cells(printIndex, 3).Value = wsSource.Cells(firstPersonIndex, 2).Value
        wsResult.Cells(printIndex, 4).Value = "---------"
        printIndex = printIndex + 1
        
        For secondPersonIndex = firstRow To lastRow
        
            ' Print the second per index
            ' wsResult.Cells(printIndex, 1).Value = wsSource.Cells(secondPersonIndex, 2).Value
            ' printIndex = printIndex + 1
        
            ' --- Check Must Cells ---
             match = True
             
             ' Check if gender is good
             If wsSource.Cells(firstPersonIndex, 8).Value = lookForFemale.Value And wsSource.Cells(secondPersonIndex, 5).Value = male.Value Then
                match = False
            End If
            If wsSource.Cells(firstPersonIndex, 8).Value = lookForMale.Value And wsSource.Cells(secondPersonIndex, 5).Value = female.Value Then
                match = False
            End If
            If wsSource.Cells(secondPersonIndex, 8).Value = lookForFemale.Value And wsSource.Cells(firstPersonIndex, 5).Value = male.Value Then
                match = False
            End If
            If wsSource.Cells(secondPersonIndex, 8).Value = lookForMale.Value And wsSource.Cells(firstPersonIndex, 5).Value = female.Value Then
                match = False
            End If
            
            ' Check if hookap is good
            If Not (wsSource.Cells(firstPersonIndex, 7).Value = wsSource.Cells(secondPersonIndex, 7).Value) Then
                match = False
            End If
            
            ' Check if age is good
            If wsSource.Cells(firstPersonIndex, 5).Value = male.Value And wsSource.Cells(secondPersonIndex, 5).Value = female.Value And CInt(wsSource.Cells(firstPersonIndex, 6).Value) < CInt(wsSource.Cells(secondPersonIndex, 6).Value) Then
                match = False
            End If
            If wsSource.Cells(secondPersonIndex, 5).Value = male.Value And wsSource.Cells(firstPersonIndex, 5).Value = female.Value And CInt(wsSource.Cells(secondPersonIndex, 6).Value) < CInt(wsSource.Cells(firstPersonIndex, 6).Value) Then
                match = False
            End If
            
            ' --- Check Data Count ---
            
            If match Then
                 ' Reset sameDataCount for each pair of persons
                 sameDataCount = 0
                 numCellsChecked = 0
                
                 ' Get the names and tele of the two persons
                 firstPersonName = wsSource.Cells(firstPersonIndex, 2).Value
                 secondPersonName = wsSource.Cells(secondPersonIndex, 2).Value
                 firstPersonTele = wsSource.Cells(firstPersonIndex, 3).Value
                 secondPersonTele = wsSource.Cells(secondPersonIndex, 3).Value
                
                 ' Compare data for each person
                 For dataIndex = 2 To lastCol
                    ' Check val is not empty
                    If Not (IsEmpty(wsSource.Cells(firstPersonIndex, dataIndex).Value) Or IsEmpty(wsSource.Cells(secondPersonIndex, dataIndex).Value)) Then
                        numCellsChecked = numCellsChecked + 1
                         ' Compare data and count if it's the same
                         If wsSource.Cells(firstPersonIndex, dataIndex).Value = wsSource.Cells(secondPersonIndex, dataIndex).Value Then
                             sameDataCount = sameDataCount + 1
                         End If
                    End If
                 Next dataIndex
                
                 ' Write results to the new sheet
                 wsResult.Cells(printIndex, 1).Value = firstPersonName
                 wsResult.Cells(printIndex, 3).Value = secondPersonName
                 wsResult.Cells(printIndex, 2).Value = firstPersonTele
                 wsResult.Cells(printIndex, 4).Value = secondPersonTele
                 wsResult.Cells(printIndex, 5).Value = sameDataCount
                 wsResult.Cells(printIndex, 6).Value = numCellsChecked
                 wsResult.Cells(printIndex, 7).Value = (sameDataCount / numCellsChecked) * 100
                 printIndex = printIndex + 1
            End If
        Next secondPersonIndex
    Next firstPersonIndex
End Sub


