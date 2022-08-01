Sub CompositionSetter()
    
    Rem Variable declaration
    Dim SkuStringToSearch As String
    Dim NewCompositionValue As String
    
    SkuStringToSearch = InputBox("Enter SKU string to search:")
    Rem Search string cannot be empty, so we need some error handling here:
    If SkuStringToSearch = "" Then
        MsgBox ("VALUE ERROR: String to search cannot be empty! exiting macro...")
        Exit Sub
    End If
    
    NewCompositionValue = InputBox("Enter new Composition value:")
    
    
    MsgBox ("This macro will replace Composition value of all SKU that contains " & SkuStringToSearch & " with Composition value " & NewCompositionValue)
    

    Rem Variable declaration

    Dim SkuColumn As Integer
    
    Rem Finding SKU column (So we no longer hard coded to assume A)
    Rem However, we are still assuming first row is where we could find the header :(
    Set SkuCell = Rows(1).Find("SKU", LookIn:=xlValues)
    If SkuCell Is Nothing Then
        MsgBox ("INVALID SHEET FORMAT: no SKU column! exiting macro...")
        Exit Sub
    Else
        SkuColumn = SkuCell.Column
    End If
    
    Dim CompositionColumn As Integer
    
    Rem Finding Composition column (So we no longer hard coded to assume B)
    Rem However, we are still assuming first row is where we could find the header :(
    Set CompositionCell = Rows(1).Find("Composition", LookIn:=xlValues)
    If CompositionCell Is Nothing Then
        MsgBox ("INVALID SHEET FORMAT: no Composition column! exiting macro...")
        Exit Sub
    Else
        CompositionColumn = CompositionCell.Column
    End If
    
    Dim LastRowIndex As Integer

    Rem Finding last row index
    Rem Reference explaination: https://stackoverflow.com/a/27066381
    LastRowIndex = Cells(Rows.Count, SkuColumn).End(xlUp).Row
    
    Dim Cell As Range
    Dim Counter As Integer
    
    Rem Loop through all value cells in SKU column
    Rem Reference: https://stackoverflow.com/a/7190567
    Rem Starting from row 2 to skip header
    For Each Cell In Range(Cells(2, SkuColumn), Cells(LastRowIndex, SkuColumn))
        
        Rem This function return position at which match is found, which we don't care, so <> 0 is sufficient
        Rem Reference: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/instr-function
        If InStr(Cell.Value, SkuStringToSearch) <> 0 Then
            Rem Debug for Pig
            Rem MsgBox ("Match at row " & Cell.Row)
            Cells(Cell.Row, CompositionColumn).Value = NewCompositionValue
        End If
        
        Counter = Counter + 1
    Next Cell
    
    Rem Debug for Pig
    Rem MsgBox (Counter)

    
End Sub

