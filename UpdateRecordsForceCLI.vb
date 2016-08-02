Attribute VB_Name = "Module2"
Sub UpdateSFRecordsByCLI(colSize As Integer, worksheetName As String, sheetRange As String, updStatusColIndex As Integer, cmdColIndex As Integer, pathToForceCli As String, sObjectAPIName As String)
    Dim rng As Range
    Dim i As Integer, j As Integer, k As Integer
    Dim ColHeaders() As String
    ReDim ColHeaders(colSize)
        
    'Select range with specified worksheet and range of cells
    Set rng = Worksheets(worksheetName).Range(sheetRange)
    
    'Loop through each row of the selected range
    For i = 1 To rng.Rows.Count
    
        'Set the initial force.com cli record update command using the provided sObject api name
        Dim sCmd As String
        sCmd = pathToForceCli & " record update " & sObjectAPIName & " "
        
        'If this is the first (header) row, loop through each column and store column header cell values which should be field api names
        If i = 1 Then
        
            For j = 1 To rng.Columns.Count
                ColHeaders(j) = rng.Cells(i, j).Value
            Next j
                    
        'Otherwise proceed to loop through each column of the data rows
        Else:
                        
            For k = 1 To rng.Columns.Count
                'If this is the first column (which should be the Id column) append the record Id value to the command string
                If k = 1 Then
                
                    sCmd = sCmd & rng.Cells(i, k).Value & " "
                
                'Otherwise store each field update value with the combination of <field api name>:<field value>
                ElseIf k <> 1 And IsEmpty(rng.Cells(i, k).Value) = False Then
                    Dim cellVal As String
                    Dim pos As Integer
                    
                    cellVal = rng.Cells(i, k).Value
                    pos = InStr(cellVal, " ")
                
                    'If boolean value convert true and false values to a format accepted by force.com cli
                    If rng.Cells(i, k).Value = True Then 'check if the value is a boolean true
                        sCmd = sCmd & ColHeaders(k) & ":" & Chr(34) & "true" & Chr(34) & " "
                    ElseIf rng.Cells(i, k).Value = False Then 'check if the cell value is a boolean false
                        sCmd = sCmd & ColHeaders(k) & ":" & Chr(34) & "false" & Chr(34) & " "
                    'If date value convert to YYYY-MM-DD format before adding to command string
                    ElseIf IsDate(rng.Cells(i, k).Value) Then
                        sCmd = sCmd & ColHeaders(k) & ":" & Year(rng.Cells(i, k).Value) & "-" & Month(rng.Cells(i, k).Value) & "-" & Day(rng.Cells(i, k).Value) & " "
                    'If string value has contains white spaces, enclose value in double quotes
                    ElseIf pos <> 0 Then
                        sCmd = sCmd & ColHeaders(k) & ":" & Chr(34) & cellVal & Chr(34) & " "
                    'Otherwise store append field name and field value combo to the command string
                    Else:
                        sCmd = sCmd & ColHeaders(k) & ":" & rng.Cells(i, k).Value & " "
                    End If
                End If
            Next k
            
            Dim oShell As Object
            Set oShell = CreateObject("WScript.Shell")
            
            Dim oExec As Object
            Dim oOutput As Object
            
            'store generated command string to the command column cell
            rng.Cells(i, cmdColIndex).Value = sCmd
            
            'execute force.com cli command
            Set oExec = oShell.Exec(sCmd)
            
            'Store stdout and stderr stream
            Set oOutput = oExec.Stdout
            Set oError = oExec.StdErr
            
            Dim sLine As String, errLine As String
            
            'Loop through stdout stream and store value to the Update Status column for current row
            While Not oOutput.AtEndOfStream
                sLine = oOutput.ReadLine
                Dim statusCurrVal As String
                statusCurrVal = rng.Cells(i, updStatusColIndex).Value
                statusCurrVal = statusCurrVal & " | " & sLine
                rng.Cells(i, updStatusColIndex).Value = statusCurrVal
            Wend
            
            'Loop through any error returned by force.com cli in stderr stream and store in the udpate status column for current row
            While Not oError.AtEndOfStream
                errLine = oError.ReadLine
                Dim statCurrVal As String
                statCurrVal = statCurrVal & " | " & errLine
                rng.Cells(i, updStatusColIndex).Value = statCurrVal
            Wend
        
        End If
    
    Next i


End Sub


