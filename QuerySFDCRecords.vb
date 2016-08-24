Attribute VB_Name = "Module3"
Sub GetRecords(queryStr As String, wkSheetName As String, pathToForceCli As String)

    Dim sCmd As String
    sCmd = pathToForceCli & " query " & Chr(34) & queryStr & Chr(34)
    
    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")
            
    Dim oExec As Object
    Dim oOutput As Object
            
    'execute force.com cli command
    Set oExec = oShell.Exec(sCmd)
            
    'Store stdout and stderr stream
    Set oOutput = oExec.Stdout
    Set oError = oExec.StdErr
    
    Dim lineCnt As Integer, rowCnt As Integer
    lineCnt = 1
    rowCnt = 1
  
    
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        Dim colCnt As Integer
        'MsgBox sLine
        'MsgBox lineCnt
        If lineCnt <> 2 Then
            Dim splitVals() As String
            splitVals = Split(sLine, "|")
            'MsgBox splitVals(1) & " " & splitVals(2)
            'MsgBox splitVals(0)
            For colCnt = 0 To UBound(splitVals)
               'MsgBox splitVals(colCnt)
               Worksheets(wkSheetName).Cells(rowCnt, colCnt + 1).Value = splitVals(colCnt)
            Next colCnt
            'iterate rowCnt to the next row
            rowCnt = rowCnt + 1
        End If
        
        lineCnt = lineCnt + 1
    Wend

End Sub
