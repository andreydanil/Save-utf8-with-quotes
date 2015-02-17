Attribute VB_Name = "CSV"
'
' @package      vba
' @author       andrey danilkovich
' @copyright    mit (c) 2015
' @since        version 1.0
' @description  type checking to add quotes around cells
'

Sub main()
    ' init file name and save path
    Dim fName As String
    Dim fPath As String
    
    On Error GoTo 0
    
    ' select all cells
    ActiveWorkbook.Activate
    With ActiveWorkbook.WebOptions
        .Encoding = msoEncodingUTF8
    End With
    
    Cells.Select
    
    fName = "abc.txt"       ' ADD NAME HERE
    fPath = "Desktop\"      ' ADD PATH HERE
    
    QuoteCommaExport fName, fPath

End Sub
Sub QuoteCommaExport(fName, fPath)
   ' Dimension all variables.
   Dim DestFile As String
   Dim FileNum As Integer
   Dim ColumnCount As Integer
   Dim RowCount As Integer
    
   ' Prompt user for destination file name.
   DestFile = fPath & fName

   ' Obtain next free file handle number.
   FileNum = FreeFile()

   ' Turn error checking off.
   On Error Resume Next

   ' Attempt to open destination file for output.
   Open DestFile For Output As #FileNum

   ' Exception handling: if an error occurs report it and end.
   If Err <> 0 Then
      MsgBox "Cannot open fName " & DestFile
      End
   End If
    
   ' Turn error checking on.
   On Error GoTo 0
   
   ' Select active cells.
   ActiveCell.CurrentRegion.Select
   
   ' Loop for each row in selection.
   For RowCount = 1 To Selection.Rows.Count

      ' Loop for each column in selection.
      For ColumnCount = 1 To Selection.Columns.Count

         ' Write current cell's text to file with quotation marks.
         Print #FileNum, """" & Selection.Cells(RowCount, _
            ColumnCount).Text & """";

         ' Check if cell is in last column.
         If ColumnCount = Selection.Columns.Count Then
            ' If so, then write a blank line.
            Print #FileNum,
         Else
            ' Otherwise, write a comma.
            Print #FileNum, ",";
         End If
      ' Start next iteration of ColumnCount loop.
      Next ColumnCount
   ' Start next iteration of RowCount loop.
   Next RowCount

   ' Close destination file.
   Close #FileNum
End Sub

'
' @description        create file if file doesn't exist
'
Function CreateFile(fName As String, contents As String)
    Dim tempFile As String
    Dim nextFileNum As Long
    
    nextFileNum = FreeFile
    tempFile = fName
     
    Open tempFile For Output As #nextFileNum
    Print #nextFileNum, contents
    Close #nextFileNum
 
End Function

'
' @description      add quotes around text
'
Sub QuotesAroundText()
    Dim c As Range
        For Each c In Selection
        If Not IsNumeric(c.Value) Then
            c.Value = """" & c.Value & """"
        End If
    Next c
End Sub

