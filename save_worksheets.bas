Sub SplitWorkbook()
'Updateby20200806
Dim FileExtStr As String
Dim FileFormatNum As Long
Dim xWs As Worksheet
Dim xWb As Workbook
Dim xNWb As Workbook
Dim WbName As String
Dim FolderName As String
Dim WbNameLength As Integer
Dim WbPath As String
Application.ScreenUpdating = False
Set xWb = Application.ThisWorkbook

WbPath = xWb.Path
WbName = xWb.Name
WbNameLength = Len(WbName) - 4
FolderName = WbPath & "\" & Left(WbName, WbNameLength)

If Val(Application.Version) < 12 Then
    FileExtStr = ".xls": FileFormatNum = -4143
Else
    Select Case xWb.FileFormat
        Case 51:
            FileExtStr = ".xlsx": FileFormatNum = 51
        Case 52:
            If Application.ActiveWorkbook.HasVBProject Then
                FileExtStr = ".xlsm": FileFormatNum = 52
            Else
                FileExtStr = ".xlsx": FileFormatNum = 51
            End If
        Case 56:
            FileExtStr = ".xls": FileFormatNum = 56
        Case Else:
            FileExtStr = ".xls": FileFormatNum = 56
        End Select
End If

If Dir(FolderName, vbDirectory) = vbNullString Then
    MkDir FolderName
End If

For Each xWs In xWb.Worksheets
On Error GoTo NErro
    If xWs.Visible = xlSheetVisible Then
    xWs.Select
    xWs.Copy
    xFile = FolderName & "\" & xWs.Name & FileExtStr
    Set xNWb = Application.Workbooks.Item(Application.Workbooks.Count)
    xNWb.SaveAs xFile, FileFormat:=FileFormatNum
    xNWb.Close False, xFile
    End If
NErro:
    xWb.Activate
Next

    MsgBox "You can find the files in " & FolderName
    Application.ScreenUpdating = True
End Sub
