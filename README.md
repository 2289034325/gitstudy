# gitstudy
study

=HYPERLINK(
    LEFT(A1, FIND("[", A1)-1) & 
    MID(A1, FIND("[", A1)+1, FIND("]", A1)-FIND("[", A1)-1) & 
    "#" & 
    MID(A1, FIND("]", A1)+1, FIND("!", A1)-FIND("]", A1)-1) & 
    "!" & 
    SUBSTITUTE(RIGHT(A1, LEN(A1)-FIND("!", A1)), "$", ""),
    "Open File & Jump"
)
Sub OpenAndFocusFromRef()
    ' This sub opens the workbook and focuses on a specific cell.
    ' Example reference: ='C:\Users\folder1\folder2\[boo1.xlsm]somesheet'!$B$3
    Dim sRef As String
    sRef = "='C:\Users\folder1\folder2\[boo1.xlsm]somesheet'!$B$3"  ' Your dynamic reference string

    Dim sStr As String
    sStr = sRef
    
    ' Remove the initial "=" and leading single quote if present
    If Left(sStr, 1) = "=" Then sStr = Mid(sStr, 2)
    If Left(sStr, 1) = "'" Then sStr = Mid(sStr, 2)
    
    ' Now sStr should look like: C:\Users\folder1\folder2\[boo1.xlsm]somesheet'!$B$3
    Dim posQuote As Long
    posQuote = InStr(sStr, "'")
    Dim mainRef As String
    mainRef = Left(sStr, posQuote - 1)  ' Gets: C:\Users\folder1\folder2\[boo1.xlsm]somesheet
    
    Dim posEx As Long
    posEx = InStr(sStr, "!")
    Dim sCell As String
    sCell = Mid(sStr, posEx + 1)  ' Gets the cell address, e.g., $B$3
    
    ' Split mainRef into folder path, file name, and sheet name
    Dim pos1 As Long, pos2 As Long
    pos1 = InStr(mainRef, "[")
    If pos1 = 0 Then
        MsgBox "Oops! The reference format seems off."
        Exit Sub
    End If
    
    Dim sPath As String, sFile As String, sSheet As String
    sPath = Left(mainRef, pos1 - 1)  ' Folder path: C:\Users\folder1\folder2\
    
    pos2 = InStr(mainRef, "]")
    If pos2 = 0 Then
        MsgBox "Hmm, couldn't find the file name part."
        Exit Sub
    End If
    sFile = Mid(mainRef, pos1 + 1, pos2 - pos1 - 1)  ' File name: boo1.xlsm
    sSheet = Mid(mainRef, pos2 + 1)  ' Sheet name: somesheet
    
    ' Debug prints (optional, for checking)
    Debug.Print "Folder: " & sPath
    Debug.Print "File: " & sFile
    Debug.Print "Sheet: " & sSheet
    Debug.Print "Cell: " & sCell
    
    ' Open the workbook and focus on the specific cell
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks.Open(sPath & sFile)
    On Error GoTo 0
    If wb Is Nothing Then
        MsgBox "Couldn't open the workbook: " & sPath & sFile
        Exit Sub
    End If
    
    wb.Worksheets(sSheet).Activate
    wb.Worksheets(sSheet).Range(sCell).Select
End Sub
