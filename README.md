<div align="center">

## Compact & Repair Database \- Enhanced


</div>

### Description

Easily Compact & Repair a MS Access Database and display the size differences.
 
### More Info
 
strDatabase as string

This code assumes that your Access database is in the same directory as your exe.

Be sure to reference:

MS DAO 3.X Object Library

MS Scripting Runtime

Boolean (True if Successful)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Cierra Computers & Consulting](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/cierra-computers-consulting.md)
**Level**          |Unknown
**User Rating**    |3.0 (6 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/cierra-computers-consulting-compact-repair-database-enhanced__1-3601/archive/master.zip)





### Source Code

```
Public Function CompactDatabase(strDatabaseName As String) As Boolean
On Error GoTo Err_CompactDatabase
Dim strPath As String
Dim strPath1 As String
Dim strPathSize As String
Dim strPathSize2 As String
Screen.MousePointer = vbHourglass
'Save Paths for Database
strPath = App.Path & "\" & strDatabaseName
strPath1 = App.Path & "\" & "BackupOf" & strDatabaseName
'Repair Database
DBEngine.RepairDatabase strPath
'Get Size of File Before Compacting
strPathSize = GetFileSize(strPath)
'Kill the file if it exists
If Dir(strPath1) <> "" Then Kill strPath1
'Compact Database to New Name
DBEngine.CompactDatabase strPath, strPath1
''Kill the file if it exists
If Dir(strPath) <> "" Then Kill strPath
'Compact back to original Name
DBEngine.CompactDatabase strPath1, strPath
'Kill the file, no need to save it
If Dir(strPath1) <> "" Then Kill strPath1
'Get Size of File After Compacting
strPathSize2 = GetFileSize(strPath)
CompactDatabase = True
'Display the Summary
MsgBox UCase(strDatabaseName) & " compacted successfully." _
 & vbNewLine & vbNewLine & "Size before compacting:" & vbTab & strPathSize _
 & vbNewLine & "Size after compacting:" & vbTab & strPathSize2, vbInformation, "Compact Successful"
Err_CompactDatabase:
 Select Case Err
 Case 0
 Case Else
 MsgBox Err & ": " & Error, vbCritical, "CompactDatabase Error"
 End Select
Screen.MousePointer = vbNormal
End Function
Public Function GetFileSize(strFile As String) As String
Dim fso As New Scripting.FileSystemObject
Dim f As File
Dim lngBytes As Long
Const KB As Long = 1024
Const MB As Long = 1024 * KB
Const GB As Long = 1024 * MB
Set f = fso.GetFile(fso.GetFile(strFile))
lngBytes = f.Size
If lngBytes < KB Then
 GetFileSize = Format(lngBytes) & " bytes"
ElseIf lngBytes < MB Then
 GetFileSize = Format(lngBytes / KB, "0.00") & " KB"
ElseIf lngBytes < GB Then
 GetFileSize = Format(lngBytes / MB, "0.00") & " MB"
Else
 GetFileSize = Format(lngBytes / GB, "0.00") & " GB"
End If
End Function
```

