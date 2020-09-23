<div align="center">

## A Basic Set of File handling controls \(updated\)


</div>

### Description

FileReal, CloseAllFiles, CopyFile, DeleteFile, GetAttrib, GetFileDate, GetFileExtension, GetFileSize, MakeDIR, RemoveDIR, SetHidden, SetReadOnly, SetSystem, SetNormal, Overwrite
 
### More Info
 
Filename, Path, Source, Destination

Filesize, File attributes, File Date/Time, File extension


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sam Truscott](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sam-truscott.md)
**Level**          |Beginner
**User Rating**    |4.6 (97 globes from 21 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sam-truscott-a-basic-set-of-file-handling-controls-updated__1-21513/archive/master.zip)





### Source Code

```
'-------FileSys V1.0-------
'----by Samuel Truscott----
'----www.pezcore.co.uk-----
Public Sub Save(filename as string)
if filereal = true then
 if msgbox("Overwrite File?", vbYesNo) = vbYes then
  deletefile(filename)
  'save file code
else
  'do NOT overwrite the file
end if
end if
End Sub
Public Function FileReal(Filename) As Boolean
On Error goto Error
If Dir(Filename) = Filename Then
FileReal = True
Else
FileReal = False
End If
Exit Function
Error:
Exit Sub
End Function
Public Function GetFileSize(FileName) As String
On Error GoTo Gfserror
Dim TempStr As String
TempStr = FileLen(FileName)
If TempStr >= "1024" Then
'KB
TempStr = CCur(TempStr / 1024) & "KB"
 Else
 If TempStr >= "1048576" Then
 'MB
 TempStr = CCur(TempStr / (1024 * 1024)) & "KB"
 Else
 TempStr = CCur(TempStr) & "B"
 End If
End If
GetFileSize = TempStr
Exit Function
Gfserror:
GetFileSize = "0B"
Resume
End Function
Public Function GetAttrib(FileName) As String
On Error GoTo GAError
Dim TempStr As String
TempStr = GetAttr(FileName)
If TempStr = "64" Then
TempStr = "Alias"
End If
If TempStr = "32" Then
TempStr = "Archive"
End If
If TempStr = "16" Then
TempStr = "Directory"
End If
If TempStr = "2" Then
TempStr = "Hidden"
End If
If TempStr = "0" Then
TempStr = "Normal"
End If
If TempStr = "1" Then
TempStr = "ReadOnly"
End If
If TempStr = "4" Then
TempStr = "System"
End If
If TempStr = "8" Then
TempStr = "Volume"
End If
GetAttrib = TempStr
Exit Function
GAError:
GetAttrib = "Unknown"
Resume
End Function
Public Sub SetHidden(FileName As String)
On Error Resume Next
SetAttr FileName, vbHidden
End Sub
Public Sub SetReadOnly(FileName As String)
On Error Resume Next
SetAttr FileName, vbReadOnly
End Sub
Public Sub SetSystem(FileName As String)
On Error Resume Next
SetAttr FileName, vbSystem
End Sub
Public Sub SetNormal(FileName As String)
On Error Resume Next
SetAttr FileName, vbNormal
End Sub
Public Function GetFileExtension(FileName As String)
On Error Resume Next
Dim TempStr As String
TempStr = Right(FileName, 2)
If Left(TempStr, 1) = "." Then
GetFileExtension = Right(FileName, 1)
Exit Function
Else
 TempStr = Right(FileName, 3)
 If Left(TempStr, 1) = "." Then
 GetFileExtension = Right(FileName, 2)
 Exit Function
 Else
 TempStr = Right(FileName, 4)
 If Left(TempStr, 1) = "." Then
 GetFileExtension = Right(FileName, 3)
 Exit Function
 Else
 TempStr = Right(FileName, 5)
 If Left(TempStr, 1) = "." Then
 GetFileExtension = Right(FileName, 4)
 Exit Function
 Else
 GetFileExtension = "Unknown"
 End If
 End If
 End If
End If
End Function
Public Function GetFileDate(FileName As String) As String
On Error Resume Next
GetFileDate = FileDateTime(FileName)
End Function
Public Sub DeleteFile(FileName As String)
On Error GoTo DelError
Kill FileName
Exit Sub
DelError:
MsgBox "Error deleting File"
Resume
End Sub
Public Sub CopyFile(Source As String, Destination As String)
On Error GoTo CopyError
FileCopy Source, Destination
Exit Sub
CopyError:
MsgBox "Error copying File"
Resume
End Sub
Public Sub MoveFile(Source As String, Destination As String)
On Error GoTo MoveError
FileCopy Source, Destination
Kill Source
Exit Sub
MoveError:
MsgBox "Error moving File"
Resume
End Sub
Public Sub MakeDIR(Path As String)
On Error GoTo DIRError
MkDir Path
Exit Sub
DIRError:
MsgBox "Error creating Directory"
Resume
End Sub
Public Sub RemoveDIR(Path As String)
On Error GoTo DIRError2
RmDir Path
Exit Sub
DIRError2:
MsgBox "Error removing Directory"
Resume
End Sub
Public Sub CloseAllFiles()
On Error Resume Next
Reset
End Sub
```

