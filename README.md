<div align="center">

## ADO Upload and Retrieve ORACLE BLOB


</div>

### Description

To upload and retrieve ORACLE BLOB File using ADO Recordset.

Never try with other than ORACLE database, but i think i can ...
 
### More Info
 
File Name

BLOB / LongRaw Field Name

ADO Blob RecordSet

I get this function somewhere from microsoft.com

Then the function is rewritten to make sure that it is easy to use and reuse...

Boolean


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Herman Ibrahim](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/herman-ibrahim.md)
**Level**          |Intermediate
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/herman-ibrahim-ado-upload-and-retrieve-oracle-blob__1-6319/archive/master.zip)





### Source Code

```
Function AddLongRaw(ByVal strFileName As String, ByRef objRecSet As ADODB.Recordset, ByVal strFieldName As String) As Boolean
 'How to call AddLongRaw function :
 'dim bool as boolean
 'dim objRecSet as new adodb.recordset
 'dim strFieldeName as string
 'strFieldName = objRecSet.Fields("YOUR_BLOB_FILE").Name
 'bool = AddLongRaw(strSourceName, objRecSet, strFieldName)
 'if bool then
  'Successfully upload the BLOB file into database
 'else
  'Failed to upload the BLOB file into database
 'End If
 AddLongRaw = False
 Dim ByteData() As Byte 'Byte array for Blob data.
 Dim SourceFile As Integer
 Dim FileLength As Long
 Dim Numblocks As Integer
 Dim LeftOver As Long
 Dim i As Integer
 Const BlockSize = 10000 'This size can be experimented with for
 SourceFile = FreeFile
 Open strFileName For Binary Access Read As SourceFile
 FileLength = LOF(SourceFile)  ' Get the length of the file.
 'Debug.Print "Filelength is " & FileLength
 If FileLength = 0 Then
  Close SourceFile
  AddLongRaw = False
  Exit Function
 Else
  Numblocks = FileLength / BlockSize
  LeftOver = FileLength Mod BlockSize
  ReDim ByteData(LeftOver)
  Get SourceFile, , ByteData()
  objRecSet.Fields(strFieldName).AppendChunk ByteData()
  ReDim ByteData(BlockSize)
   For i = 1 To Numblocks
   Get SourceFile, , ByteData()
   objRecSet.Fields(strFieldName).AppendChunk ByteData()
   Next i
   AddLongRaw = True
   Close SourceFile
 End If
End Function
Function GetLongRaw(strFileName As String, objRecSet As ADODB.Recordset, strBLOBFieldName As String) As Boolean
 GetLongRaw = False
 Dim ByteData() As Byte 'Byte array for file.
 Dim DestFileNum As Integer
 Dim DiskFile As String
 Dim FileLength As Long
 Dim Numblocks As Integer
 Const BlockSize = 10000
 Dim LeftOver As Long
 Dim i As Integer
 FileLength = objRecSet.Fields(strBLOBFieldName).ActualSize
 ' Remove any existing destination file.
 DiskFile = strFileName
 If Len(Dir$(DiskFile)) > 0 Then
  Kill DiskFile
 End If
 DestFileNum = FreeFile
 Open DiskFile For Binary As DestFileNum
 Numblocks = FileLength / BlockSize
 LeftOver = FileLength Mod BlockSize
 ByteData() = objRecSet.Fields(strBLOBFieldName).GetChunk(LeftOver)
 Put DestFileNum, , ByteData()
 For i = 1 To Numblocks
  ByteData() = objRecSet.Fields(strBLOBFieldName).GetChunk(BlockSize)
  Put DestFileNum, , ByteData()
 Next i
 Close DestFileNum
 GetLongRaw = True
'============
'The object file is now located at strFileName
'============
End Function
```

