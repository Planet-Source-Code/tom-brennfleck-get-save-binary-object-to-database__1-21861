<div align="center">

## Get/Save Binary Object To Database


</div>

### Description

Save binary objects to a DAO database, eg: pictures, exe files, dll's etz. Its a generic class module that allows saving/extracting from any access database. Also gets additional fileds if required for example a persons name if it was stored in the same table
 
### More Info
 
KillFile- Kill the file if its present

ObjectKeyFieldName- key field name to the database

ObjectTableName- table name holding the bin object

ObjectKey - binary object key to extract

SubFieldData - other field data to extract/save

SubFieldNames- other field names to extract/save

ObjectFieldName- field holding the binary object

DB- database holding the binary object table

BlockSize- block size to use

FileName- filename to extract to or import from

FileName - returns file name if it was changed

eg: temp file was used

ReturnData - returns a variant array of the aditional database fields that were requested

none known


<span>             |<span>
---                |---
**Submitted On**   |2001-03-23 18:19:44
**By**             |[Tom Brennfleck](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tom-brennfleck.md)
**Level**          |Intermediate
**User Rating**    |5.0 (55 globes from 11 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD174483232001\.zip](https://github.com/Planet-Source-Code/tom-brennfleck-get-save-binary-object-to-database__1-21861/archive/master.zip)

### API Declarations

```
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
```





