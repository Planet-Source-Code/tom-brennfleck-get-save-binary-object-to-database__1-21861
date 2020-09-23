VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   4650
      TabIndex        =   6
      Top             =   2760
      Width           =   1965
   End
   Begin VB.TextBox txtFileName 
      Height          =   345
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   240
      Width           =   6375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save To DB"
      Height          =   405
      Left            =   4650
      TabIndex        =   3
      Top             =   1860
      Width           =   1995
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get From DB"
      Height          =   405
      Left            =   4650
      TabIndex        =   2
      Top             =   3840
      Width           =   1995
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Picture"
      Height          =   405
      Left            =   4650
      TabIndex        =   1
      Top             =   1290
      Width           =   1995
   End
   Begin VB.CommandButton cmdGetPicture 
      Caption         =   "Get Picture From Disk"
      Height          =   405
      Left            =   4650
      TabIndex        =   0
      Top             =   720
      Width           =   1995
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   30
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      Caption         =   "Label2"
      Height          =   375
      Left            =   30
      TabIndex        =   7
      Top             =   4650
      Width           =   6735
   End
   Begin VB.Image pPicture 
      BorderStyle     =   1  'Fixed Single
      Height          =   3525
      Left            =   240
      Stretch         =   -1  'True
      Top             =   720
      Width           =   4305
   End
   Begin VB.Label Label1 
      Caption         =   "File Name"
      Height          =   225
      Left            =   270
      TabIndex        =   4
      Top             =   0
      Width           =   2235
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************     Form Level Declaration
Private WithEvents HO As cBinaryDBObject
Attribute HO.VB_VarHelpID = -1
Private sFileName As String
Private DocumentDB As Database
'
'
'
Private Sub Command2_Click()

    GetBinaryObject

End Sub

Private Sub Command3_Click()

    SaveBinaryObject

End Sub

'***************     Form Level Procedures
Private Sub Form_Load()

    'open the database
    OpenDB DocumentDB, True
    txtFileName = ""
    LoadListBox

End Sub

Private Sub LoadListBox()
Dim rs As Recordset
Dim SQL As String

    List1.Clear

    ' open the table.
    SQL = "SELECT * FROM tblFileObject"
    Set rs = DocumentDB.OpenRecordset(SQL)
    With rs
        If .EOF And .BOF Then
            .Close
            GoTo LoadListBoxExit
        End If
        .MoveFirst
        Do
            List1.AddItem !ID
            .MoveNext
        Loop Until .EOF
    End With

LoadListBoxExit:

End Sub

Public Sub OpenDB(MyDB As Database, Optional OpenMDB As Boolean = True)

    If OpenMDB Then
        '/* Password protected database file */
        Set MyDB = Workspaces(0).OpenDatabase(App.Path & "\db1.mdb", False, False, "")
    Else
        MyDB.Close
        Set MyDB = Nothing
    End If

End Sub

Private Sub HO_Error(ID As Long, Msg As String)

    MsgBox ID & ":  " & Msg

End Sub

Private Sub HO_Status(ID As Long, Msg As String)

    lblStatus = CStr(ID) & ":  " & Msg

End Sub

Private Sub GetBinaryObject()
Dim FieldNames(1) As Variant           'names of the other fields to return
Dim RD() As Variant                    'store for the returned data, not the binary field
Dim FN As String                       'Binary file name to use as storage
Dim i As Integer

    If List1.SelCount > 0 Then
        Set HO = New cBinaryDBObject       'create the new bd object

        FieldNames(0) = "ID"               'return the ID field
        FieldNames(1) = "FileName"         'return the filename

        With HO
            .KillFile = True                            'kill the filename if it exists
            Set .DB = DocumentDB                        'pass the database
            .ObjectKeyFieldName = "ID"                  'the key/index field is
            .ObjectKey = List1.List(List1.ListIndex)    'the value to search for is
            .ObjectFieldName = "OLEModule"              'name of the field that contains the binary file
            .ObjectTableName = "tblFileObject"          'table that contains the binary files
            .SubFieldNames = FieldNames                 'pass in the field names to return
            .FileName = App.Path & "\picture.bmp"       'file name to use
            .GetObject                                  'get the file from the database
            .ReturnData RD()                            'return any aditional data
            FN = .FileName                              'actual file name used - if default was used
        End With
        Set HO = Nothing

        pPicture.Picture = LoadPicture(FN)

        For i = 0 To UBound(RD)
            Debug.Print RD(i)                      'print aditional info returned
        Next

    End If

End Sub


Private Sub SaveBinaryObject()
Dim FieldNames(1) As Variant           'names of the other fields to return
Dim FieldData(1) As Variant            'names of the other fields to return
Dim RD() As Variant                    'store for the returned data, not the binary field
Dim FN As String                       'Binary file name to use as storage
Dim i As Integer

    If sFileName = "" Then
        Exit Sub
    End If

    Set HO = New cBinaryDBObject       'create the new bd object

    FieldNames(0) = "ID"               'return the ID field
    FieldNames(1) = "FileName"         'return the filename
    FieldData(0) = Null                  'return the ID field
    FieldData(1) = sFileName           'return the filename

    With HO
        .KillFile = False                       'kill the filename if it exists
        Set .DB = DocumentDB                   'pass the database
        .ObjectKeyFieldName = "ID"             'the key/index field is
        .ObjectKey = -1                        'the value to search for is
        .ObjectFieldName = "OLEModule"         'name of the field that contains the binary file
        .ObjectTableName = "tblFileObject"     'table that contains the binary files
        .SubFieldNames = FieldNames            'pass in the field names to return
        .SubFieldData = FieldData
        .FileName = sFileName                  'file name to use
        .SaveObject                            'get the file from the database
        .ReturnData RD()                       'return any aditional data
        FN = .FileName                         'actual file name used - if default was used
    End With
    Set HO = Nothing

    LoadListBox

    For i = 0 To UBound(RD)
        Debug.Print RD(i)                      'print aditional info returned
    Next
End Sub

Private Sub cmdGetPicture_Click()
On Error Resume Next

    With CommonDialog1
         .FileName = ""
         .DialogTitle = "Extract to"
         .CancelError = True
         .Filter = "BMP (*.BMP)|*.bmp|All (*.*)|*.*"
         .FileName = sFileName
         .ShowSave
         If Err.Number > 0 Then
              Exit Sub
         End If
         sFileName = .FileName
    End With
    txtFileName = sFileName
    pPicture.Picture = LoadPicture(sFileName)


End Sub
'
Private Sub Command1_Click()

    pPicture.Picture = LoadPicture("")

End Sub
