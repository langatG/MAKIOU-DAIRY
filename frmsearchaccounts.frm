VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmsearchaccounts 
   Caption         =   "SEARCH ACCOUNT NUMBERS"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   Icon            =   "frmsearchaccounts.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2550
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      Enabled         =   0   'False
      Height          =   375
      Left            =   510
      TabIndex        =   0
      Top             =   3240
      Width           =   1095
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "frmsearchaccounts.frx":08CA
      DataField       =   "accno"
      DataSource      =   "Adodc1"
      Height          =   2595
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4577
      _Version        =   393216
      ListField       =   "accno"
      BoundColumn     =   "accno"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3960
      Top             =   3240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "Investar"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmsearchaccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    On Error GoTo 10
    strName = DataList1.BoundText
    Unload Me
    Exit Sub
10:    MsgBox Err.description
End Sub

Private Sub Command2_Click()
    strName = ""
    Unload Me
End Sub

Private Sub DataList1_Click()
    On Error Resume Next
    Command1.Enabled = True
End Sub

Private Sub DataList1_DblClick()
On Error Resume Next
    Call Command1_Click
End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)
On Error Resume Next
'
If KeyAscii = 13 Then
End If
End Sub

Private Sub Form_Load()
    On Error GoTo 10
    Dim myclass As Object
    Dim strQ
    Dim cn As Connection
    Set myclass = New cdbase
    Set cn = CreateObject("adodb.connection")
    Provider = myclass.OpenCon
    cn.Open Provider, "bi"
    Adodc1.ConnectionString = cn
    With Adodc1
        .RecordSource = "select distinct ACCNo from CUb order by ACCno "
        .Refresh
    End With
    Exit Sub
10:    MsgBox Err.description
End Sub






