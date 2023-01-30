VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmstations 
   Caption         =   "Create Station"
   ClientHeight    =   3045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvwstation 
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "StationNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "StationName"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox Txtno 
      Height          =   405
      Left            =   960
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtstation 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "StationNo"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmstations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdnew_Click()
  lvwstation.Visible = False
  newstation
End Sub
Sub newstation()
   sql = "select top 1 * from stations order by stationno desc"
   Set rs = oSaccoMaster.GetRecordset(sql)
   If Not rs.EOF Then
    Txtno = Format(rs!Stationno, "00") + 1
   Else
   Txtno = "01"
   End If
End Sub

Private Sub cmdsave_Click()
   If Txtno = "" Then
    MsgBox "Enter station No", vbInformation
    Txtno.SetFocus
    Exit Sub
   End If
    If txtstation = "" Then
        MsgBox "Enter station No", vbInformation
        txtstation.SetFocus
    Exit Sub
   End If

   sql = "insert into Stations(StationNo,StationName) Values(" & Txtno & ",'" & txtstation & "')"
   oSaccoMaster.ExecuteThis (sql)
    MsgBox "Station Saved Successfully", vbInformation
    Form_Load
    Exit Sub
End Sub

Private Sub Form_Load()
  sql = "select * from Stations"
  Set Rst = oSaccoMaster.GetRecordset(sql)
  With Rst
    While Not .EOF
        Set li = lvwstation.ListItems.Add(, , !Stationno & "")
           li.SubItems(1) = !StationName & ""
     .MoveNext
    Wend
  End With
End Sub
