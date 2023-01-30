VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmsharestransactions 
   Caption         =   "Shares Transactions"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14190
   Icon            =   "frmsharestransactions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   14190
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdPrint 
      Caption         =   "Print HShare Statement"
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   8400
      TabIndex        =   15
      Top             =   7320
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPregdate 
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   1320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Format          =   145358849
      CurrentDate     =   40637
   End
   Begin VB.TextBox txtsno 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   960
      TabIndex        =   9
      Top             =   0
      Width           =   2415
   End
   Begin MSComctlLib.ListView lvwshares 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   9340
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Transaction Date"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Amount"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Balance"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Posted By"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Label txtbal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9360
      TabIndex        =   14
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Balance"
      Height          =   375
      Left            =   8520
      TabIndex        =   13
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Name"
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Reg Date"
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label txtidno 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8760
      TabIndex        =   8
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "S No."
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label txtxmemberno 
      Caption         =   "Member No."
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label txtlocation 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Location"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Label txtname 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   6495
   End
   Begin VB.Label Label1 
      Caption         =   "ID No."
      Height          =   255
      Left            =   8040
      TabIndex        =   1
      Top             =   1440
      Width           =   735
   End
End
Attribute VB_Name = "frmsharestransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub CmdPrint_Click()
  reportname = "Sharestatement.rpt"
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub Form_Load()
'pick items from deduction from the
End Sub

Private Sub txtsno_Change()
Dim bal As Double

sql = "SET dateformat dmy SELECT     s.sno, s.Names, s.Type, s.Location, d.mno,s.Type, d.transdate, d.Cash,d.bal"
sql = sql & " From d_Shares d INNER JOIN d_suppliers s on s.sno=d.code WHERE  d.code = '" & txtsno & "'" 'Period = '" & Enddate & "' AND

Set rs2 = oSaccoMaster.GetRecordset(sql)
If rs2.RecordCount > 0 Then
txtidno = IIf(IsNull(rs2.Fields(0)), 0, rs2.Fields(0))
txtname = IIf(IsNull(rs2.Fields(1)), 0, rs2.Fields(1))
'cboSex = rs2.Fields(2)
txtlocation = IIf(IsNull(rs2.Fields(3)), 0, rs2.Fields(3))
Label6 = IIf(IsNull(rs2.Fields(4)), 0, rs2.Fields(4))
txtbal = IIf(IsNull(rs2.Fields(8)), 0, rs2.Fields(8))
DTPregdate = IIf(IsNull(rs2.Fields(6)), Date, rs2.Fields(6))
'optCash.value = rs2.Fields(6).value
'DTPicker2 = Enddate
End If
Dim tamount As Double
bal = 0
'//populate the items on the listview
'Set rs = oSaccoMaster.GetRecordset("SELECT     id, transdate, amount, bal, transdescription, auditid  FROM  d_sconribution where idno='" & txtidno & "'")
lvwshares.ListItems.Clear
Set rs = oSaccoMaster.GetRecordset("SELECT     code, transdate, amnt, bal, type, auditid  FROM  d_shares where code='" & txtsno & "' ORDER BY TRANSDATE")
tamount = 0
With rs
While Not rs.EOF
  
   
   Set li = lvwshares.ListItems.Add(, , IIf(IsNull(!code), 1, !code))
   If rs.Fields("transdate") <> "" Then li.ListSubItems.Add , , rs.Fields("transdate")
   If rs.Fields("Amnt") <> "" Then li.ListSubItems.Add , , rs.Fields("Amnt")
     bal = bal + rs.Fields("Amnt")
     li.ListSubItems.Add , , bal
 '  If rs.Fields("transdescription") <> "" Then li.ListSubItems.Add , , rs.Fields("transdescription")
 '  If rs.Fields("type") <> "" Then li.ListSubItems.Add , , rs.Fields("type")
   If rs.Fields("type") <> "" Then li.ListSubItems.Add , , IIf(IsNull(!Type), "blank", !Type)
   If rs.Fields("auditid") <> "" Then li.ListSubItems.Add , , rs.Fields("auditid")
   tamount = tamount + rs.Fields("Amnt")
   .MoveNext

Wend
End With
If tamount = 0 Then
txtbal = txtbal
Else
txtbal = tamount
End If
lvwshares.View = lvwReport
End Sub
