VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmlocalsale 
   BackColor       =   &H80000002&
   Caption         =   "REJECT & CF"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtbf 
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   4080
      Width           =   1935
   End
   Begin VB.TextBox txtfrom 
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox txtto 
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtintake 
      Height          =   375
      Left            =   2040
      TabIndex        =   16
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox txtspillage 
      Height          =   375
      Left            =   2040
      TabIndex        =   14
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   5040
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Left            =   4080
      TabIndex        =   12
      Top             =   1200
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "PostDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Branch"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "SQnty"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "RQnty"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "QntyC/F"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Q/Spillage"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtbranch 
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   600
      Width           =   3735
   End
   Begin VB.ComboBox CboBranch 
      Height          =   315
      Left            =   2040
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtquantity 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdpost 
      Caption         =   "Post"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtRejects 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtcf 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   2040
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   145817601
      CurrentDate     =   41926
   End
   Begin VB.Label Label10 
      Caption         =   "B/F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Transfer Fro Station"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Transfer To Station"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Total Intake"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Spillage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Branch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Post_date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Quantity Sold"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Quantity Rejected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Physical Qnty Cf"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
   End
End
Attribute VB_Name = "frmlocalsale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CboBranch_Change()
Set Rst = oSaccoMaster.GetRecordset("SELECT BName FROM d_Branch where bcode='" & CboBranch & "'")
If Not Rst.EOF Then
txtbranch.Text = Rst.Fields("BName")
Else
txtbranch.Text = ""
End If
If Trim(CboBranch.Text) <> "All" Then
Set rs = oSaccoMaster.GetRecordset("SET DATEFORMAT DMY SELECT SUM(Dispatch)Quantity FROM Dispatch WHERE  Transdate ='" & DTPicker1 & "'")
If Not rs.EOF Then
txtquantity = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
Else
txtquantity = "0"
End If
 sql = "set dateformat dmy SELECT  ISNULL(SUM(RIntake),0) QSupplied FROM  RouteIntake where TransDate='" & DTPicker1 & "'"
    Set Rst = oSaccoMaster.GetRecordset(sql)
    If Not Rst.EOF Then
      txtintake = IIf(IsNull(Rst("QSupplied")), 0, Rst("QSupplied"))
     Else
     txtintake = 0
    End If

Else
Set rs = oSaccoMaster.GetRecordset("SET DATEFORMAT DMY SELECT SUM(Dispatch)Quantity FROM Dispatch WHERE TransDate ='" & DTPicker1 & "'")
If Not rs.EOF Then
txtquantity = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
Else
txtquantity = "0"
End If
sql = "set dateformat dmy SELECT  SUM(RIntake) QSupplied FROM  RouteIntake where TransDate='" & DTPDispatchDate & "' "
    Set Rst = oSaccoMaster.GetRecordset(sql)
    If Not Rst.EOF Then
      txtintake = IIf(IsNull(Rst("QSupplied")), 0, Rst("QSupplied"))
     Else
     txtintake = 0
    End If
End If

txtFrom = getFromstation(Trim(CboBranch), DTPicker1)
txtTo = getTostation(Trim(CboBranch), DTPicker1)
txtbf = getBf(Trim(CboBranch), DTPicker1)

DTPicker1_Change

End Sub

Private Sub CboBranch_Click()
CboBranch_Change
End Sub

Private Sub cmdPost_Click()
Dim Total, Intake As Double
If CboBranch.Text = "" Then
MsgBox "Select the branch to proceed", vbInformation
CboBranch.SetFocus
Exit Sub
End If
If txtcf.Text = "" Then
MsgBox "Enter Quantity C/f", vbInformation
txtcf.SetFocus
Exit Sub
End If

If txtquantity.Text = "" Then
MsgBox "Enter Quantity sold", vbInformation
txtquantity.SetFocus
Exit Sub
End If
If txtRejects.Text = "" Then
MsgBox "Enter Rejected Quantity", vbInformation
txtRejects.SetFocus
Exit Sub
End If
If txtspillage.Text = "" Then
MsgBox "Enter Rejected Quantity", vbInformation
txtspillage.SetFocus
Exit Sub
End If
Total = CDbl(txtquantity) + CDbl(txtcf) + CDbl(txtRejects) + CDbl(txtspillage) + CDbl(txtTo)
Intake = CDbl(txtintake) + CDbl(txtFrom) + CDbl(txtbf)

If Total > Intake Then
'  MsgBox ("Totals cannot Exceed Milk Intake For the Day,Ensure Entries Are Correct")
'  Exit Sub
  
End If

  Set Rst = oSaccoMaster.GetRecordset(" set dateformat dmy select * from milkcontrol2 where transdate='" & DTPicker1 & "'and bcode='" & CboBranch.Text & "' ")
If Rst.EOF Then
 oSaccoMaster.ExecuteThis " set dateformat dmy insert into milkcontrol2 (Intake,SQuantity,Reject,transdate,auditid,cfa,Spillage,FromStation,Tostation,Bf,bcode)values ('" & txtintake & "','" & txtquantity & "','" & txtRejects & "','" & DTPicker1 & "','" & User & "','" & txtcf.Text & "','" & txtspillage & "','" & txtFrom & "','" & txtTo & "','" & txtbf & "','" & CboBranch.Text & "') "
  MsgBox "Saved successfully ", vbInformation
  Else
oSaccoMaster.ExecuteThis " set dateformat dmy update  milkcontrol2  set intake='" & txtintake & "',SQuantity='" & txtquantity & "',Reject= '" & txtRejects & "',spillage= '" & txtspillage & "',Bcode= '" & CboBranch & "',auditid='" & User & "',cfa='" & txtcf & "',FromStation='" & txtFrom & "' ,Tostation='" & txtTo & "',bf='" & txtbf & "'   where transdate='" & DTPicker1 & "'  AND bcode='" & CboBranch.Text & "' "
  MsgBox " updated  successfully ", vbInformation
  End If
  txtquantity.Text = ""
  txtRejects.Text = ""
  txtcf.Text = ""
  txtbranch.Text = ""
  Form_Load
End Sub



Private Sub Command1_Click()
ListView1.Visible = True
End Sub

Private Sub DTPicker1_Change()
ListView1.ListItems.Clear
    sql = "set dateformat dmy select transdate,Squantity,reject,cfa,spillage,bcode from milkcontrol2 where transdate='" & DTPicker1 & "' and bcode='" & CboBranch.Text & "'  "
    Set Rst = oSaccoMaster.GetRecordset(sql)

    While Not Rst.EOF
        Set li = ListView1.ListItems.Add(, , Rst("transdate"))
         li.ListSubItems.Add , , Rst("bcode")
        li.ListSubItems.Add , , Rst("Squantity")
        li.ListSubItems.Add , , Rst("reject")
        li.ListSubItems.Add , , Rst("cfa")
        li.ListSubItems.Add , , Rst("Spillage")
        Rst.MoveNext
    Wend
   ' CboBranch_Change
End Sub



Private Sub Form_Load()
CboBranch.Clear
ListView1.ListItems.Clear
ListView1.Visible = False
 DTPicker1 = Format(Get_Server_Date, "dd/mm/yyyy")
 Set rs = CreateObject("adodb.recordset")
    rs.Open "SELECT bcode FROM d_Branch order by bcode", cn
    If rs.EOF Then Exit Sub
    With rs
        While Not .EOF
         CboBranch.AddItem rs.Fields(0)
         .MoveNext
        Wend
    End With
    sql = "set dateformat dmy select transdate,isnull(Squantity,0),isnull(reject,0),isnull(cfa,0),bcode from milkcontrol2 where transdate='" & DTPicker1 & "' and bcode='" & CboBranch.Text & "'"
    Set Rst = oSaccoMaster.GetRecordset(sql)

    While Not Rst.EOF
        Set li = ListView1.ListItems.Add(, , Rst("transdate"))
        li.ListSubItems.Add , , Rst("bcode")
        li.ListSubItems.Add , , Rst("Squantity")
        li.ListSubItems.Add , , Rst("reject")
        li.ListSubItems.Add , , Rst("cfa")
        li.ListSubItems.Add , , Rst("Spillage")
'        li.ListSubItems.Add , , rst("AgentPrice")
'        li.ListSubItems.Add , , rst("RetailPrice")
'        li.ListSubItems.Add , , rst("accno")
        Rst.MoveNext
    Wend
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
   CboBranch.Text = Item.ListSubItems(1)
   txtquantity.Text = Item.ListSubItems(2)
    txtRejects.Text = Item.ListSubItems(3)
    txtcf.Text = Item.ListSubItems(4)
    
End Sub
Public Function getFromstation(bcode As String, ddate As Date) As Double
 Dim rsgetfrom As New ADODB.Recordset
 Set rsgetfrom = oSaccoMaster.GetRecordset("  set dateformat dmy Select isnull(sum(fromStation),0) as Fromstation from Milktransfer where   toBranch='" & bcode & "' and Transdate= '" & ddate & "'")
    If Not rsgetfrom.EOF Then
     getFromstation = IIf(IsNull(rsgetfrom(0)), 0, rsgetfrom(0))
    Else
     getFromstation = 0
    End If
   
End Function
Public Function getTostation(bcode As String, ddate As Date) As Double
 Dim rsgetto As New ADODB.Recordset
 Set rsgetto = oSaccoMaster.GetRecordset("  set dateformat dmy Select isnull(sum(fromStation),0) as Tostation from Milktransfer where   fromBranch='" & bcode & "' and Transdate= '" & ddate & "'")
    If Not rsgetto.EOF Then
     getTostation = IIf(IsNull(rsgetto(0)), 0, rsgetto(0))
    Else
     getTostation = 0
    End If
   
End Function
Public Function getBf(bcode As String, ddate As Date) As Double
 Dim rsBf As New ADODB.Recordset
 Dim yday As Date
 yday = DateSerial(year(ddate), month(ddate), Day(ddate) - 1)
 Set rsBf = oSaccoMaster.GetRecordset("  set dateformat dmy Select isnull(sum(cfa),0) as Bf from milkcontrol2 where   Bcode='" & bcode & "' and Transdate= '" & yday & "'")
    If Not rsBf.EOF Then
     getBf = IIf(IsNull(rsBf(0)), 0, rsBf(0))
    Else
     getBf = 0
    End If
   
End Function
