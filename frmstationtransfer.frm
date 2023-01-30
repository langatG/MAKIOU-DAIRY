VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmstationtransfer 
   Caption         =   " "
   ClientHeight    =   4875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7050
   ScaleHeight     =   4875
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdpost 
      Caption         =   "Post"
      Height          =   495
      Left            =   3000
      TabIndex        =   18
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Height          =   495
      Left            =   1440
      TabIndex        =   17
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transfer/Recieve"
      Height          =   615
      Left            =   2040
      TabIndex        =   13
      Top             =   1680
      Width           =   3255
      Begin VB.OptionButton optfrom 
         Caption         =   "From station"
         Height          =   195
         Left            =   1320
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optto 
         Caption         =   "To station"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.ComboBox cbotransferto 
      Height          =   315
      Left            =   1920
      TabIndex        =   11
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txttransferto 
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox txtfrom 
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtto 
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   3240
      Width           =   1935
   End
   Begin VB.ComboBox CboBranch 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Text            =   "`"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtbranch 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.TextBox txtintake 
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   3720
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   112197633
      CurrentDate     =   41926
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
      TabIndex        =   16
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "To Branch"
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
      TabIndex        =   12
      Top             =   1080
      Width           =   1335
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
      TabIndex        =   8
      Top             =   3240
      Width           =   1815
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
      TabIndex        =   6
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "From Branch"
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
      Top             =   480
      Width           =   1335
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
      TabIndex        =   4
      Top             =   3720
      Width           =   1815
   End
End
Attribute VB_Name = "frmstationtransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim edit As Boolean
 Dim Tostation, Intake, Total As Double

Private Sub cbobranch_Change()
Set Rst = oSaccoMaster.GetRecordset("SELECT BName FROM d_Branch where bcode='" & CboBranch & "'")
If Not Rst.EOF Then
txtbranch.Text = Rst.Fields("BName")
Else
txtbranch.Text = ""
End If


 sql = "set dateformat dmy SELECT  BranchCode, ISNULL(SUM(QSupplied),0) QSupplied FROM  d_Milkintake where TransDate='" & DTPicker1 & "' AND  branchcode='" & CboBranch & "' GROUP BY branchcode"
    Set Rst = oSaccoMaster.GetRecordset(sql)
    If Not Rst.EOF Then
      txtintake = IIf(IsNull(Rst("QSupplied")), 0, Rst("QSupplied"))
     Else
     txtintake = 0
    End If
    Tostation = getTostation(Trim(CboBranch), Format(DTPicker1, "dd/mm/yyyy"))
    txtintake = CDbl(txtintake) - Tostation
End Sub
Public Function getTostation(bcode As String, ddate As Date) As Double
 Dim rsgetto As New ADODB.Recordset
 Set rsgetto = oSaccoMaster.GetRecordset("  set dateformat dmy Select isnull(sum(Tostation),0) as Tostation from Milktransfer where   ToBranch='" & bcode & "' and Transdate= '" & ddate & "'")
    If Not rsgetto.EOF Then
     getTostation = IIf(IsNull(rsgetto(0)), 0, rsgetto(0))
    Else
     getTostation = 0
    End If
   
End Function

Private Sub cbobranch_Click()
 cbobranch_Change
End Sub

Private Sub cbotransferto_Change()
Set Rst = oSaccoMaster.GetRecordset("SELECT BName FROM d_Branch where bcode='" & cbotransferto & "'")
If Not Rst.EOF Then
txttransferto.Text = Rst.Fields("BName")
Else
txttransferto.Text = ""
End If
End Sub

Private Sub cbotransferto_Click()
 cbotransferto_Change
End Sub

Private Sub cmdpost_Click()
  If optto.value = False And optfrom = False Then
     MsgBox "Select Transfer Option Either from or To Station", vbInformation
    Exit Sub
  End If
  
  If optfrom.value = True Then
    If txtfrom.Text = "" Then
      MsgBox "Enter Rejected Quantity", vbInformation
      txtfrom.SetFocus
      Exit Sub
    End If
  End If
  If optto.value = True Then
    If txtto.Text = "" Then
        MsgBox "Enter Rejected Quantity", vbInformation
        txtto.SetFocus
        Exit Sub
    End If
  End If
  If Trim(CboBranch.Text) = Trim(cbotransferto.Text) Then
     MsgBox "Can't Transfer from to the same branch", vbInformation
   Exit Sub
  End If
  If txtto = "" Then txtto = 0
  If txtfrom = "" Then txtfrom = 0
  
  Total = CDbl(txtfrom)
  Intake = CDbl(txtintake)

If Total > Intake Then
   If MsgBox("Totals cannot Exceed Milk Intake For the Day,Are Sure the Entries Are Correct?", vbQuestion + vbYesNo) = vbNo Then
     Exit Sub
   End If
End If

  'Set Rst = oSaccoMaster.GetRecordset(" set dateformat dmy select * from Milktransfer where transdate='" & DTPicker1 & "'and ToBranch='" & cbotransferto.Text & "' ")
  If Not edit Then
   oSaccoMaster.ExecuteThis " set dateformat dmy insert into Milktransfer (Transdate, fromStation, Tostation, FromBranch, ToBranch, auditid, Intake) " _
                           & "   values ('" & DTPicker1 & "','" & txtfrom & "','" & txtto & "','" & CboBranch.Text & "','" & cbotransferto.Text & "','" & user & "','" & txtintake & "') "
  MsgBox "Saved successfully ", vbInformation
  Else
  oSaccoMaster.ExecuteThis " set dateformat dmy update  Milktransfer  set intake='" & txtintake & "',fromStation='" & txtfrom & "',Tostation= '" & txtto & "',FromBranch= '" & CboBranch & "',ToBranch= '" & cbotransferto & "',auditid='" & user & "'    where transdate='" & DTPicker1 & "'  AND ToBranch='" & cbotransferto.Text & "' and FromBranch='" & CboBranch.Text & "' "
  MsgBox " updated  successfully ", vbInformation
  End If
  
  
  
  txtfrom.Text = "0"
  txtto.Text = "0"
  txtintake = 0
  edit = False
End Sub

Private Sub Command1_Click()
  edit = True
End Sub

Private Sub Form_Load()
   DTPicker1 = Format(Get_Server_Date, "dd/mm/yyyy")
    Set rs = CreateObject("adodb.recordset")
     edit = False
     
    CboBranch.Clear
    cbotransferto.Clear
    rs.Open "SELECT bcode FROM d_Branch order by bcode", cn
    If rs.EOF Then Exit Sub
    With rs
        While Not .EOF
         CboBranch.AddItem rs.Fields(0)
         .MoveNext
        Wend
    End With
     Set rs = CreateObject("adodb.recordset")
    rs.Open "SELECT bcode FROM d_Branch order by bcode", cn
    If rs.EOF Then Exit Sub
    With rs
        While Not .EOF
         cbotransferto.AddItem rs.Fields(0)
         .MoveNext
        Wend
    End With
End Sub

Private Sub optfrom_Click()
   If optfrom.value = True Then
    txtto.Visible = False
    Label8.Visible = False
    txtfrom.Visible = True
    Label9.Visible = True
 End If
End Sub

Private Sub optto_Click()
 If optto.value = True Then
  txtto.Visible = True
  Label8.Visible = True
  txtfrom.Visible = False
  Label9.Visible = False
  
 End If
End Sub
