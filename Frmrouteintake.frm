VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frmrouteintake 
   BackColor       =   &H00FF00FF&
   Caption         =   "Route Intake Summary"
   ClientHeight    =   3045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtintake 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtroute 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   480
      Width           =   3735
   End
   Begin VB.ComboBox Cboroute 
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      Text            =   "`"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtrouteintake 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdpost 
      Caption         =   "Post"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   2400
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   161677313
      CurrentDate     =   41926
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
      TabIndex        =   10
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Route Name"
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
      TabIndex        =   9
      Top             =   480
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
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Route Summary Intake"
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
      TabIndex        =   7
      Top             =   1320
      Width           =   1935
   End
End
Attribute VB_Name = "Frmrouteintake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim edit As Boolean
 Dim Tostation, Intake, Total As Double

Private Sub Cboroute_Change()
Set Rst = oSaccoMaster.GetRecordset("SELECT RouteName FROM D_Routes where RCode='" & Cboroute & "'")
If Not Rst.EOF Then
txtroute.Text = Rst.Fields("RouteName")
Else
txtroute.Text = ""
End If


 sql = "set dateformat dmy SELECT  AuditId, ISNULL(SUM(QSupplied),0) QSupplied FROM  d_Milkintake where TransDate='" & DTPicker1 & "' AND  AuditId='" & txtroute & "' GROUP BY AuditId"
    Set Rst = oSaccoMaster.GetRecordset(sql)
    If Not Rst.EOF Then
      txtintake = IIf(IsNull(Rst("QSupplied")), 0, Rst("QSupplied"))
     Else
     txtintake = 0
    End If
    txtintake = CDbl(txtintake)
End Sub

Private Sub Cboroute_Click()
 Cboroute_Change
End Sub

Private Sub cmdPost_Click()

    If txtrouteintake.Text = "" Then
      MsgBox "Enter Rejected Quantity", vbInformation
      txtrouteintake.SetFocus
      Exit Sub
    End If

  If Trim(Cboroute.Text) = "" Then
     MsgBox "Select the Route First", vbInformation
     Cboroute.SetFocus
   Exit Sub
  End If
  
  If txtrouteintake = "" Then txtrouteintake = 0
  
  Total = CDbl(txtrouteintake)
  Intake = CDbl(txtintake)

  Set Rst = oSaccoMaster.GetRecordset(" set dateformat dmy select * from RouteIntake where transdate='" & DTPicker1 & "' and Route= '" & Cboroute & "' ")
  If Rst.EOF Then
   oSaccoMaster.ExecuteThis " set dateformat dmy insert into RouteIntake (Transdate, RIntake, Route,auditid, TIntake) " _
                           & "   values ('" & DTPicker1 & "','" & CDbl(txtrouteintake) & "','" & Cboroute.Text & "','" & User & "','" & CDbl(txtintake) & "') "
  MsgBox "Saved successfully ", vbInformation
  Else
  oSaccoMaster.ExecuteThis " set dateformat dmy update  RouteIntake  set Tintake='" & CDbl(txtintake) & "',RIntake='" & CDbl(txtrouteintake) & "',Route= '" & Cboroute & "',auditid='" & User & "'    where transdate='" & DTPicker1 & "' and Route='" & Cboroute.Text & "' "
  MsgBox " updated  successfully ", vbInformation
  End If
  
  txtrouteintake.Text = "0"

  edit = False
End Sub

Private Sub Command1_Click()
  edit = True
End Sub

Private Sub DTPicker1_Change()
  Cboroute_Change
End Sub

Private Sub Form_Load()
   DTPicker1 = Format(Get_Server_Date, "dd/mm/yyyy")
    Set rs = CreateObject("adodb.recordset")
     edit = False
     
    Cboroute.Clear
    rs.Open "SELECT RCode FROM D_Routes order by RCode", cn
    If rs.EOF Then Exit Sub
    With rs
        While Not .EOF
         Cboroute.AddItem rs.Fields(0)
         .MoveNext
        Wend
    End With
End Sub
