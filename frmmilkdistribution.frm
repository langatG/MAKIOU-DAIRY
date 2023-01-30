VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmmilkdistribution 
   BackColor       =   &H00C0C000&
   Caption         =   "Milk Distribution"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   7410
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Height          =   375
      Left            =   1440
      Picture         =   "frmmilkdistribution.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   11
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdpost 
      Caption         =   "Post"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtdispatch 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtintake 
      Height          =   375
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtNames 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   5415
   End
   Begin VB.ComboBox cboDCode 
      Height          =   315
      ItemData        =   "frmmilkdistribution.frx":02C2
      Left            =   120
      List            =   "frmmilkdistribution.frx":02C4
      TabIndex        =   0
      Text            =   "cboDCode"
      Top             =   720
      Width           =   1215
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
      Format          =   148701185
      CurrentDate     =   41926
   End
   Begin VB.Label Label2 
      Caption         =   "Customer"
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
      TabIndex        =   10
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Kilos Dispatched"
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
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
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
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
End
Attribute VB_Name = "frmmilkdistribution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim edit As Boolean
 Dim Tostation, Intake, Total As Double

Private Sub cboDCode_Click()
 cboDCode_Change
End Sub

Private Sub cboDCode_Change()
Set Rst = oSaccoMaster.GetRecordset("select p.dname from d_debtors p  where p.dcode='" & cboDCode & "'")
    If Not Rst.EOF Then
        txtNames.Text = Rst("dname")
    Else
        txtNames.Text = ""
    End If
End Sub
Private Sub cmdPost_Click()
    If txtdispatch.Text = "" Then
      MsgBox "Enter  Quantity To Dispatch", vbInformation
      txtdispatch.SetFocus
      Exit Sub
    End If

  If Trim(cboDCode.Text) = "" Then
     MsgBox "Select the Customer First To Dipatch to", vbInformation
     cboDCode.SetFocus
   Exit Sub
  End If
  
  If txtdispatch = "" Then txtdispatch = 0
  
  Total = CDbl(txtdispatch)
  Intake = CDbl(txtintake)

  Set Rst = oSaccoMaster.GetRecordset(" set dateformat dmy select * from Dispatch where transdate='" & DTPicker1 & "' and Dcode= '" & cboDCode & "' ")
  If Rst.EOF Then
   oSaccoMaster.ExecuteThis " set dateformat dmy insert into Dispatch (Transdate, Dispatch, Dcode,auditid, TIntake) " _
                           & "   values ('" & DTPicker1 & "','" & CDbl(txtdispatch) & "','" & cboDCode.Text & "','" & User & "','" & CDbl(txtintake) & "') "
  MsgBox "Saved successfully ", vbInformation
  Else
  oSaccoMaster.ExecuteThis " set dateformat dmy update  Dispatch  set Tintake='" & CDbl(txtintake) & "',Dispatch='" & CDbl(txtdispatch) & "',auditid='" & User & "' where transdate='" & DTPicker1 & "' and Dcode='" & cboDCode.Text & "' "
  MsgBox " updated  successfully ", vbInformation
  End If
  
  txtdispatch.Text = "0"

  edit = False
End Sub

Private Sub Command1_Click()
  edit = True
End Sub

Private Sub DTPicker1_Change()
 sql = "set dateformat dmy SELECT  ISNULL(SUM(QSupplied),0) QSupplied FROM  d_Milkintake where TransDate='" & DTPicker1 & "'"
    Set Rst = oSaccoMaster.GetRecordset(sql)
    If Not Rst.EOF Then
      txtintake = IIf(IsNull(Rst("QSupplied")), 0, Rst("QSupplied"))
     Else
     txtintake = 0
    End If
    txtintake = CDbl(txtintake)
End Sub

Private Sub Form_Load()
   DTPicker1 = Format(Get_Server_Date, "dd/mm/yyyy")
    Set rs = CreateObject("adodb.recordset")
    DTPicker1_Change
     edit = False
End Sub


Private Sub Picture5_Click()
        frmSearchDebtors.Show vbModal
        cboDCode.Text = sel
        cboDCode_Change
End Sub
    

