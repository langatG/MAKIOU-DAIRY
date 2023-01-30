VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmtriptranscomplain 
   BackColor       =   &H00FFFF80&
   Caption         =   "Trip Transporters Complain Desk"
   ClientHeight    =   2145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   1440
      Picture         =   "frmtriptranscomplain.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   13
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox txtNames 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   360
      Width           =   3615
   End
   Begin VB.TextBox txtTCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   1680
      Width           =   855
   End
   Begin VB.CheckBox chkActive 
      Caption         =   "Active"
      Enabled         =   0   'False
      Height          =   195
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   915
   End
   Begin VB.TextBox txtSubsidy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CheckBox chkTrip 
      Caption         =   "Trip"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtkgs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtrate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtamount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CheckBox chkperkg 
      Caption         =   "Kgs Delivered"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtptransdate 
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   145817601
      CurrentDate     =   42135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Names"
      Height          =   195
      Left            =   1200
      TabIndex        =   19
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Trans Code"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   825
   End
   Begin VB.Label Label18 
      Caption         =   "Kgs Delivered"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lbltrip 
      Caption         =   "Rate"
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Subsidy (Per Kg)"
      Height          =   195
      Left            =   5520
      TabIndex        =   15
      Top             =   840
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "Amount"
      Height          =   255
      Left            =   5520
      TabIndex        =   14
      Top             =   1320
      Width           =   615
   End
End
Attribute VB_Name = "frmtriptranscomplain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, Trip As Boolean, PerKg As Boolean
Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmdnew_Click()
Form_Load
txtTCode = ""
txtamount = "0"
txtNames = ""
End Sub

Private Sub cmdSave_Click()
If txtTCode = "" Then
    MsgBox "Please Enter Valid Transporter code ", vbInformation
 Exit Sub
End If
'If chkActive = vbUnchecked Then
'   MsgBox ("The Transporter is inActive,Update his/her Details first")
'   Exit Sub
'End If
If chkTrip = vbUnchecked And chkperkg = vbUnchecked Then
   MsgBox ("The Transporter is Not Paid by the company/Per Kg,Update his/her Details first")
   Exit Sub
End If

If txtamount = 0 Or txtamount = "" Then
 MsgBox "Update Transporters Details Trip Rate first", vbInformation
 txtamount.SetFocus
 Exit Sub
End If

Enddate = DateSerial(Year(dtptransdate), month(dtptransdate) + 1, 1 - 1)
Set rs2 = oSaccoMaster.GetRecordset("set dateformat dmy SELECT Amount  FROM d_TripTransDetailed where Trans_Code='" & txtTCode & "' and transdate='" & dtptransdate & "'")
 If Not rs2.EOF Then
     If MsgBox("Transporter  " & txtNames & " had Delivered milk today. Add this?", vbYesNo + vbQuestion, "Delivery REPEAT") = vbNo Then
       Exit Sub
      Else
     End If
 End If
 
 sql = "set dateformat dmy  set dateformat dmy insert into  d_TripTransDetailed(QNTY,Amount, Subsidy, Trans_Code,Transdate ,EndPeriod,auditid, auditdatetime, BR)" _
      & " values(" & txtkgs & "," & txtamount & ",'" & txtSubsidy & "','" & txtTCode & "','" & dtptransdate & "','" & Enddate & "','" & user & "','" & Get_Server_Date & "','0')"
 oSaccoMaster.ExecuteThis (sql)
 MsgBox "Transporter Delivery Saved Successfully"
 cmdnew_Click
 
End Sub

Private Sub Form_Load()
 dtptransdate = Format(Get_Server_Date, "dd/mm/yyyy")
End Sub


Private Sub Picture5_Click()
frmSearchTransporter.Show vbModal
        txtTCode = sel
        txtTCode_Validate True
       
End Sub

Private Sub txtkgs_Change()
If chkperkg = vbChecked Then
 If txtkgs <> "" Then
 If IsNull(txtSubsidy) Then txtSubsidy = 0
 txtamount = Format((CDbl(txtrate) * CDbl(txtkgs)), Cfmt)
 End If
End If
End Sub

Private Sub txtkgs_Click()
txtkgs_Change
End Sub

Private Sub txtkgs_KeyPress(KeyAscii As Integer)
If (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii = 45) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "Please enter a number "
End If
End Sub

Private Sub txtTCode_Validate(Cancel As Boolean)
PerKg = False
Trip = False
a = False
Set rs = New ADODB.Recordset
sql = "d_sp_SelectTrans '" & txtTCode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then txtNames = rs.Fields(0)
    If Not IsNull(rs.Fields(8)) Then txtSubsidy = rs.Fields(8)
    If Not IsNull(rs.Fields(12)) Then a = rs.Fields(12)
    If Not IsNull(rs.Fields(15)) Then Trip = rs.Fields(17)
    If Not IsNull(rs.Fields(15)) Then PerKg = rs.Fields(18)
    If Not IsNull(rs.Fields(14)) Then txtrate = rs.Fields(15)
    If a = True Then
     chkActive = vbChecked
    Else
     chkActive = vbUnchecked
    End If
    If Trip = True Then
     chkTrip = vbChecked
     chkperkg = vbUnchecked
     txtamount = CDbl(txtrate)
     txtkgs = CDbl(txtrate)
     txtkgs.Locked = True
    Else
     chkTrip = vbUnchecked
    End If
    
    If PerKg = True Then
     chkTrip = vbUnchecked
     chkperkg = vbChecked
     txtkgs.Locked = fase
     txtkgs = 0
    Else
     chkperkg = vbUnchecked
    End If
    'cmdEdit.Enabled = True
    cmdsave.Enabled = True
End If
End Sub


