VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTransportersDetails 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Transporters Details"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   7605
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF80&
      Caption         =   "Contacts"
      Height          =   1695
      Left            =   120
      TabIndex        =   24
      Top             =   2280
      Width           =   7455
      Begin VB.TextBox txtTown 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   32
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   26
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox txtPAddress 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "E - Mail"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Phone"
         Height          =   195
         Left            =   2880
         TabIndex        =   30
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Postal Address"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Town"
         Height          =   195
         Left            =   2880
         TabIndex        =   28
         Top             =   960
         Width           =   405
      End
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1560
      TabIndex        =   23
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2880
      TabIndex        =   22
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   21
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   480
      TabIndex        =   20
      Top             =   6960
      Width           =   735
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF80&
      Caption         =   "Other Details"
      Height          =   1695
      Left            =   120
      TabIndex        =   17
      Top             =   5280
      Width           =   7455
      Begin VB.CheckBox chkperkg 
         Caption         =   "Kgs Delivered"
         Height          =   255
         Left            =   2400
         TabIndex        =   53
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtrate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         TabIndex        =   52
         Text            =   "0"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         TabIndex        =   49
         Text            =   "0"
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox chkTrip 
         Caption         =   "Trip"
         Height          =   255
         Left            =   2520
         TabIndex        =   48
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtflatratet 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         TabIndex        =   39
         Text            =   "0"
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chkflatrate 
         Caption         =   "Flat Rate"
         Height          =   255
         Left            =   2520
         TabIndex        =   38
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtSubsidy 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "Per Kg"
         Height          =   255
         Left            =   6240
         TabIndex        =   55
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblkg 
         Caption         =   "Rate"
         Height          =   255
         Left            =   3840
         TabIndex        =   54
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label23 
         Caption         =   "Per  Kg"
         Height          =   255
         Left            =   6240
         TabIndex        =   51
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbltrip 
         Caption         =   "Rate"
         Height          =   255
         Left            =   3840
         TabIndex        =   50
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Rate"
         Height          =   255
         Left            =   3840
         TabIndex        =   41
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "Per Day"
         Height          =   255
         Left            =   6240
         TabIndex        =   40
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Subsidy (Per Kg)"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1170
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Personal Details"
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7455
      Begin VB.PictureBox Picture5 
         Height          =   255
         Left            =   960
         Picture         =   "frmTransportersDetails.frx":0000
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   37
         Top             =   480
         Width           =   255
      End
      Begin VB.ComboBox cboBranch 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmTransportersDetails.frx":02C2
         Left            =   240
         List            =   "frmTransportersDetails.frx":02C4
         TabIndex        =   35
         Top             =   1680
         Width           =   2895
      End
      Begin VB.CheckBox chkActive 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Active"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   33
         Top             =   1680
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.ComboBox cboLocation 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   16
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtId 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         TabIndex        =   13
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtNames 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   12
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtTCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPRegDate 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   108855297
         CurrentDate     =   40096
      End
      Begin VB.Label Label22 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   1320
         TabIndex        =   47
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label21 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   2640
         TabIndex        =   46
         Top             =   840
         Width           =   135
      End
      Begin VB.Label Label20 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   6840
         TabIndex        =   45
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label19 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   720
         TabIndex        =   44
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label18 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   1800
         TabIndex        =   43
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label17 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   135
         Left            =   960
         TabIndex        =   42
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Branch"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   1440
         Width           =   510
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Location"
         Height          =   195
         Left            =   1920
         TabIndex        =   15
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Id Number/Business No"
         Height          =   195
         Left            =   5040
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Date registered"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Names"
         Height          =   195
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Trans Code"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Bank Details"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Width           =   7455
      Begin VB.ComboBox cboBBranch 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         TabIndex        =   34
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cboBName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtAccNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         TabIndex        =   4
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Account Number"
         Height          =   195
         Left            =   4920
         TabIndex        =   3
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Bank Name"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "Bank Branch"
         Height          =   195
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmTransportersDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flat, Trip As Integer, PerKgs As Integer
Private Sub chkflatrate_Click()
If chkflatrate = vbChecked Then
Flat = 1
PerKgs = 0
Trip = 0
chkTrip.Value = vbUnchecked
chkperkg.Value = vbUnchecked
txtamount = 0
txtrate = 0
Else
chkflatrate.Value = vbUnchecked
Flat = 0
PerKgs = 0
Trip = 0
End If
End Sub

Private Sub chkperkg_Click()
If chkperkg.Value = vbChecked Then
Flat = 0
Trip = 0
PerKgs = 1
txtamount = 0
txtflatratet = 0
chkTrip = vbUnchecked
chkflatrate = vbUnchecked
Else
chkperkg.Value = vbUnchecked
Flat = 0
Trip = 0
PerKgs = 0
End If

End Sub

Private Sub chkTrip_Click()
If chkTrip = vbChecked Then
Flat = 0
Trip = 1
PerKgs = 0
txtrate = 0
txtflatratet = 0
chkflatrate.Value = vbUnchecked
chkperkg.Value = vbUnchecked
Else
chkTrip.Value = vbUnchecked
Flat = 0
Trip = 0
PerKgs = 0
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdedit_Click()
txtAccNo.Locked = False
txtEmail.Locked = False
txtId.Locked = False
txtNames.Locked = False
txtPAddress.Locked = False
txtPhone.Locked = False
txtSubsidy.Locked = False
txtTCode.Locked = False
txtTown.Locked = False
cboBBranch.Locked = False
cboBName.Locked = False
cboLocation.Locked = False

'cmdEdit.Enabled = False
'cmdSave.Enabled = False
cmdSave.Enabled = True
End Sub

Private Sub cmdnew_Click()
txtAccNo = ""
txtEmail = ""
txtId = ""
txtNames = ""
txtPAddress = ""
txtPhone = ""
txtSubsidy = ""
txtTCode = ""
txtTown = ""
cboBBranch.Text = ""
cboBName.Text = ""
cboLocation.Text = ""
cboBranch.Text = ""

txtAccNo.Locked = False
txtEmail.Locked = False
txtId.Locked = False
txtNames.Locked = False
txtPAddress.Locked = False
txtPhone.Locked = False
txtSubsidy.Locked = False
txtTCode.Locked = False
txtTown.Locked = False
cboBBranch.Locked = False
cboBName.Locked = False
cboLocation.Locked = False
'cmdEdit.Enabled = False
'cmdSave.Enabled = False
cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
Dim Active As String
Dim rate As Currency, TripAmount As Currency
On Error GoTo ErrorHandler

If txtTCode = "" Then
MsgBox "Please enter the transporters code ", vbInformation, "Missing Information"
txtTCode.SetFocus
Exit Sub
End If

If txtSubsidy = "" Then
txtSubsidy = "0"
End If

If chkActive.Value = vbChecked Then
    Active = "1"
Else
    Active = "0"
End If
If chkflatrate = vbChecked Then
Flat = 1
If txtflatratet = "" Then
MsgBox "The rate per day is not privided for yet you have select the use of Flat Rate", vbInformation, "EASYMA"
Exit Sub
End If
rate = txtflatratet
Else
If txtflatratet = "" Then txtflatratet = 0
rate = txtflatratet
Flat = Flat
End If
If cboBranch = "" Then
MsgBox "Please select the branch before you proceed", vbInformation, "EASYMA"
Exit Sub
End If
TripAmount = 0
'"""""""""""" PAID BY COMPANY PER KG
If chkTrip = vbChecked Then
    If txtamount = "" Then
        MsgBox "The rate per Trip is not provided for yet you have select the use of Trip Rate", vbInformation, "EASYMA"
        txtamount.SetFocus
        Exit Sub
    End If
  TripAmount = txtamount
End If

If chkperkg = vbChecked Then
    If txtrate = "" Then
        MsgBox "The rate per Kg is not provided for yet you have select the use of Per Kg Rate", vbInformation, "EASYMA"
        txtrate.SetFocus
        Exit Sub
    End If
 TripAmount = txtrate
End If

Set cn = New ADODB.Connection
sql = "d_sp_Transporter '" & txtTCode & "','" & txtNames & "','" & txtId & "','" & cboLocation & "','" & DTPRegDate & "','" & txtEmail & "','" & txtPhone & "','" & txtTown & "','" & txtPAddress & "'," & txtSubsidy & ",'" & txtAccNo & "','" & cboBName & "'," & Active & ",'" & Replace(cboBBranch, "'", "") & "','" & cboBranch & "','" & User & "'," & Flat & "," & Trip & "," & rate & ""
oSaccoMaster.ExecuteThis (sql)

sql = ""
sql = "UPDATE  d_Transporters SET   Trip=" & Trip & ",PerKg=" & PerKgs & ",Amount=" & TripAmount & ",Active=" & Active & ",status =0, status2 =0, status3 =0, status4 =0, status5 =0, status6 =0 where TransCode ='" & txtTCode & "'"
oSaccoMaster.ExecuteThis (sql)
cmdnew_Click
cmdSave.Enabled = False

MsgBox "Records successively updated."
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub Form_Load()
Dim myclass As cdbase

txtAccNo.Locked = True
txtEmail.Locked = True
txtId.Locked = True
txtNames.Locked = True
txtPAddress.Locked = True
txtPhone.Locked = True
txtSubsidy.Locked = True
'txtTCode.Locked = True
txtTown.Locked = True
cboBBranch.Locked = True
cboBName.Locked = True
cboLocation.Locked = True
cmdEdit.Enabled = False
cmdSave.Enabled = False

    
    Set myclass = New cdbase

    Provider = myclass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
    cn.Open Provider, "bi"

Set rs = CreateObject("adodb.recordset")
    
    rs.Open "SELECT LName FROM d_Location", cn
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         If Not IsNull(rs.Fields("LName")) Then
         cboLocation.AddItem rs.Fields("LName")
         End If
         .MoveNext
        
        Wend
    
    End With
    
    
    Set rs = CreateObject("adodb.recordset")
    
    rs.Open "SELECT BankName,BranchName FROM d_BANKS", cn
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         If Not IsNull(rs.Fields(0)) Then cboBName.AddItem rs.Fields(0)
         If Not IsNull(rs.Fields(1)) Then cboBBranch.AddItem rs.Fields(1)
         
         .MoveNext
        
        Wend
    
    End With
    
     Set rs = CreateObject("adodb.recordset")
    
    rs.Open "SELECT BName FROM d_Branch", cn
    
    If rs.EOF Then Exit Sub
    
    With rs
        
        While Not .EOF
         
         If Not IsNull(rs.Fields(0)) Then cboBranch.AddItem rs.Fields(0)
         
         .MoveNext
        
        Wend
    
    End With

Flat = 0
DTPRegDate = Format(Get_Server_Date, "dd/mm/yyyy")
End Sub

Private Sub Picture5_Click()
Me.MousePointer = vbHourglass
         frmSearchTransporter.Show vbModal
        txtTCode = sel
        txtTCode_Validate True
        Me.MousePointer = 0
End Sub

Private Sub txtTCode_Validate(Cancel As Boolean)
Dim a As Boolean
Set rs = New ADODB.Recordset
sql = "d_sp_SelectTrans '" & txtTCode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtNames = rs.Fields(0)
If Not IsNull(rs.Fields(1)) Then txtId = rs.Fields(1)
If Not IsNull(rs.Fields(2)) Then cboLocation = rs.Fields(2)
'If Not IsNull(rs.Fields(3)) Then DTPRegDate = rs.Fields(3)
If Not IsNull(rs.Fields(4)) Then txtEmail = rs.Fields(4)
If Not IsNull(rs.Fields(5)) Then txtPhone = rs.Fields(5)
If Not IsNull(rs.Fields(6)) Then txtTown = rs.Fields(6)
If Not IsNull(rs.Fields(7)) Then txtPAddress = rs.Fields(7)
If Not IsNull(rs.Fields(8)) Then txtSubsidy = rs.Fields(8)
If Not IsNull(rs.Fields(9)) Then txtAccNo = rs.Fields(9)
If Not IsNull(rs.Fields(10)) Then cboBName = rs.Fields(10)
If Not IsNull(rs.Fields(11)) Then cboBBranch = rs.Fields(11)
If Not IsNull(rs.Fields(17)) Then Trip = rs.Fields(17)
If Not IsNull(rs.Fields(18)) Then PerKgs = rs.Fields(18)
If Not IsNull(rs.Fields(16)) Then Flat = rs.Fields(16)
If Trip Then
chkTrip = vbChecked
chkTrip_Click
txtamount = IIf(IsNull(rs.Fields(15)), 0, rs.Fields(15))
End If
If PerKgs Then
chkperkg = vbChecked
chkperkg_Click
txtrate = IIf(IsNull(rs.Fields(15)), 0, rs.Fields(15))
End If
If Flat Then
chkflatrate = vbChecked
chkflatrate_Click
txtflatratet = IIf(IsNull(rs.Fields(14)), 0, rs.Fields(14))
End If
If PerKgs = False And Trip = False Then
Flat = True
chkflatrate = vbChecked
End If
'If Not IsNull(rs.Fields(12)) Then a = rs.Fields(12)
If Not IsNull(rs.Fields(13)) Then cboBranch = rs.Fields(13)
If a = True Then
chkActive = vbChecked
Else
chkActive = vbUnchecked
End If
cmdEdit.Enabled = True
cmdSave.Enabled = True
End If
End Sub
