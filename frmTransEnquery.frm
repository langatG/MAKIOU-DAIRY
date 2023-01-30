VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTransEnquery 
   Caption         =   "Transporter's Enquery"
   ClientHeight    =   9600
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   14670
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3360
      TabIndex        =   16
      Top             =   0
      Width           =   4575
   End
   Begin VB.TextBox txtBank 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      TabIndex        =   15
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtTelNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9240
      TabIndex        =   14
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox txtBox 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3360
      TabIndex        =   13
      Top             =   360
      Width           =   4575
   End
   Begin VB.TextBox txtTransport 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9240
      TabIndex        =   12
      Top             =   0
      Width           =   3015
   End
   Begin VB.TextBox txtTCode 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox txtBBranch 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4080
      TabIndex        =   11
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtAccNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9240
      TabIndex        =   10
      Top             =   840
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   7455
      Begin VB.CommandButton cmdShow 
         Caption         =   "Show"
         Default         =   -1  'True
         Height          =   375
         Left            =   6000
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   146472961
         CurrentDate     =   40157
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   375
         Left            =   4080
         TabIndex        =   7
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   146472961
         CurrentDate     =   40157
      End
      Begin VB.Label Label9 
         Caption         =   "Date From"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Date To"
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   2160
      Picture         =   "frmTransEnquery.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox TXTIDNO 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin MSComctlLib.ListView lvwEnguery 
      Height          =   6615
      Left            =   0
      TabIndex        =   17
      Top             =   2280
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   11668
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description/SNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "DR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Balance"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblTKgs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   34
      Top             =   9120
      Width           =   2415
   End
   Begin VB.Label Label13 
      Caption         =   "Total Kgs"
      Height          =   375
      Left            =   240
      TabIndex        =   33
      Top             =   9120
      Width           =   1335
   End
   Begin VB.Label Label16 
      Caption         =   "Gross"
      Height          =   255
      Left            =   4440
      TabIndex        =   32
      Top             =   9120
      Width           =   735
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   31
      Top             =   9120
      Width           =   1935
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8640
      TabIndex        =   30
      Top             =   9120
      Width           =   2055
   End
   Begin VB.Label Label19 
      Caption         =   "Deductions"
      Height          =   255
      Left            =   7680
      TabIndex        =   29
      Top             =   9120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Bank :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   28
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   27
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Box :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   26
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Telephone :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   25
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Transport :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   24
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "Branch :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   22
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Account Number :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7320
      TabIndex        =   21
      Top             =   840
      Width           =   1800
   End
   Begin VB.Label Label11 
      Caption         =   "Loc :"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblNPay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   615
      Left            =   8160
      TabIndex        =   19
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label Label12 
      Caption         =   "ID:"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   1080
      Width           =   495
   End
End
Attribute VB_Name = "frmTransEnquery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Flat As Boolean, PerKg As Boolean, Trip As Boolean
Private Sub cmdshow_Click()
txtTCode_Validate True
End Sub

Private Sub Form_Load()
dtpFrom = Format(Get_Server_Date, "dd/mm/yyyy")
dtpFrom = DateSerial(year(dtpFrom), month(dtpFrom), 1)
dtpTo = DateSerial(year(dtpFrom), month(dtpFrom) + 1, 1 - 1)
WindowState = vbMaximized
End Sub

Private Sub LoadData()
Dim I As Integer, TotalDed As Double
Dim bal As Double, Total As Double, Kgs As Double
bal = 0
I = 0
Total = 0
TotalDed = 0
Kgs = 0
lvwEnguery.ListItems.Clear
'd_sp_UpdateTripTranstmpEnquery
If Flat Then
oSaccoMaster.ExecuteThis ("d_sp_UpdateTranstmpEnquery '" & txtTCode & "','" & dtpTo & "'")
Else
oSaccoMaster.ExecuteThis ("d_sp_UpdateTripTranstmpEnquery '" & txtTCode & "','" & dtpTo & "'")
End If
oSaccoMaster.ExecuteThis ("d_sp_UpdateTranstmpEnqueryDed '" & txtTCode & "','" & dtpFrom & "','" & dtpTo & "'")

'Set rs = oSaccoMaster.GetRecordset("SELECT TransDate, SNo, CR, DR, Bal FROM d_tmpTransEnquery WHERE Code ='" & txtTCode & "' ORDER BY TransDate")
Set rs = oSaccoMaster.GetRecordset("SELECT TransDate, SNo, CR, DR, Bal FROM d_tmpTransEnquery WHERE Code ='" & txtTCode & "' ORDER BY CASE WHEN SNO LIKE '%[^0-9]%' THEN 9E99 ELSE CAST(SNO AS INTEGER) END, SNO")
'CASE IsNumeric(sno) WHEN 1 THEN Replicate(Char(0), 100 - Len(sno)) + sno ELSE sno END
'case when isnumeric(your_column) = 1 then your_column else 999999999 end,
'your_colum
'CASE WHEN col LIKE '%[^0-9]%' THEN 9E99 ELSE CAST(col AS INTEGER) END, col
With rs
While Not rs.EOF
   '//check if it is a flat rate case
   Dim Rst As New ADODB.Recordset, rate As Double, samson As Integer, Amount As Double, L3 As Double
  
   Set Rst = oSaccoMaster.GetRecordset("SELECT     *  FROM         d_Transporters   WHERE     (transcode = '" & txtTCode & "') ")

   If Not Rst.EOF Then
   Set li = lvwEnguery.ListItems.Add(, , IIf(IsNull(!transdate), "", !transdate))
   If Flat Then
   li.SubItems(1) = IIf(IsNull(!sno), "", !sno)
   End If
   If Trip Then
   li.SubItems(1) = "Per Trip"
   End If
   If PerKg Then
   li.SubItems(1) = "Per Kgs"
   End If
   li.SubItems(2) = IIf(IsNull(!Cr), 0, !Cr)
   li.SubItems(3) = IIf(IsNull(!Dr), 0, !Dr)
   bal = bal + li.SubItems(2) - li.SubItems(3)
   li.SubItems(4) = bal
   Else
    Set li = lvwEnguery.ListItems.Add(, , dtpFrom)
    rate = Rst.Fields("rate")
    If I = 0 Then
    li.SubItems(1) = "Flat Rate"
     End If
   samson = Days_In_Month(month(dtpTo), month(dtpTo))
   Amount = samson * rate
   If I = 0 Then
   li.SubItems(2) = Amount
   End If
   li.SubItems(3) = IIf(IsNull(!Dr), 0, !Dr)
   L3 = li.SubItems(3)
   If I > 1 Then
   If L3 > 0 Then
   li.SubItems(1) = IIf(IsNull(!sno), "", !sno)
   bal = bal - li.SubItems(3)
   End If
   Else
   If I = 0 Then
   bal = bal + li.SubItems(2) - li.SubItems(3)
   End If
   End If
   li.SubItems(4) = bal
   End If
   
   Total = Total + li.SubItems(2)
   TotalDed = TotalDed + li.SubItems(3)
   I = I + 1
   .MoveNext
 
Wend
End With
lblNPay = "Net Pay :" & Format(bal, "#,##0.00")
Label17 = Format(Total, "#,##0.00")
Label18 = Format(TotalDed, "#,##0.00")

    Set rs = oSaccoMaster.GetRecordset("d_sp_TransTotal '" & txtTCode & "'," & month(dtpFrom) & ", " & year(dtpFrom))
     If Not rs.EOF Then
       Kgs = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
     End If
lblTKgs = Format(Kgs, "#,##0.00")
End Sub

Private Sub Picture5_Click()
Me.MousePointer = vbHourglass
         frmSearchTransporter.Show vbModal
        txtTCode = sel
        txtTCode_Validate True
        Me.MousePointer = 0
End Sub

Private Sub txtTCode_Validate(Cancel As Boolean)

Set rs = New ADODB.Recordset
sql = "d_sp_TransEnquiry  '" & txtTCode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
    ' SELECT     TransName, CertNo, Locations, Phoneno, Address + ' ' + Town AS Expr1, Bcode, BBranch, Accno
    'From dbo.d_Transporters
    If Not IsNull(rs.Fields(0)) Then txtName = rs.Fields(0)
    If Not IsNull(rs.Fields(1)) Then txtidno = rs.Fields(1)
    If Not IsNull(rs.Fields(2)) Then txtlocation = rs.Fields(2)
    If Not IsNull(rs.Fields(3)) Then txtTelNo = rs.Fields(3)
    If Not IsNull(rs.Fields(4)) Then txtBox = rs.Fields(4)
    If Not IsNull(rs.Fields(5)) Then txtBank = rs.Fields(5)
    If Not IsNull(rs.Fields(6)) Then txtBBranch = rs.Fields(6)
    If Not IsNull(rs.Fields(7)) Then txtAccNo = rs.Fields(7)
   
End If
sql = "d_sp_SelectTrans  '" & txtTCode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
    If Not IsNull(rs.Fields(17)) Then Trip = rs.Fields(17)
    If Not IsNull(rs.Fields(18)) Then PerKg = rs.Fields(18)
    If Not IsNull(rs.Fields(16)) Then Flat = rs.Fields(16)
    If Trip Then
    Trip = True
    End If
    If PerKg Then
    PerKg = True
    End If
    If Flat Then
    Flat = True
    End If
    If PerKg = False And Trip = False Then
    Flat = True
    End If
End If
LoadData
End Sub
