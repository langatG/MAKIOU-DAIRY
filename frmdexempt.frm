VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmdexempt 
   BackColor       =   &H00FFC0FF&
   Caption         =   "DEDUCTIONS EXEMPT"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2280
      TabIndex        =   11
      Top             =   600
      Width           =   4575
   End
   Begin VB.CommandButton cmdremove 
      Caption         =   "Remove"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.ComboBox cboDeduct 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmdexempt.frx":0000
      Left            =   1440
      List            =   "frmdexempt.frx":0002
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox txtSNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtpSDate 
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   149159937
      CurrentDate     =   40209
   End
   Begin VB.Label lblRemarks 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Exempt Date"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2520
      TabIndex        =   7
      Top             =   0
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Deduction"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1230
   End
   Begin VB.Label lblSNo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "SNo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   480
   End
End
Attribute VB_Name = "frmdexempt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboDeduct_Change()
If UCase(cboDeduct.Text) = "OTHERS" Then
lblRemarks.Visible = True
txtRemarks.Visible = True
txtRemarks = ""
Else
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdcancel_Click()
 Unload Me
End Sub

Private Sub cmdremove_Click()
If txtSNo = "" Then
MsgBox "Enter the Supplier Number", vbInformation
Exit Sub
End If
If txtName = "" Then
MsgBox "Enter Valid Supplier Number", vbInformation
txtSNo.SetFocus
Exit Sub
End If
If cboDeduct.Text = "" Then
  MsgBox "Select Deduction Type", vbInformation
  cboDeduct.SetFocus
 Exit Sub
End If

If MsgBox("Do You Want to Remove Supplier " & txtSNo & "  From " & cboDeduct & "  Exemptions", vbYesNo + vbInformation, Me.Caption) = vbNo Then
  Exit Sub
End If
Set cn = New ADODB.Connection
sql = "delete from d_exempt where sno= " & txtSNo & " and description='" & cboDeduct & "'"
oSaccoMaster.ExecuteThis (sql)
txtSNo = ""
MsgBox "Record Removed successfully!"
End Sub

Private Sub cmdsave_Click()
On Error GoTo ErrorHandler
If txtSNo = "" Then
MsgBox "Enter the Supplier Number", vbInformation
txtSNo = ""
Exit Sub
End If
If txtName = "" Then
MsgBox "Enter Valid Supplier Number", vbInformation
txtSNo.SetFocus
Exit Sub
End If
If cboDeduct.Text = "" Then
  MsgBox "Select Deduction Type", vbInformation
  cboDeduct.SetFocus
 Exit Sub
End If
If MsgBox("Do You Want to Exempt Supplier " & txtSNo & "  From " & cboDeduct & "  Deductions", vbYesNo + vbInformation, Me.Caption) = vbNo Then
  Exit Sub
End If
Set cn = New ADODB.Connection
sql = "d_sp_exempt " & txtSNo & ",'" & cboDeduct & "','" & User & "','" & txtRemarks & "','" & dtpSDate & "'"
oSaccoMaster.ExecuteThis (sql)
txtSNo = ""
cboDeduct.Locked = False
MsgBox "Records Saved successfully!"
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    dtpSDate = Format(Get_Server_Date, "dd/mm/yyyy")
    cboDeduct.Clear
    Set rs = oSaccoMaster.GetRecordset("SELECT Description FROM d_DCodes")
    If Not rs.EOF Then
    With rs
        While Not .EOF
         cboDeduct.AddItem rs.Fields("Description")
         .MoveNext
        Wend
    End With
   End If
End Sub

Private Sub txtSNo_Change()
Set rs = New ADODB.Recordset
If txtSNo <> "" Then
    sql = "d_sp_SupplierEnquiry " & txtSNo & ""
    Set rs = oSaccoMaster.GetRecordset(sql)
    If rs.EOF Then
     txtName = ""
        MsgBox "There is no supplier with number " & txtSNo
        Exit Sub
    Else
    If Not IsNull(rs.Fields(0)) Then txtName = rs.Fields(0)
    End If
Else
txtName = ""
End If
End Sub
