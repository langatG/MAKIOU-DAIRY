VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmpaybills 
   BackColor       =   &H00800080&
   Caption         =   "Pay Bills"
   ClientHeight    =   8310
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   10170
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtremarks 
      Height          =   405
      Left            =   1440
      TabIndex        =   29
      Top             =   7080
      Width           =   5895
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   3120
      Picture         =   "frmpaybills.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   27
      Top             =   5040
      Width           =   255
   End
   Begin VB.TextBox txtNames 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   3480
      TabIndex        =   26
      Top             =   4920
      Width           =   4575
   End
   Begin VB.TextBox txtTCode 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1440
      TabIndex        =   25
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   7680
      Width           =   855
   End
   Begin VB.TextBox txtinvoiceNo 
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox txtamount 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   4440
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   285
      Left            =   3075
      Picture         =   "frmpaybills.frx":02C2
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   5520
      Width           =   300
   End
   Begin VB.TextBox txtcreditorAcc 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      TabIndex        =   5
      Top             =   5520
      Width           =   1530
   End
   Begin VB.TextBox txtnarration 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   6720
      Width           =   5775
   End
   Begin VB.PictureBox Picture4 
      Height          =   285
      Left            =   3000
      Picture         =   "frmpaybills.frx":0584
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   6000
      Width           =   300
   End
   Begin VB.TextBox txtcontra 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   6000
      Width           =   1530
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox txtoriginalamount 
      Height          =   285
      Left            =   5040
      TabIndex        =   0
      Top             =   4440
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTPtransdate 
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   3840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   166526977
      CurrentDate     =   41927
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   3735
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "InvoiceNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "InvoiceDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "DueDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Supplier"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "AccNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "SupplierID"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComCtl2.DTPicker DtpDueDate 
      Height          =   375
      Left            =   8400
      TabIndex        =   31
      Top             =   4680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   166526977
      CurrentDate     =   41927
   End
   Begin VB.Label Label13 
      Caption         =   " Due Date"
      Height          =   255
      Left            =   8400
      TabIndex        =   32
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Remarks"
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Suppliers"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   " InvoiceNo"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   " Amount"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Creditor Acc"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label lblcreditorname 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3480
      TabIndex        =   21
      Top             =   5520
      Width           =   4530
   End
   Begin VB.Label Label3 
      Caption         =   "Narration"
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label lblcontra 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3480
      TabIndex        =   19
      Top             =   6000
      Width           =   4530
   End
   Begin VB.Label Label2 
      Caption         =   "Debit Acc"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   5520
      Width           =   735
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080FF80&
      Caption         =   "ACCOUNT TO DEBIT"
      Height          =   255
      Left            =   8160
      TabIndex        =   17
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080FF80&
      Caption         =   "PAYMENT ACC"
      Height          =   255
      Left            =   8160
      TabIndex        =   16
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   " Payment Date"
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Balance"
      Height          =   255
      Left            =   3840
      TabIndex        =   14
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "ReceiptNo"
      Height          =   255
      Left            =   7200
      TabIndex        =   13
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lblrno 
      Height          =   375
      Left            =   8280
      TabIndex        =   12
      Top             =   3840
      Width           =   1815
   End
End
Attribute VB_Name = "frmpaybills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdnew_Click()
 txtcontra = ""
 lblcontra = ""
 txtcreditorAcc = ""
 lblcreditorname = ""
 txtamount = 0
 txtnarration = ""
 txtremarks = ""
 DTPtransdate = Format(Get_Server_Date, "dd/mm/yyyy")
End Sub

Private Sub cmdsave_Click()
Dim Amount As Double, DRaccno As String, Craccno As String, _
TransSource As String, TransDescription As String, CashBook As String, doc_posted As String, chequeno As String
 
 If txtinvoiceNo = "" Then
   MsgBox "Enter Invoice No ", vbInformation, Me.Caption
    txtnarration.SetFocus
  Exit Sub
 End If
  
  If txtcontra = "" Then
   MsgBox "Enter Debit Gl Account Item ", vbInformation, Me.Caption
    txtcontra.SetFocus
  Exit Sub
 End If

  If txtcreditorAcc = "" Then
   MsgBox "Enter  Account To Debit ", vbInformation, Me.Caption
    txtcreditorAcc.SetFocus
  Exit Sub
 End If
   If txtnarration = "" Then
   MsgBox "Enter Narration ", vbInformation, Me.Caption
    txtnarration.SetFocus
  Exit Sub
 End If
  
  If txtamount = "" Or txtamount = 0 Then
    MsgBox "Enter invoice Amount", vbInformation, Me.Caption
    txtamount.SetFocus
    Exit Sub
   End If
   
    transdate = Format(DTPtransdate, "dd/mm/yyyy")
    If transdate > Format(Get_Server_Date, "dd/mm/yyyy") Then
     MsgBox "  Cant Transact on a future Date"
     DTPtransdate.SetFocus
     Exit Sub
    End If
    
    Amount = CDbl(txtamount)
    DRaccno = txtcreditorAcc
    Craccno = txtcontra
    DocumentNo = txtinvoiceNo
    TransSource = lblcreditorname
    TransDescription = txtnarration
    CashBook = 1
    doc_posted = 1
    GetTransactionNo
    
         NewTransaction Amount, transdate, TransDescription
         
       If Not Save_GLTRANSACTION(transdate, Amount, DRaccno, Craccno, DocumentNo, _
      TransSource, User, ErrorMessage, TransDescription, CashBook, doc_posted, chequeno, transactionNo, "") Then
          If ErrorMessage <> "" Then
              MsgBox ErrorMessage, vbInformation, Me.Caption
              ErrorMessage = ""
          End If
      End If
      
       sql = " set dateformat dmy  INSERT INTO InvoicePayments"
       sql = sql & " (InvoiceNo,SupplierId,SupplierAccNo,Amount,TransactionNo,TransDate,DueDate, Particulars,Transtype,Remarks, receiptno,Auditid) "
       sql = sql & "  VALUES     (" & txtinvoiceNo & ",'" & txtTCode & "','" & txtcreditorAcc & "'," & Amount & " ,'" & transactionNo & "',"
       sql = sql & "  '" & transdate & "','" & Format(DtpDueDate, "dd/mm/yyyy") & "','" & TransDescription & "','DR','" & txtremarks & "','" & txtReceiptno & "','" & User & "')"
       oSaccoMaster.ExecuteThis (sql)
       
       MsgBox "Invoice Paid Successfuly", vbInformation, Me.Caption
        txtinvoiceNo = ""
        lblrno = ""
        txtoriginalamount = "0"
        txtsupplier = ""
        
       Form_Load
       cmdnew_Click
       Exit Sub
ErrorHandler:
MsgBox err.description

End Sub



Private Sub Form_Load()
cmdnew_Click
lvwItems.ListItems.Clear
Set rs = oSaccoMaster.GetRecordset("d_sp_PendingLiabilities")
While Not rs.EOF
Set li = lvwItems.ListItems.Add(, , rs.Fields(0))
                li.SubItems(1) = IIf(IsNull(rs.Fields(1)), "", rs.Fields(1)) & ""
                li.SubItems(2) = IIf(IsNull(rs.Fields(2)), "", rs.Fields(2)) & ""
                li.SubItems(3) = IIf(IsNull(rs.Fields(3)), "", rs.Fields(3)) & ""
                li.SubItems(4) = IIf(IsNull(rs.Fields(5)), "", rs.Fields(5)) & ""
                li.SubItems(5) = IIf(IsNull(rs.Fields(4)), "", rs.Fields(4)) & ""
                li.SubItems(6) = IIf(IsNull(rs.Fields(7)), "", rs.Fields(7)) & ""
                li.SubItems(7) = IIf(IsNull(rs.Fields(10)), "", rs.Fields(10)) & ""
                
rs.MoveNext
Wend

End Sub
Sub LoadInvoice(Supplier As String)
lvwItems.ListItems.Clear
Set rs = oSaccoMaster.GetRecordset("d_sp_PendingSLiabilities '" & Supplier & "'")
While Not rs.EOF
Set li = lvwItems.ListItems.Add(, , rs.Fields(0))
                li.SubItems(1) = IIf(IsNull(rs.Fields(1)), "", rs.Fields(1)) & ""
                li.SubItems(2) = IIf(IsNull(rs.Fields(2)), "", rs.Fields(2)) & ""
                li.SubItems(3) = IIf(IsNull(rs.Fields(3)), "", rs.Fields(3)) & ""
                li.SubItems(4) = IIf(IsNull(rs.Fields(5)), "", rs.Fields(5)) & ""
                li.SubItems(5) = IIf(IsNull(rs.Fields(4)), "", rs.Fields(4)) & ""
                li.SubItems(6) = IIf(IsNull(rs.Fields(7)), "", rs.Fields(7)) & ""
                li.SubItems(7) = IIf(IsNull(rs.Fields(10)), "", rs.Fields(10)) & ""
                
rs.MoveNext
Wend
End Sub

Private Sub lvwItems_DblClick()
On Error GoTo sysError
    If lvwItems.ListItems.Count > 0 Then
        txtinvoiceNo = lvwItems.SelectedItem
        txtTCode = lvwItems.SelectedItem.SubItems(7)
        DtpDueDate = lvwItems.SelectedItem.SubItems(2)
        txtoriginalamount = getBillBalance(txtTCode, Trim(txtinvoiceNo))
    End If
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Picture1_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtcreditorAcc = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub Picture4_Click()
   frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtcontra = SearchValue
            SearchValue = ""
        End If
    End If
End Sub

Private Sub Picture5_Click()
     frmsearchcreditors.Show vbModal
        txtTCode = sel
        txtTCode_Change
        LoadInvoice txtTCode
End Sub

Private Sub txtAmount_LostFocus()
On Error Resume Next
    If Val(txtamount) = 0 Then
        MsgBox "Please enter a valid Amount", vbInformation, Me.Caption
        txtrate.SetFocus
        Beep
        Exit Sub
    End If
End Sub

Private Sub txtcontra_Change()
Dim Account As Acc_Details
    Account = Get_Acc_Details(txtcontra, ErrorMessage)
    If Account.accno <> "" Then
        lblcontra = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblcontra = ""
    End If
End Sub

Private Sub txtcreditorAcc_Change()
Dim Account As Acc_Details
    Account = Get_Acc_Details(txtcreditorAcc, ErrorMessage)
    If Account.accno <> "" Then
        lblcreditorname = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblcreditorname = ""
    End If
End Sub
Private Sub txtTCode_Change()
sql = "d_sp_SelectCreditors '" & txtTCode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtNames = rs.Fields(0)
If Not IsNull(rs.Fields(4)) Then txtcreditorAcc = rs.Fields(4)
Else
txtNames = ""
txtcontra = ""
End If
End Sub

Private Sub txtTCode_Click()
  txtTCode_Change
End Sub


