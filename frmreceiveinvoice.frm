VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmreceiveinvoice 
   Caption         =   "Receive Invoice"
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10650
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   10650
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtoriginalamount 
      Height          =   285
      Left            =   6480
      TabIndex        =   24
      Top             =   5040
      Width           =   3015
   End
   Begin VB.TextBox txtsupplier 
      Height          =   375
      Left            =   6480
      TabIndex        =   22
      Top             =   4440
      Width           =   3975
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   8040
      Width           =   1095
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
      Left            =   1320
      TabIndex        =   7
      Top             =   5520
      Width           =   1170
   End
   Begin VB.PictureBox Picture4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      Picture         =   "frmreceiveinvoice.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   5520
      Width           =   300
   End
   Begin VB.TextBox txtnarration 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   7440
      Width           =   5775
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
      Left            =   1320
      TabIndex        =   4
      Top             =   6840
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2595
      Picture         =   "frmreceiveinvoice.frx":02C2
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   6840
      Width           =   300
   End
   Begin VB.TextBox txtamount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox txtinvoiceNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   8040
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DTPtransdate 
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   3840
      Width           =   1815
      _ExtentX        =   3201
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
      Format          =   147980289
      CurrentDate     =   41927
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   3735
      Left            =   0
      TabIndex        =   20
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
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ReceiptNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ord Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ProductName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Ordered Qyt"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Delivery Qty"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Balance"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblrno 
      Height          =   375
      Left            =   8280
      TabIndex        =   26
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "RNO"
      Height          =   255
      Left            =   7200
      TabIndex        =   25
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Original Amount"
      Height          =   255
      Left            =   4560
      TabIndex        =   23
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Creditor:"
      Height          =   255
      Left            =   4560
      TabIndex        =   21
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   " Invoice Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   19
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080FF80&
      Caption         =   "CREDITOR ACCOUNT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080FF80&
      Caption         =   "ACCOUNT TO DEBIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   17
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Debit Acc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   5520
      Width           =   1095
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
      Left            =   3000
      TabIndex        =   15
      Top             =   5520
      Width           =   4170
   End
   Begin VB.Label Label3 
      Caption         =   "Narration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   7440
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
      Left            =   3000
      TabIndex        =   13
      Top             =   6840
      Width           =   4170
   End
   Begin VB.Label Label8 
      Caption         =   "Creditor Acc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   " Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   " InvoiceNo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Width           =   975
   End
End
Attribute VB_Name = "frmreceiveinvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdnew_Click()
 txtcontra = ""
 lblcontra = ""
 txtcreditorAcc = ""
 lblcreditorname = ""
 txtAmount = 0
 txtNarration = ""
 dtptransdate = Format(Get_Server_Date, "dd/mm/yyyy")
End Sub

Private Sub cmdsave_Click()
Dim Amount As Double, DRaccno As String, Craccno As String, _
      TransSource As String, TransDescription As String, CashBook As String, doc_posted As String, chequeno As String
   If txtinvoiceNo = "" Then
   MsgBox "Enter Invoice No ", vbInformation, Me.Caption
    txtNarration.SetFocus
  Exit Sub
 End If
  
  If txtcontra = "" Then
   MsgBox "Enter Debit Gl Account Item ", vbInformation, Me.Caption
    txtcontra.SetFocus
  Exit Sub
 End If

  If txtcreditorAcc = "" Then
   MsgBox "Enter Creditor Account ", vbInformation, Me.Caption
    txtcreditorAcc.SetFocus
  Exit Sub
 End If
   If txtNarration = "" Then
   MsgBox "Enter Narration ", vbInformation, Me.Caption
    txtNarration.SetFocus
  Exit Sub
 End If
   
    transdate = Format(dtptransdate, "dd/mm/yyyy")
    If transdate > Format(Get_Server_Date, "dd/mm/yyyy") Then
     MsgBox "  Cant Transact on a future Date"
     dtptransdate.SetFocus
     Exit Sub
    End If
    
    Amount = CDbl(txtAmount)
    DRaccno = txtcontra
    Craccno = txtcreditorAcc
    DocumentNo = txtinvoiceNo
    TransSource = lblcreditorname
    TransDescription = txtNarration
    CashBook = 1
    doc_posted = 1
    GetTransactionNo
    
       If Not Save_GLTRANSACTION(transdate, Amount, DRaccno, Craccno, DocumentNo, _
      TransSource, User, ErrorMessage, TransDescription, CashBook, doc_posted, chequeno, transactionNo, "") Then
          If ErrorMessage <> "" Then
              MsgBox ErrorMessage, vbInformation, Me.Caption
              ErrorMessage = ""
          End If
      End If
      
       sql = "set dateformat dmy  INSERT INTO InvoiceReceived"
       sql = sql & " (InvoiceNo,CreditorAccNo, DRAccNo,Amount,Transdate, Transdescription,Auditid,RNO) "
       sql = sql & "  VALUES     (" & txtinvoiceNo & ",'" & txtcreditorAcc & "','" & txtcontra & "'," & Amount & " ,"
       sql = sql & "  '" & transdate & "','" & TransDescription & "','" & User & "','" & lblrno & "')"
       oSaccoMaster.ExecuteThis (sql)
       
       
       sql = "Update d_Requisition set [status]='Invoiced'   where RNo='" & lblrno & "'"
  oSaccoMaster.ExecuteThis (sql)
       
       MsgBox "Invoice Received Successfuly", vbInformation, Me.Caption
        txtinvoiceNo = ""
        lblrno = ""
        txtoriginalamount = ""
        txtsupplier = ""
        
       Form_Load
       cmdnew_Click
       Exit Sub
ErrorHandler:
MsgBox err.description

End Sub



Private Sub Form_Load()
cmdnew_Click
Lvwitems.ListItems.Clear
Set rs = oSaccoMaster.GetRecordset("d_sp_loadOrderedGoods1")
'Set rs = oSaccoMaster.GetRecordset("d_sp_IOrdered")
While Not rs.EOF
Set li = Lvwitems.ListItems.Add(, , rs.Fields(0))
                li.SubItems(1) = rs.Fields(1) & ""
                li.SubItems(2) = rs.Fields(2) & ""
                li.SubItems(3) = rs.Fields(3) & ""
                li.SubItems(4) = rs.Fields(4) & ""
                li.SubItems(5) = "0" & ""
                li.SubItems(6) = rs.Fields(4) & ""
                
rs.MoveNext
Wend

End Sub

Private Sub lvwItems_DblClick()
edit Lvwitems.SelectedItem
End Sub
Public Sub edit(selected As String)
'//
sql = ""
sql = "SELECT * FROM d_Requisition WHERE RNo='" & Trim(selected) & "' "
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
txtoriginalamount = rs.Fields("totalprice")
txtsupplier = rs.Fields("costcentre")
lblrno = selected
'//GET THE GL NAME BEFORE POSTINGS
sql = ""
sql = "SELECT GLNO FROM ag_Supplier1 WHERE CompanyName ='" & txtsupplier & "'"
Set Rst = oSaccoMaster.GetRecordset(sql)
If Not Rst.EOF Then
txtcreditorAcc = IIf(IsNull(Rst.Fields(0)), "", Rst.Fields(0))
End If

Else
txtoriginalamount = ""
txtsupplier = ""
End If
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

Private Sub txtAmount_LostFocus()
On Error Resume Next
    If Val(txtAmount) = 0 Then
        MsgBox "Please enter a valid Amount", vbInformation, Me.Caption
        txtrate.SetFocus
        Beep
        Exit Sub
    End If
End Sub

Private Sub txtcontra_Change()
Dim Account As Acc_Details
    Account = Get_Acc_Details(txtcontra, ErrorMessage)
    If Account.AccNo <> "" Then
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
    If Account.AccNo <> "" Then
        lblcreditorname = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblcreditorname = ""
    End If
End Sub
