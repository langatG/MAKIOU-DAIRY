VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmcreatecustomerbill 
   BackColor       =   &H00C000C0&
   Caption         =   "Create Customer Bill"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRemove 
      Caption         =   "<<Remove"
      Height          =   345
      Left            =   3480
      TabIndex        =   33
      Top             =   4560
      Width           =   1050
   End
   Begin VB.TextBox txtdocNo 
      Height          =   285
      Left            =   8160
      TabIndex        =   16
      Top             =   4650
      Width           =   1815
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   2400
      Picture         =   "frmcreatecustomerbill.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   15
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox txtNames 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2760
      TabIndex        =   14
      Top             =   1920
      Width           =   4215
   End
   Begin VB.TextBox txtTCode 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1080
      TabIndex        =   13
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "RePrint"
      Height          =   255
      Left            =   7200
      TabIndex        =   12
      Top             =   4680
      Width           =   855
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   960
      TabIndex        =   11
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox txtinvoiceNo 
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtamount 
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   1200
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   285
      Left            =   2355
      Picture         =   "frmcreatecustomerbill.frx":02C2
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   300
   End
   Begin VB.TextBox txtCreditorAcc 
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
      Left            =   1080
      TabIndex        =   7
      Top             =   2880
      Width           =   1170
   End
   Begin VB.TextBox txtnarration 
      Height          =   405
      Left            =   1080
      TabIndex        =   6
      Top             =   3360
      Width           =   5895
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Add >>"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   4560
      Width           =   1095
   End
   Begin VB.PictureBox Picture4 
      Height          =   285
      Left            =   2400
      Picture         =   "frmcreatecustomerbill.frx":0584
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   2400
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
      Left            =   1080
      TabIndex        =   3
      Top             =   2400
      Width           =   1170
   End
   Begin VB.TextBox txtremarks 
      Height          =   405
      Left            =   1080
      TabIndex        =   2
      Top             =   3960
      Width           =   5895
   End
   Begin VB.CommandButton cmdpost 
      Caption         =   "Post"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox txttotal 
      Height          =   375
      Left            =   7440
      TabIndex        =   0
      Top             =   3720
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPtransdate 
      Height          =   375
      Left            =   960
      TabIndex        =   17
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
      Format          =   166461441
      CurrentDate     =   41927
   End
   Begin MSComCtl2.DTPicker DtpDueDate 
      Height          =   375
      Left            =   4080
      TabIndex        =   18
      Top             =   120
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
      Format          =   166526977
      CurrentDate     =   41927
   End
   Begin MSComctlLib.ListView lvwInvoice 
      Height          =   2535
      Left            =   0
      TabIndex        =   19
      Top             =   5040
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4471
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "InvoiceNo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "SupplierId"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "UnitPrice"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "SupplierAccNO"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ContraAccNo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "TransDate"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "DueDate"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Remarks"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "InvoiceNo"
      Height          =   255
      Left            =   8520
      TabIndex        =   32
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Suppliers"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   " InvoiceNo"
      Height          =   255
      Left            =   3000
      TabIndex        =   30
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FF80&
      Caption         =   " Amount"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "Suppliers Cr"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblCreditorname 
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
      Left            =   2760
      TabIndex        =   27
      Top             =   2880
      Width           =   4170
   End
   Begin VB.Label Label3 
      Caption         =   "Remarks"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "TransDate"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "DueDate"
      Height          =   255
      Left            =   3240
      TabIndex        =   24
      Top             =   240
      Width           =   735
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
      Left            =   2760
      TabIndex        =   23
      Top             =   2400
      Width           =   4170
   End
   Begin VB.Label Label2 
      Caption         =   "Debit Acc"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Particulars"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   " Total"
      Height          =   255
      Left            =   7920
      TabIndex        =   20
      Top             =   3480
      Width           =   975
   End
End
Attribute VB_Name = "frmcreatecustomerbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pushed As Currency
Private Sub cmdnew_Click()
    DTPtransdate = Format(Get_Server_Date, "dd/mm/yyyy")
    DtpDueDate = DateSerial(year(DTPtransdate), month(DTPtransdate) + 1, 1 - 1)
'    DtpDueDate = DateSerial(year(DtpDueDate), month(DtpDueDate), 1)
    txtremarks = ""
    txtnarration = ""
    txtcontra = ""
    txtkilos = 0
    txtamount = 0
    txtkilos = 0
    txtrate = 0
    txtCreditorAcc = ""
    lblcontra = ""
    lblCreditorname = ""
    lvwInvoice.ListItems.Clear
    Generate_InvoiceNo

End Sub

Private Sub cmdPost_Click()
Dim Kilos As Double, Price As Double, Qty As Double, scode As String, CreditorAcc As String, ContrAcc As String
Dim Cess As Double, CessAcc As String, Amount As Double, REMARKS As String
Dim DRaccno As String, Craccno As String, chequeno As String, tdate As Date, edate As Date, _
TransSource As String, TransDescription As String, CashBook As String, doc_posted As String

    For I = 1 To lvwInvoice.ListItems.Count
    ' If lvwInvoice.ListItems.Item(I).Checked = True Then
            Set li = lvwInvoice.ListItems(I)
            DocumentNo = li
            scode = lvwInvoice.ListItems(I).SubItems(1)
            Qty = lvwInvoice.ListItems(I).SubItems(2)
            Price = lvwInvoice.ListItems(I).SubItems(3)
            Amount = lvwInvoice.ListItems(I).SubItems(4)
            ContrAcc = lvwInvoice.ListItems(I).SubItems(5)
            CreditorAcc = lvwInvoice.ListItems(I).SubItems(6)
            tdate = lvwInvoice.ListItems(I).SubItems(8)
            edate = lvwInvoice.ListItems(I).SubItems(9)
            TransDescription = lvwInvoice.ListItems(I).SubItems(7)
            REMARKS = lvwInvoice.ListItems(I).SubItems(10)
            CashBook = 1
            doc_posted = 1
            TransSource = scode
            GetTransactionNo
            DRaccno = ContrAcc
            Craccno = CreditorAcc
                
                 NewTransaction Amount, tdate, TransDescription
                 
                 If Not Save_GLTRANSACTION(tdate, Amount, DRaccno, Craccno, DocumentNo, _
                    TransSource, User, ErrorMessage, TransDescription, CashBook, doc_posted, chequeno, transactionNo, "") Then
                        If ErrorMessage <> "" Then
                            MsgBox ErrorMessage, vbInformation, Me.Caption
                            ErrorMessage = ""
                        End If
                 End If
      
                sql = " set dateformat dmy  INSERT INTO InvoicePayments"
                sql = sql & " (InvoiceNo,SupplierId,SupplierAccno,Amount,Price,Qty,TransactionNo,TransDate,DueDate, Particulars,Transtype,Remarks, receiptno,Auditid) "
                sql = sql & "  VALUES     (" & DocumentNo & ",'" & scode & "','" & CreditorAcc & "'," & Amount & " ," & Price & "," & Qty & ",'" & transactionNo & "',"
                sql = sql & "  '" & tdate & "','" & edate & "','" & TransDescription & "','CR','" & REMARKS & "','" & txtReceiptno & "','" & User & "')"
                oSaccoMaster.ExecuteThis (sql)
       Next I
       
       MsgBox "Bill Created Successfuly", vbInformation, Me.Caption
       txtdocNo = txtinvoiceNo
        reportname = "Bill.rpt"
        STRFORMULA = "{InvoicePayments.InvoiceNo}=" & txtdocNo & ""
        Show_Sales_Crystal_Report STRFORMULA, reportname, ""
        cmdnew_Click
End Sub

Private Sub cmdprint_Click()
        txtdocNo = txtinvoiceNo
        reportname = "Bill.rpt"
        STRFORMULA = "{InvoicePayments.InvoiceNo}=" & txtdocNo & ""
        Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdRemove_Click()
     On Error GoTo sysError
    With lvwInvoice
        If .ListItems.Count > 0 Then
            If MsgBox("Do you want to remove " & .SelectedItem.SubItems(1) & _
            " From the list?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                Exit Sub
            End If
            pushed = pushed - .SelectedItem.ListSubItems(2)
            .ListItems.Remove (.SelectedItem.Index)
        End If
    End With
    Recalculate
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdsave_Click()
Dim Amount As Double, DRaccno As String, Craccno As String, _
TransSource As String, TransDescription As String, CashBook As String, doc_posted As String, chequeno As String
      
 If txtamount = "" Or txtamount = 0 Then
    MsgBox "Enter invoice Amount", vbInformation, Me.Caption
    txtamount.SetFocus
    Exit Sub
 End If

 If txtcontra = "" Then
   MsgBox "Enter Income Item ", vbInformation, Me.Caption
    txtcontra.SetFocus
  Exit Sub
 End If
 
 If txtCreditorAcc = "" Then
   MsgBox "Enter Debtor Accno ", vbInformation, Me.Caption
    txtCreditorAcc.SetFocus
  Exit Sub
 End If
 
 If txtnarration = "" Then
   MsgBox "Enter Narration ", vbInformation, Me.Caption
    txtnarration.SetFocus
  Exit Sub
 End If

    Set li = lvwInvoice.ListItems.Add(, , txtinvoiceNo)
      li.SubItems(1) = txtTCode
      li.SubItems(2) = 1
      li.SubItems(3) = 1
      li.SubItems(4) = Format(CDbl(txtamount), "#,##0.00")
      li.SubItems(5) = txtcontra
      li.SubItems(6) = txtCreditorAcc
      li.SubItems(7) = txtnarration
      li.SubItems(8) = Format(DTPtransdate, "dd/mm/yyyy")
      li.SubItems(9) = Format(DtpDueDate, "dd/mm/yyyy")
      li.SubItems(10) = txtremarks
 
 Recalculate
 txtkilos = 0
 txtrate = 0
 txtamount = 0
 txtremarks = ""
 txtnarration = ""
       Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub Form_Load()
cmdnew_Click
End Sub



Private Sub mnucreditorsaging_Click()
'//start the staff here
Dim invoiceno As String
Dim Amount As Double
Dim dcode As String
Dim rsp As New ADODB.Recordset
Dim gl As String
Dim glamount As Double
Dim days As Integer
Dim sdate As Date
Dim invdate As Date
Dim aamount As Double
Dim d1 As Integer, d2 As Integer, d3 As Integer, d4 As Integer, d5 As Integer
sql = ""
sql = "truncate table d_debtors_aging"
oSaccoMaster.ExecuteThis (sql)

'first loop then second loop
sql = "select invoiceno from invoice order by invoiceno"
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
invoiceno = rs.Fields(0)
        sql = ""
        sql = "set dateformat dmy SELECT invoiceno,dcode,supplieracc,amount,startdate FROM invoice where invoiceno=" & invoiceno & " ORDER BY Invoiceno,dcode"
        Set Rst = oSaccoMaster.GetRecordset(sql)
        While Not Rst.EOF
        invoiceno = Rst.Fields(0)
        Amount = Rst.Fields(3)
        dcode = Rst.Fields(1)
        gl = Rst.Fields(2)
        sdate = Format(Get_Server_Date, "dd / mm / yyyy")
        invdate = Rst.Fields(4)
        days = DateDiff("d", invdate, sdate)
        '/check if the amount is paid
        sql = ""
        sql = "SELECT amount FROM  GLTRANSACTIONS WHERE DrAccNo ='" & gl & "' AND DocumentNo='" & invoiceno & "'"
        Set rsp = oSaccoMaster.GetRecordset(sql)
        If Not rsp.EOF Then
        glamount = rsp.Fields(0)
        Else
        glamount = 0
        End If
        If glamount > 0 Then
        aamount = Amount - glamount
        Else
        aamount = Amount
        End If
        
        If aamount = 0 Then
       GoTo horola1
        End If
        If days <= 30 Then
        d1 = days
        GoTo horola
        End If
        If days > 30 And days <= 60 Then
        d2 = days
        GoTo horola
        End If
        If days > 60 And days <= 90 Then
        d3 = days
        GoTo horola
        End If
        If days > 90 And days <= 180 Then
        d4 = days
        GoTo horola
        End If
        If days > 180 Then
        d5 = days
        GoTo horola
        End If
horola:
        sql = ""
        sql = "INSERT INTO d_debtors_aging(dcode,Invoiceno,amount,[upto 30],[upto 60],[upto 90],[upto 180],[over 180]) VALUES('" & dcode & "','" & invoiceno & "'," & Amount & "," & d1 & "," & d2 & "," & d3 & "," & d4 & "," & d5 & ")"
        oSaccoMaster.ExecuteThis (sql)
        d1 = 0
        d2 = 0
        d3 = 0
        d4 = 0
        d5 = 0
        
        Rst.MoveNext
        Wend
horola1:
rs.MoveNext
Wend

'd_aging.rpt
 reportname = "d_aging.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
MsgBox "Record successfully generated", vbInformation

End Sub

Private Sub optmilk_Click()
 If optmilk.Value = True Then
   optothers.Value = False
   txtrate.Visible = True
   Label1.Visible = True
   Label5.Visible = True
   txtkilos.Visible = True
   txtamount.Locked = True
End If
   
End Sub

Private Sub optothers_Click()
  If optothers.Value = True Then
   optmilk.Value = False
   txtrate.Visible = False
   Label5.Visible = False
   Label1.Visible = False
   txtkilos.Visible = False
   txtamount.Locked = False
End If
End Sub

Private Sub Picture1_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtCreditorAcc = SearchValue
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
    Account = Get_Acc_Details(txtCreditorAcc, ErrorMessage)
    If Account.accno <> "" Then
        lblCreditorname = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblCreditorname = ""
    End If
End Sub
Sub Generate_InvoiceNo()
 sql = "select isnull(max(invoiceno),0) from InvoicePayments"
  Set Rst = oSaccoMaster.GetRecordset(sql)
   If Not Rst.EOF Then
    txtinvoiceNo = Rst.Fields(0) + 1
   End If
  
End Sub


Private Sub txtTCode_Change()
sql = "d_sp_SelectCreditors '" & txtTCode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtNames = rs.Fields(0)
If Not IsNull(rs.Fields(4)) Then txtCreditorAcc = rs.Fields(4)
Else
txtNames = ""
End If
End Sub

Private Sub txtTCode_Click()
  txtTCode_Change
End Sub
Private Sub Recalculate()
Dim balance As Double
    txttotal.Text = 0
    balance = 0
    If lvwInvoice.ListItems.Count > 0 Then
        For I = 1 To lvwInvoice.ListItems.Count
            balance = balance + CDbl(lvwInvoice.ListItems(I).SubItems(4))
        Next I
    End If
    txttotal = Format(balance, Cfmt)
End Sub




