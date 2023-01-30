VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frminvoicecustomer 
   BackColor       =   &H00FF00FF&
   Caption         =   "Create Customer Invoice"
   ClientHeight    =   8100
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsave 
      Caption         =   "Add >>"
      Height          =   375
      Left            =   1080
      TabIndex        =   20
      Top             =   4800
      Width           =   975
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
      TabIndex        =   18
      Top             =   2640
      Width           =   1170
   End
   Begin VB.PictureBox Picture4 
      Height          =   285
      Left            =   2400
      Picture         =   "frminvoicecustomer.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   17
      Top             =   2640
      Width           =   300
   End
   Begin VB.TextBox txtnarration 
      Height          =   525
      Left            =   1680
      TabIndex        =   16
      Top             =   3840
      Width           =   5535
   End
   Begin VB.TextBox txtdebtorAcc 
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
      TabIndex        =   15
      Top             =   3120
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      Height          =   285
      Left            =   2355
      Picture         =   "frminvoicecustomer.frx":02C2
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   14
      Top             =   3120
      Width           =   300
   End
   Begin VB.TextBox txtamount 
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtinvoiceNo 
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "RePrint"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox txtTCode 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1080
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtNames 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2760
      TabIndex        =   8
      Top             =   2160
      Width           =   4215
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   2400
      Picture         =   "frminvoicecustomer.frx":0584
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox txtdocNo 
      Height          =   285
      Left            =   5160
      TabIndex        =   6
      Top             =   4770
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Option"
      Height          =   615
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   4335
      Begin VB.OptionButton optmilk 
         Caption         =   "Millk Sales"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optothers 
         Caption         =   "Other Services"
         Height          =   240
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdpost 
      Caption         =   "Post"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "<<Remove"
      Height          =   345
      Left            =   7320
      TabIndex        =   1
      Top             =   4680
      Width           =   1530
   End
   Begin VB.TextBox txttotal 
      Height          =   375
      Left            =   7800
      TabIndex        =   0
      Top             =   4080
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPtransdate 
      Height          =   375
      Left            =   4320
      TabIndex        =   19
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
      Format          =   167641089
      CurrentDate     =   41927
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1560
      TabIndex        =   21
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
      Format          =   167641089
      CurrentDate     =   41927
   End
   Begin MSComctlLib.ListView lvwInvoice 
      Height          =   2655
      Left            =   0
      TabIndex        =   22
      Top             =   5280
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4683
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "InvoiceNo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Dcode"
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
         Text            =   "DebtorAccNO"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ContraAccNo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Particulars"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Cess"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "CessAccNo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "StartDate"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "EndDate"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Income Item"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "To"
      Height          =   255
      Left            =   3600
      TabIndex        =   32
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   120
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
      Left            =   2760
      TabIndex        =   30
      Top             =   2640
      Width           =   4170
   End
   Begin VB.Label Label3 
      Caption         =   "Narration"
      Height          =   255
      Left            =   360
      TabIndex        =   29
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lbldebtorname 
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
      TabIndex        =   28
      Top             =   3120
      Width           =   4170
   End
   Begin VB.Label Label8 
      Caption         =   "Debtors Acc"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   " Amount"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   " InvoiceNo"
      Height          =   255
      Left            =   3840
      TabIndex        =   25
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Debtors"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   " Total"
      Height          =   255
      Left            =   8280
      TabIndex        =   23
      Top             =   3840
      Width           =   975
   End
End
Attribute VB_Name = "frminvoicecustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pushed As Currency
Private Sub cmdnew_Click()
    DTPicker1 = Format(Get_Server_Date, "dd/mm/yyyy")
    DTPtransdate = DateSerial(year(DTPicker1), month(DTPicker1) + 1, 1 - 1)
    DTPicker1 = DateSerial(year(DTPicker1), month(DTPicker1), 1)
 
    txtnarration = ""
    txtcontra = ""
    txtkilos = 0
    txtamount = 0
    txtkilos = 0
    txtrate = 0
    txtdebtorAcc = ""
    lblcontra = ""
    lbldebtorname = ""
    txtcess = 0
    txttotal = 0
'    chkcess.Value = vbUnchecked
    Generate_InvoiceNo
    lvwInvoice.ListItems.Clear
End Sub

Private Sub cmdPost_Click()
Dim Kilos As Double, Price As Double, Qty As Double, dcode As String, DebtorAcc As String, ContrAcc As String
Dim Cess As Double, CessAcc As String, Amount As Double
Dim DRaccno As String, Craccno As String, chequeno As String, tdate As Date, edate As Date, _
TransSource As String, TransDescription As String, CashBook As String, doc_posted As String

    For I = 1 To lvwInvoice.ListItems.Count
    ' If lvwInvoice.ListItems.Item(I).Checked = True Then
            Set li = lvwInvoice.ListItems(I)
            DocumentNo = li
            dcode = lvwInvoice.ListItems(I).SubItems(1)
            Kilos = lvwInvoice.ListItems(I).SubItems(2)
            Price = lvwInvoice.ListItems(I).SubItems(3)
            Amount = lvwInvoice.ListItems(I).SubItems(4)
            DebtorAcc = lvwInvoice.ListItems(I).SubItems(5)
            ContrAcc = lvwInvoice.ListItems(I).SubItems(6)
            Cess = lvwInvoice.ListItems(I).SubItems(8)
            CessAcc = lvwInvoice.ListItems(I).SubItems(9)
            tdate = lvwInvoice.ListItems(I).SubItems(10)
            edate = lvwInvoice.ListItems(I).SubItems(11)
            TransDescription = lvwInvoice.ListItems(I).SubItems(7)
            CashBook = 1
            doc_posted = 1
            TransSource = dcode
            GetTransactionNo
          
           NewTransaction Amount, tdate, TransDescription
           
          If Not Save_GLTRANSACTION(tdate, Amount, DebtorAcc, ContrAcc, DocumentNo, _
            TransSource, User, ErrorMessage, TransDescription, CashBook, doc_posted, DocumentNo, transactionNo, "") Then
              If ErrorMessage <> "" Then
                  MsgBox ErrorMessage, vbInformation, Me.Caption
                  ErrorMessage = ""
              End If
          End If
          
          sql = " set dateformat dmy  INSERT INTO invoice"
           sql = sql & " (InvoiceNo,Dcode,SupplierAcc, IncomeAcc,Amount,StartDate, EndDate, Transdescription, Rate, Kilos,Auditid,TransactionNo,Transtype) "
           sql = sql & "  VALUES     (" & DocumentNo & ",'" & dcode & "','" & DebtorAcc & "','" & ContrAcc & "'," & Amount & " ,"
           sql = sql & "  '" & tdate & "','" & edate & "','" & TransDescription & "'," & CDbl(Price) & "," & CDbl(Kilos) & ",'" & User & "','" & transactionNo & "','CR')"
           oSaccoMaster.ExecuteThis (sql)

    Next I
       
        MsgBox "Invoice Created Successfuly", vbInformation, Me.Caption
        txtdocNo = txtinvoiceNo
        If optmilk.Value = True Then
          reportname = "Invoice.rpt"
        Else
         reportname = "Invoice.rpt"
        End If
        STRFORMULA = "{Invoice.InvoiceNo}=" & txtdocNo & ""
        Show_Sales_Crystal_Report STRFORMULA, reportname, ""
       
       cmdnew_Click
End Sub

Private Sub cmdprint_Click()
    STRFORMULA = "{Invoice.InvoiceNo}=" & txtdocNo & ""
    reportname = "Invoice.rpt"
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

Private Sub cmdsave_Click()
Dim Amount As Double, DRaccno As String, Craccno As String, _
    TransSource As String, TransDescription As String, CashBook As String, doc_posted As String, chequeno As String
  
  If optmilk.Value = False And optothers.Value = False Then
   MsgBox "Select the Invoice option first", vbInformation, Me.Caption
   Exit Sub
  End If

 If optothers.Value = True Then
'   txtkilos = 0
'   txtrate = 0
   If txtamount = "" Or txtamount = 0 Then
    MsgBox "Enter invoice Amount", vbInformation, Me.Caption
    txtamount.SetFocus
    Exit Sub
   End If
 End If
 
 If txtcontra = "" Then
   MsgBox "Enter Income Item ", vbInformation, Me.Caption
    txtcontra.SetFocus
  Exit Sub
 End If
 
  If txtdebtorAcc = "" Then
   MsgBox "Enter Debtor Accno ", vbInformation, Me.Caption
    txtdebtorAcc.SetFocus
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
    li.SubItems(5) = txtdebtorAcc
    li.SubItems(6) = txtcontra
    li.SubItems(7) = txtnarration '& " Between " & DTPicker1 & " And " & dtpTransDate
    li.SubItems(8) = 0
    li.SubItems(9) = ""
    li.SubItems(10) = DTPicker1
    li.SubItems(11) = DTPtransdate
    
 Recalculate
 txtkilos = 0
 txtrate = 0
 txtamount = 0
 txtnarration = ""

       Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub Form_Load()
lvwInvoice.ForeColor = vbBlue
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
   txtnarration = "Milk Sales"
End If
   
End Sub

Private Sub optothers_Click()
  If optothers.Value = True Then
   optmilk.Value = False
'   txtrate.Visible = False
'   Label5.Visible = False
'   Label1.Visible = False
'   txtkilos.Visible = False
   txtamount.Locked = False
   chkcess.Value = vbUnchecked
End If
End Sub

Private Sub Picture1_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtdebtorAcc = SearchValue
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
     frmSearchDebtors.Show vbModal
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

Private Sub txtdocNo_LostFocus()
If Val(txtdocNo) = 0 Then
        MsgBox "Please enter a valid Invoice No", vbInformation, Me.Caption
        txtrate.SetFocus
        Beep
        Exit Sub
    End If
End Sub


Private Sub txtdebtorAcc_Change()
 Dim Account As Acc_Details
    Account = Get_Acc_Details(txtdebtorAcc, ErrorMessage)
    If Account.accno <> "" Then
        lbldebtorname = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lbldebtorname = ""
    End If
End Sub
Sub Generate_InvoiceNo()
 sql = "select isnull(max(invoiceno),0) from Invoice"
  Set Rst = oSaccoMaster.GetRecordset(sql)
   If Not Rst.EOF Then
    txtinvoiceNo = Rst.Fields(0) + 1
   End If
  
End Sub

Private Sub txtTCode_Change()
sql = "d_sp_Selectdebtors '" & txtTCode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtNames = rs.Fields(0)
If Not IsNull(rs.Fields(15)) Then txtdebtorAcc = rs.Fields(15)
If Not IsNull(rs.Fields(16)) Then txtcontra = rs.Fields(16)
End If

End Sub

Private Sub txtTCode_Click()
  txtTCode_Change
End Sub


