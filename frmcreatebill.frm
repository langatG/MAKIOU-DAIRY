VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmcreatebill 
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   Picture         =   "frmcreatbill.frx":0000
   ScaleHeight     =   5010
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   23
      Top             =   2280
      Width           =   1170
   End
   Begin VB.PictureBox Picture4 
      Height          =   285
      Left            =   2400
      Picture         =   "frmcreatbill.frx":0342
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   22
      Top             =   2280
      Width           =   300
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox txtnarration 
      Height          =   525
      Left            =   1320
      TabIndex        =   10
      Top             =   3360
      Width           =   5535
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
      TabIndex        =   9
      Top             =   2760
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      Height          =   285
      Left            =   2355
      Picture         =   "frmcreatbill.frx":0604
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   300
   End
   Begin VB.TextBox txtamount 
      Height          =   375
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtinvoiceNo 
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox txtTCode 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1080
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtNames 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   2760
      TabIndex        =   2
      Top             =   1680
      Width           =   4215
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   2400
      Picture         =   "frmcreatbill.frx":08C6
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtdocNo 
      Height          =   285
      Left            =   5040
      TabIndex        =   0
      Top             =   4290
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPtransdate 
      Height          =   375
      Left            =   960
      TabIndex        =   11
      Top             =   240
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
      Format          =   123207681
      CurrentDate     =   41927
   End
   Begin MSComCtl2.DTPicker DtpDueDate 
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   240
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
      Format          =   123207681
      CurrentDate     =   41927
   End
   Begin VB.Label Label2 
      Caption         =   "Debit Acc"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2280
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
      Left            =   2760
      TabIndex        =   24
      Top             =   2280
      Width           =   4170
   End
   Begin VB.Label Label4 
      Caption         =   "DueDate"
      Height          =   255
      Left            =   3240
      TabIndex        =   21
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "TransDate"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Narration"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3600
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
      TabIndex        =   18
      Top             =   2760
      Width           =   4170
   End
   Begin VB.Label Label8 
      Caption         =   "Suppliers Cr"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   " Amount"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   " InvoiceNo"
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Suppliers"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   735
   End
End
Attribute VB_Name = "frmcreatebill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdnew_Click()
    DTPicker1 = Format(Get_Server_Date, "dd/mm/yyyy")
    DTPtransdate = DateSerial(year(DTPicker1), month(DTPicker1) + 1, 1 - 1)
    DTPicker1 = DateSerial(year(DTPicker1), month(DTPicker1), 1)
 
    txtnarration = ""
    txtcontra = ""
    txtkilos = 0
    txtamount = 0
    txtkilos = 0
    txtRate = 0
    txtCreditorAcc = ""
    lblcontra = ""
    lbldebtorname = ""
    
    Generate_InvoiceNo

End Sub

Private Sub cmdprint_Click()
    STRFORMULA = "{Invoice.InvoiceNo}=" & txtdocNo & ""
        reportname = "Invo.rpt"
        Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdsave_Click()
Dim Amount As Double, DRaccno As String, Craccno As String, _
      TransSource As String, TransDescription As String, CashBook As String, doc_posted As String, chequeno As String
   
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
 
 transdate = Format(DTPtransdate, "dd/mm/yyyy")
 Amount = CDbl(txtamount)
 DRaccno = txtCreditorAcc
 Craccno = txtcontra
 DocumentNo = txtinvoiceNo
 TransSource = lblcontra
 TransDescription = txtnarration
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
      
      
       sql = " set dateformat dmy  INSERT INTO invoice"
       sql = sql & " (InvoiceNo,Dcode,SupplierAcc, IncomeAcc,Amount,StartDate, EndDate, Transdescription, Rate, Kilos,Auditid) "
       sql = sql & "  VALUES     (" & txtinvoiceNo & ",'" & txtTCode & "','" & txtCreditorAcc & "','" & txtcontra & "'," & Amount & " ,"
       sql = sql & "  '" & Format(DTPicker1, "dd/mm/yyyy") & "','" & transdate & "','" & TransDescription & "'," & CDbl(txtRate) & "," & CDbl(txtkilos) & ",'" & User & "')"
       oSaccoMaster.ExecuteThis (sql)
       
       MsgBox "Invoice Created Successfuly", vbInformation, Me.Caption
        txtdocNo = txtinvoiceNo
        If optmilk.Value = True Then
          reportname = "Invoice.rpt"
        Else
         reportname = "Invoicess.rpt"
        End If
        STRFORMULA = "{Invoice.InvoiceNo}=" & txtdocNo & ""
        
        Show_Sales_Crystal_Report STRFORMULA, reportname, ""
       
       cmdnew_Click
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
   txtRate.Visible = True
   Label1.Visible = True
   Label5.Visible = True
   txtkilos.Visible = True
   txtamount.Locked = True
End If
   
End Sub

Private Sub optothers_Click()
  If optothers.Value = True Then
   optmilk.Value = False
   txtRate.Visible = False
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
            txtsupplierAcc = SearchValue
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
        Me.MousePointer = 0
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
        txtRate.SetFocus
        Beep
        Exit Sub
    End If
End Sub

Private Sub txtkilos_LostFocus()
If Val(txtkilos) = 0 Then
        MsgBox "Please enter a valid Amount", vbInformation, Me.Caption
        txtRate.SetFocus
        Beep
        Exit Sub
    End If
If txtRate = "" Then txtRate = 0
If txtkilos = "" Then txtkilos = 0
txtamount = CDbl(txtRate * txtkilos)
End Sub



Private Sub txtrate_LostFocus()
If Val(txtRate) = 0 Then
        MsgBox "Please enter a valid Amount", vbInformation, Me.Caption
        txtRate.SetFocus
        Beep
        Exit Sub
    End If
If txtRate = "" Then txtRate = 0
If txtkilos = "" Then txtkilos = 0
txtamount = CDbl(txtRate * txtkilos)
End Sub

Private Sub txtCreditorAcc_Change()
 Dim Account As Acc_Details
    Account = Get_Acc_Details(txtCreditorAcc, ErrorMessage)
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
If Not IsNull(rs.Fields(15)) Then txtCreditorAcc = rs.Fields(15)
Else
txtNames = ""

End If
End Sub

Private Sub txtTCode_Click()
  txtTCode_Change
End Sub


