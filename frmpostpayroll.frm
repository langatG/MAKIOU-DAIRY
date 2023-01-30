VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmpostpayroll 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Post Processed Payroll"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPurchaseAcc 
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
      TabIndex        =   13
      Top             =   1080
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      Height          =   285
      Left            =   2715
      Picture         =   "frmpostpayroll.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   300
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load"
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdpost 
      Caption         =   "Post"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   2160
      Width           =   975
   End
   Begin VB.PictureBox Picture4 
      Height          =   285
      Left            =   2760
      Picture         =   "frmpostpayroll.frx":02C2
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   1560
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
      TabIndex        =   1
      Top             =   1560
      Width           =   1170
   End
   Begin VB.TextBox txttotal 
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Text            =   "0"
      Top             =   240
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPtransdate 
      Height          =   375
      Left            =   480
      TabIndex        =   2
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
      Format          =   165806081
      CurrentDate     =   41927
   End
   Begin VB.Label lblPurchasename 
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
      Left            =   3120
      TabIndex        =   14
      Top             =   1080
      Width           =   4170
   End
   Begin VB.Label Label1 
      Caption         =   " EndMonth"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Farmers Creditors"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Milk Purchases"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
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
      Left            =   3120
      TabIndex        =   4
      Top             =   1560
      Width           =   4170
   End
   Begin VB.Label Label12 
      Caption         =   " Gross"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "frmpostpayroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim Rspayroll As ADODB.Recordset, mMonth As Integer, yYear As Integer
Private Sub cmdload_Click()
 Enddate = DateSerial(year(DTPtransdate), month(DTPtransdate) + 1, 1 - 1)
 mMonth = month(Enddate)
 yYear = year(Enddate)
 sql = "select isnull(sum(Gpay),0)Gpay from  d_Payroll where Mmonth=" & mMonth & " and Yyear=" & yYear & ""
 Set rs2 = oSaccoMaster.GetRecordset(sql)
 If Not rs2.EOF Then
 txttotal = Format(rs2.Fields(0), Cfmt)
 End If
End Sub

Private Sub cmdnew_Click()
txttotal = 0
End Sub

Private Sub cmdPost_Click()
Dim Kilos As Double, Price As Double, Qty As Double, dcode As String, DebtorAcc As String, ContrAcc As String
Dim Cess As Double, CessAcc As String, Amount As Double
Dim DRaccno As String, Craccno As String, chequeno As String, Tdate As Date, edate As Date, _
TransSource As String, TransDescription As String, CashBook As String, doc_posted As String

    If txttotal = "0" Then
       MsgBox "Please Load the Payroll Gross First Before Posting"
     Exit Sub
    End If
    
    If MsgBox("Post The Payroll To Affect Legders?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    Set TransPPosting = New ADODB.Connection
    TransPPosting.Open oSaccoMaster.goConn

    On Error GoTo TransError
    TransPPosting.BeginTrans
    
            Tdate = DTPtransdate
            CashBook = 1
            doc_posted = 1
            TransSource = "Farmers Payroll"
            DocumentNo = (mMonth) & "-" & yYear
            GetTransactionNo
            TransDescription = "Milk Purchases- " & MonthName(mMonth) & "-" & yYear
            Amount = CDbl(txttotal)
     
           NewTransaction Amount, Tdate, TransDescription
         
          If Not Save_GLTRANSACTION(Tdate, Amount, txtPurchaseAcc, txtcontra, DocumentNo, _
            TransSource, User, ErrorMessage, TransDescription, CashBook, doc_posted, DocumentNo, transactionNo, "") Then
              If ErrorMessage <> "" Then
                  MsgBox ErrorMessage, vbInformation, Me.Caption
                  ErrorMessage = ""
              End If
          End If
          MsgBox "Farmers Payroll Posted Successfully"
Exit Sub
TransError:
    MsgBox err.description
    TransPPosting.RollbackTrans
End Sub



Private Sub Form_Load()
DTPtransdate = Format(Get_Server_Date, "dd/mm/yyyy")
DTPtransdate = DateSerial(year(DTPtransdate), month(DTPtransdate), 1 - 1)
txtPurchaseAcc = "E101"
txtcontra = "L007"
End Sub

Private Sub Picture1_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtPurchaseAcc = SearchValue
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

Private Sub txtPurchaseAcc_Change()
Dim Account As Acc_Details
    Account = Get_Acc_Details(txtPurchaseAcc, ErrorMessage)
    If Account.accno <> "" Then
        lblPurchasename = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblPurchasename = ""
    End If
End Sub
