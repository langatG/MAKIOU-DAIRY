VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransactions 
   BackColor       =   &H00C0C000&
   Caption         =   "Reverse Transactions"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13125
   LinkTopic       =   "Form6"
   ScaleHeight     =   8670
   ScaleWidth      =   13125
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboUser 
      Height          =   315
      ItemData        =   "frmtransactions.frx":0000
      Left            =   1560
      List            =   "frmtransactions.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtDocumentno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   19
      Top             =   720
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker dtpCurrentDate 
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   106889217
      CurrentDate     =   40522
   End
   Begin VB.CheckBox chkSorted 
      Caption         =   "Sorted"
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtpTransDate 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   121307137
      CurrentDate     =   40367
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   12632064
      TabCaption(0)   =   "TRANSACTIONS"
      TabPicture(0)   =   "frmtransactions.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lvwTransactions"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtReason"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "WITHDRAWN MEMBERS"
      TabPicture(1)   =   "frmtransactions.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "lvwWithdrawn"
      Tab(1).Control(2)=   "txtReasons"
      Tab(1).Control(3)=   "Frame2"
      Tab(1).ControlCount=   4
      Begin VB.TextBox txtReason 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Top             =   6720
         Width           =   5055
      End
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   -66600
         TabIndex        =   14
         Top             =   2400
         Width           =   1215
         Begin VB.CommandButton cmdWReverse 
            Caption         =   "GO!"
            Height          =   495
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   855
         End
         Begin VB.OptionButton optWReverse 
            Caption         =   "Reverse"
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   9600
         TabIndex        =   10
         Top             =   6600
         Width           =   2295
         Begin VB.OptionButton optDelete 
            Caption         =   "Delete"
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optReverse 
            Caption         =   "Reverse"
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton cmdReverse 
            Caption         =   "GO!"
            Height          =   495
            Left            =   960
            TabIndex        =   11
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox txtReasons 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71880
         TabIndex        =   8
         Top             =   4440
         Width           =   5175
      End
      Begin MSComctlLib.ListView lvwTransactions 
         Height          =   6135
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   10821
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777088
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TransNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "TransDate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "DocumentNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "UserId"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "AuditTime"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "TransDescription"
            Object.Width           =   7937
         EndProperty
      End
      Begin MSComctlLib.ListView lvwWithdrawn 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6376
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777088
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Memberno"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Names"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Withdrawal Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Withdrawal Date"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Transactionno"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label4 
         Caption         =   "Amendment Reason"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Amendment Reason"
         Height          =   255
         Left            =   -73560
         TabIndex        =   9
         Top             =   4440
         Width           =   1575
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Documentno"
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "UserId"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Current Date"
      Height          =   255
      Left            =   5880
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Transaction Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim totalshares As Double
Dim ReceiptNo As String

Private Sub chkSorted_Click()
    LoadTransactions
End Sub

Private Sub cmdreverse_Click()
    Dim TempAccNo As String, rsrepay As Recordset, LBalance As Double
    Dim j As Integer, fperiod As Integer, description As String, chequeno As String
    Dim action As String, Maturitydate As Date, IntRate As Double
    Dim post As New ADODB.Connection, CurrentDate As Date, Amount As Double
    
    j = 0
    Dim rsContrib As ADODB.Recordset
''    On Error GoTo UndoTrans
''    If lvwTransactions.ListItems.Count = 0 Then
''        Exit Sub
''    End If
''    For I = 1 To lvwTransactions.ListItems.Count
''        If lvwTransactions.ListItems(I).Checked = True Then
''            j = j + 1
''        End If
''    Next I
''    If optDelete.value = True Then
''        action = "Delete"
''    Else
''        action = "Reverse"
''    End If
''
''    If txtReason = "" Then
''        MsgBox "Reason for Reversal?", vbInformation
''        txtReason.SetFocus
''      Exit Sub
''    End If
''
''
''    serverDate = Format(Get_Server_Date, "dd/mm/yyyy")
''    CurrentDate = Format(dtpTransDate, "dd/mm/yyyy")
''
''    If CurrentDate < "04/03/2017" Then
''       MsgBox " Transaction Date is Previous System Change over ,Please Seek Permission from the Administration  ", vbCritical
''      Exit Sub
''    End If
''
''    If CurrentDate > serverDate Then
''       MsgBox "Oops! Check the Transdate,Can not Transact on a future date", vbCritical
''      Exit Sub
''    End If
''
''    If MsgBox("Are you sure you want to " & action & " the selected " & j & " Transaction(s), Noting that the Process is Irreversible?", vbQuestion + vbYesNo) = vbNo Then
''        Exit Sub
''    End If
''
''With post
''   .Open "FOSA"
''    .BeginTrans
''        Dim repayid As Double
''        With lvwTransactions
''            If optDelete.value Then
''            MsgBox "Consult System Admin for help!", vbInformation, Me.Caption
''                If Not .ListItems.Count <= 0 Then
''                    For I = 1 To .ListItems.Count
''                        If .ListItems(I).Checked = True Then
''                            TransNo = .ListItems(I).text
''                            transactionNo = TransNo
''                            Amount = Format(lvwTransactions.ListItems(I).ListSubItems(2), Cfmt)
''                            '************Start the Delition process**************
''                            'Start with Loans
''                            sql = "select distinct Loanno,sum(principal) as principal,sum(interest) as interest,sum(intrOwed) as intOwed,max(paymentno) as paymentno,sum(loanbalance) loanbalance from repay where transactionno='" & transactionNo & "' group by loanno"
''                            Set rsrepay = oSaccoMaster.GetRecordset(sql)
''                            With rsrepay
''                                If Not .EOF Then
''                                    While Not .EOF
''                                        LoanNo = .Fields("Loanno")
''                                        principal = .Fields("Principal")
''                                        interest = .Fields("Interest")
''                                        PaymentNo = .Fields("PaymentNo")
''                                        Loanbalance = .Fields("LoanBalance")
''                                        Set rs = oSaccoMaster.GetRecordset("SELECT INTROWED,penalty,datereceived FROM REPAY WHERE LOANNO='" & LoanNo & "' AND PAYMENTNO=" & PaymentNo - 1 & "")
''                                        If Not rs.EOF Then
''                                            intOwed = IIf(IsNull(rs("IntrOwed")), 0, rs("IntrOwed"))
''                                            Penalty = IIf(IsNull(rs("Penalty")), 0, rs("Penalty"))
''                                            lastrepay = rs("datereceived")
''                                        End If
''                                        PaymentNo = PaymentNo + 1
''                                        'Just Delete this transaction
''                                        oSaccoMaster.ExecuteThis ("Update loanbal set balance=" & Loanbalance + principal & ",intrOwed=" & intOwed & ",lastDate='" & lastrepay & "' where loanno='" & LoanNo & "'")
''                                        If success = False Then
''                                            GoTo UndoTrans
''                                        End If
''                                        .MoveNext
''                                    Wend
''                                    If Not oSaccoMaster.Execute("Delete from repay where transactionno='" & transactionNo & "'") Then
''                                        GoTo UndoTrans
''                                    End If
''                                Else
''
''                                End If
''                            End With
''
''                            'Next: Shares
''                            oSaccoMaster.ExecuteThis ("delete from contrib where transactionno='" & transactionNo & "'")
''                            If success = False Then
''                                GoTo UndoTrans
''                            End If
''
''                        End If
''                    Next I
''                    'Delete the transaction from gltransactions table
''                    oSaccoMaster.ExecuteThis ("Delete from gltransactions where transactionNo='" & transactionNo & "' ")
''                    If success = False Then
''                        GoTo UndoTrans
''                    End If
''
''                       ' find customer balance transactions
''                       Set rs = oSaccoMaster.GetRecordset("select * from CustomerBalance where transactionno='" & transactionNo & "'")
''
''                    If Not rs.EOF Then
''                       Description = rs!TransDescription
''
''                     '//Update Balance in CUB
''                    sql = "SELECT distinct (accno),balance From CuB where AccNo in(select AccNo from CustomerBalance where transactionNo='" & transactionNo & "') "
''                    Set rs2 = New ADODB.Recordset
''                    Set rs2 = oSaccoMaster.GetRecordset(sql)
''                    If Not rs2.EOF Then
''                        LBalance = IIf(rs2!balance <= 0, rs!balance, rs2!balance)
''                            If Description = "Account Deposit" Then
''
''                             LBalance = LBalance - Amount
''                             ElseIf Description = "Account Withdrawal" Then
''
''                            LBalance = LBalance + Amount
''                           End If
''                           ' Delete from customerbalance table
''                           oSaccoMaster.Execute ("delete from CustomerBalance where transactionno='" & transactionNo & "'")
''
''                     End If
''                       ' update Cub Now
''                    oSaccoMaster.Execute ("Update Cub set Balance= " & LBalance & " where Accno='" & rs2.Fields(0) & "'")
''
''                    End If
''
''                        'delete from cheque Deposits
''
''                        oSaccoMaster.Execute ("delete from  ChequeDeposits where transactionno='" & transactionNo & "'")
''
''                                        'Update Transaction as Deleted
''
''                   Set rs = oSaccoMaster.GetRecordset("Update Transactions set Status='DELETED',Reason='" & txtReason.text & "',Deletedby='" & User & "' where transactionno='" & transactionNo & "'")
''
''
''                    ' ****** Delete  Now VoucherNo from PaymentBooking *******/
''                             oSaccoMaster.Execute ("delete from PaymentBooking where transactionno='" & transactionNo & "'")
''
''                    ' ****** Delete  Now VoucherNo from ReceiptBooking *******/
''                             oSaccoMaster.Execute ("delete from ReceiptBooking where transactionno='" & transactionNo & "'")
''
''
''                    If success = False Then
''                        GoTo UndoTrans
''                    End If
''                Else
''                    Exit Sub
''''                End If
''            ElseIf optReverse.value = True Then
''                If Not .ListItems.Count <= 0 Then
''                    For I = 1 To .ListItems.Count
''                        If .ListItems(I).Checked = True Then
''                            TransNo = .ListItems(I).Text
''                            transactionNo = TransNo
''                            dtpTransDate.value = Format(.ListItems(I).ListSubItems(1), "dd/mm/yyyy")
''                            'Start with Loans
''                            sql = "select distinct Loanno,sum(principal) as principal,sum(interest) as interest,max(paymentno) as paymentno,Receiptno from repay where transactionno='" & transactionNo & "' group by loanno,Receiptno"
''                            Set rsrepay = oSaccoMaster.GetRecordset(sql)
''                            With rsrepay
''                                If Not .EOF Then
''                                    While Not .EOF
''                                        LoanNo = .Fields("Loanno")
''                                        principal = .Fields("Principal")
''                                        interest = .Fields("Interest")
''                                        paymentno = .Fields("PaymentNo")
''                                        ReceiptNo = IIf(IsNull(.Fields("Receiptno")), "", .Fields("Receiptno"))
''                                        Set rs = oSaccoMaster.GetRecordset("SELECT INTROWED,penalty FROM REPAY WHERE LOANNO='" & LoanNo & "' AND transactionno='" & transactionNo & "'")
''                                        If Not rs.EOF Then
''                                            intOwed = IIf(IsNull(rs("IntrOwed")), 0, rs("IntrOwed"))
''                                            Penalty = IIf(IsNull(rs("Penalty")), 0, rs("Penalty"))
''                                        End If
''
''                                        paymentno = paymentno + 1
''                                        'update loanbalance
''                                        oSaccoMaster.ExecuteThis ("Update loanbal set balance=balance+" & principal & ",intrOwed=" & intOwed & ",lastDate='" & dtpTransDate.value & "' where loanno='" & LoanNo & "'")
''                                        If success = False Then
''                                            GoTo UndoTrans
''                                        End If
''
''                                        Set rs = oSaccoMaster.GetRecordset("Select * from repay where transactionNo='" & transactionNo & "' ")
''
''                                        sql = "Insert into Repay(Loanno,Datereceived,Paymentno,Amount,Principal,Interest,IntrOwed,Penalty,intbalance,Loanbalance,Receiptno,TransBy,Remarks,auditid,IntrCharged,TransactionNo)" _
''                                        & " Values('" & LoanNo & "','" & dtpTransDate.value & "'," & paymentno & "," & (principal + interest) * (-1) & "," & principal * (-1) & "," & interest * (-1) & "," & intOwed & "," & Penalty * (-1) & "" _
''                                        & " ," & rs!intbalance * (-1) & "," & rs!loanbalance + principal & ",'" & ReceiptNo & "','Reversal','" & "Reversal" & "','" & User & "',0,'" & transactionNo & "')"
''
''                                        oSaccoMaster.ExecuteThis (sql)
''
''                                        If success = False Then
''                                            GoTo UndoTrans
''                                        End If
''                                        .MoveNext
''                                    Wend
''                                Else
''
''                                End If
''                            End With
''
''                             'Non Milk Advance
''                             Dim AdvNo As String
''
''
''
''
''                            'Next: BOSA Shares
''                            sql = "select distinct *  from Contrib where transactionno='" & transactionNo & "'"
''                            Set rsContrib = oSaccoMaster.GetRecordset(sql)
''                            With rsContrib
''                                If Not .EOF Then
''                                    While Not .EOF
''                                        mMemberNo = .Fields("memberno")
''                                        sharesCode = .Fields("Sharescode")
''                                        Amount = .Fields("Amount")
''                                        ReceiptNo = .Fields("Receiptno")
''                                        IntrAmount = .Fields("interest")
''                                        fperiod = .Fields("fperiod")
''                                        Maturitydate = .Fields("Maturitydate")
''                                        IntRate = .Fields("intRate")
''
''                                        'Reverse share transaction
''                                        Set rs = oSaccoMaster.GetRecordset("SELECT isnull(max(REFNO),0) FROM CONTRIB WHERE MEMBERNO='" & mMemberNo & "'")
''                                        If Not rs.EOF Then
''                                            Refno = rs(0) + 1
''                                        Else
''                                            Refno = 1
''                                        End If
''
''                                        totalshares = Amount
''                                        oSaccoMaster.ExecuteThis ("set dateformat dmy Insert into Contrib(memberno,contrdate,refno,Amount,sharebal,interest,transby,ChequeNo,receiptno,remarks,auditid,sharescode,transactionno)" _
''                                        & " Values('" & mMemberNo & "','" & Format(dtpTransDate, "DD/MM/YYYY") & "'," & Refno & "," & Amount * (-1) & "," & totalshares & "," & IntrAmount * (-1) & ",'','','" & ReceiptNo & "','Reversal -" & ReceiptNo & "' ,'" & User & "','" & sharesCode & "','" & transactionNo & "')")
''
''                                        If success = False Then
''                                            GoTo UndoTrans:
''                                        End If
''                                        .MoveNext
''                                    Wend
''                                Else
''
''                                End If
''                            End With
''
''                                If Format(dtpTransDate, "DD/MM/YYYY") < "06/02/2018" Then
''                                   MsgBox "Cannot Suucessfully Reverse Entrance Fee Before 05/02/2018", vbInformation, Me.Caption
''                                  GoTo Jump:
''                                End If
''                            'FOSA ENTRANCE FEE
''                            sql = "select distinct *  from Contrib where transactionno='" & transactionNo & "'"
''                            Set rsContrib = oSaccoMaster.GetRecordset(sql)
''                            With rsContrib
''                                If Not .EOF Then
''                                    While Not .EOF
''                                        mMemberNo = .Fields("memberno")
''                                        sharesCode = .Fields("Sharescode")
''                                        Amount = .Fields("Amount")
''                                        ReceiptNo = .Fields("Receiptno")
''
''                                        totalshares = Amount
''                                        oSaccoMaster.Execute ("set dateformat dmy Insert into Contrib(memberno,contrdate,refno,Amount,sharebal,transby,ChequeNo,receiptno,remarks,auditid,sharescode,transactionno)" _
''                                        & " Values('" & mMemberNo & "','" & Format(dtpTransDate, "DD/MM/YYYY") & "',0," & Amount * (-1) & "," & totalshares & ",'','','" & ReceiptNo & "','Reversal -" & ReceiptNo & "' ,'" & User & "','" & sharesCode & "','" & transactionNo & "')")
''
''                                        If success = False Then
''                                            GoTo UndoTrans:
''                                        End If
''                                        .MoveNext
''                                    Wend
''                                Else
''
''                                End If
''                            End With
''
''Jump:
                            
                            
                            'Reverse cheque Deposits
''                            Set rst = oSaccoMaster.GetRecordset("select * from  ChequeDeposits where transactionno='" & transactionNo & "'")
''                            If Not rst.EOF Then
''                              While Not rst.EOF
''                                'Description = rst!TransDescription
''                                BankAcc = IIf(IsNull(rst!bank), "", rst!bank)
''                                vno = rst!vno
''                                customerAcc = rst!AccNo
''                                Amount = rst!Amount
''                                chequeNo = rst!chequeNo
''                                Maturitydate = rst!valuedate
''
''                                sql = "insert into chequeDeposits(AccNo,DateReceived,Amount,VNo,CType,ChequeNo,Bank,AuditId,valueDate,drawer,transactionNo)" _
''                                & "Values('" & customerAcc & "','" & Format(dtpTransDate, "DD/MM/YYYY") & "'," & Amount & ",'" & vno & "','','" & chequeNo & "','" & BankAcc & "','" & User & "','" & Format(dtpTransDate, "dd/mm/yyyy") & "','','" & transactionNo & "')"
''                               oSaccoMaster.Execute (sql)
''                               rst.MoveNext
''                              Wend
''                            End If
                       
                            'Reverse the transaction from gltransactions table
''                            Dim Source As String, DRAcc As String, CRAcc As String, DocumentNo As String
''                            Set rs = oSaccoMaster.GetRecordset("select distinct Amount,DrAccNo,CrAccNo,DocumentNo,SOURCE,TransDescript,chequeNo from gltransactions where transactionNo='" & transactionNo & "' ")
''                            If Not rs.EOF Then
''                                While Not rs.EOF
''                                    DRAcc = IIf(IsNull(rs("DrAccNo")), "", rs("DrAccNo"))
''                                    CRAcc = IIf(IsNull(rs("CrAccNo")), "", rs("CrAccNo"))
''                                    Amount = IIf(IsNull(rs("Amount")), "", rs("Amount"))
''                                    DocumentNo = IIf(IsNull(rs("Documentno")), "", rs("Documentno"))
''                                    description = IIf(IsNull(rs("TransDescript")), "", rs("TransDescript"))
''                                    Source = IIf(IsNull(rs("Source")), "", rs("Source"))
''                                    chequeno = IIf(IsNull(rs("chequeNo")), "", rs("chequeNo"))
''
''                                    'switch the accounts
''                                    TempAccNo = DRAcc
''                                    DRAcc = CRAcc
''                                    CRAcc = TempAccNo
''
''                                    If Not Save_GLTRANSACTION(dtpTransDate, Amount, DRAcc, CRAcc, DocumentNo, Source, User, "", "Reversal- " & DocumentNo & "-" & description, 0, 1, chequeno, transactionNo, "") Then
''                                        GoTo UndoTrans
''                                    End If
''
''                                      'Check if loan issued
''                                        If Trim(UCase(description)) = "LOAN ISSUED" Then
''                                             sql = "SELECT loanNo   FROM LOANBAL  WHERE (LoanNo = '" & DocumentNo & "')"
''                                            Set RsLoans = oSaccoMaster.GetRecordset(sql)
''                                            If Not RsLoans.EOF Then
''                                                 LoanNo = RsLoans.Fields(0)
''                                                'delete from loans
''                                                oSaccoMaster.ExecuteThis ("Delete from loans where loanno='" & LoanNo & "'")
''                                                  'delete from appraisal
''                                                oSaccoMaster.ExecuteThis ("Delete from Appraisal where loanno='" & LoanNo & "'")
''
''                                                ' delete from Endmain
''                                                oSaccoMaster.ExecuteThis ("Delete from EndMAIN where loanno='" & LoanNo & "'")
''
''                                                'Delete from DisbursementDeduction
''
''                                                oSaccoMaster.ExecuteThis ("Delete from DisbursementDeduction where loanno='" & LoanNo & "'")
''
''                                                'Delete from Cheques
''                                                oSaccoMaster.ExecuteThis ("Delete from Cheques where loanno='" & LoanNo & "'")
''                                                'Delete from Cheques
''                                                oSaccoMaster.ExecuteThis ("Delete from LOANBAL where loanno='" & LoanNo & "'")
''                                            End If
''                                        End If
''                                    rs.MoveNext
''                                Wend
''                            End If
''
''                            'Update Transaction as Reversed
''                            oSaccoMaster.ExecuteThis ("Update Transactions set Status='REVERSED',reason='" & txtReason & "',Deletedby='" & Current_User.UserId & "' where transactionno='" & transactionNo & "'")
''                            If success = False Then
''                                GoTo UndoTrans
''                            End If
''                        End If
''                    Next I
''                Else
''                Exit Sub
''                End If
''            End If
''        End With
''        .CommitTrans
''        MsgBox "Transaction  Complete!"
''        chkSorted_Click
''        txtReason.Text = ""
''    Exit Sub
''UndoTrans:
''        .RollbackTrans
''        MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
''End With
End Sub

Private Sub cmdWReverse_Click()
    On Error GoTo Capture
    Dim message As String
    Exit Sub
Capture:
    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
End Sub


Private Sub LoadTransactions()
    On Error GoTo Capture
    Dim trs As Double

    If chkSorted.Value = 1 Then 'sorted
        If cboUser <> "" And txtDocumentno <> "" Then
            sql = "set dateformat dmy SELECT T.*,GT.DocumentNo FROM Transactions T "
           sql = sql & " RIGHT OUTER JOIN "
           sql = sql & " (SELECT G.TransactionNo,G.DocumentNo FROM GLTRANSACTIONS G WHERE G.DocumentNo='" & txtDocumentno & "') GT ON GT.TransactionNo=T.TransactionNo "
           sql = sql & "  WHERE  T.status='Active' and t.auditId='" & cboUser & "' AND T.transDescription not like'%Milk Advance%' ORDER BY T.AuditTime"
           
'            sql = "set dateformat dmy select tr.*,gl.DocumentNo from transactions tr inner join Gltransactions gl on tr.transactionNo=gl.transactionNo  where tr.TransDate ='" & Format(dtpTransDate, "DD/MM/YYYY") _
'            & "' and GL.DocumentNo='" & txtDocumentno & "' and tr.auditId='" & cboUser & "' and  tr.status='Active' and tr.transDescription not like'%Milk Advance%'   order by tr.audittime"
        ElseIf cboUser <> "" And txtDocumentno = "" Then
           sql = "set dateformat dmy SELECT T.*,GT.DocumentNo FROM Transactions T "
           sql = sql & " RIGHT OUTER JOIN "
           sql = sql & " (SELECT G.TransactionNo,G.DocumentNo FROM GLTRANSACTIONS G ) GT ON GT.TransactionNo=T.TransactionNo "
           sql = sql & "  WHERE  T.transdate='" & Format(DTPtransdate, "DD/MM/YYYY") & "' AND T.status='Active' and t.auditId='" & cboUser & "' AND T.transDescription not like'%Milk Advance%' ORDER BY T.AuditTime"

'         sql = "set dateformat dmy select tr.*,gl.DocumentNo from transactions tr inner join Gltransactions gl on tr.transactionNo=gl.transactionNo  where tr.TransDate ='" & Format(dtpTransDate, "DD/MM/YYYY") _
'            & "'  and tr.auditId='" & cboUser & "' and  tr.status='Active' and tr.transDescription not like'%Milk Advance%'   order by tr.audittime"
        ElseIf cboUser = "" And txtDocumentno <> "" Then
           sql = "set dateformat dmy SELECT T.*,GT.DocumentNo FROM Transactions T "
           sql = sql & " RIGHT OUTER JOIN "
           sql = sql & " (SELECT G.TransactionNo,G.DocumentNo FROM GLTRANSACTIONS G WHERE G.DocumentNo='" & txtDocumentno & "') GT ON GT.TransactionNo=T.TransactionNo "
           sql = sql & "  WHERE  T.status='Active' AND T.transDescription not like'%Milk Advance%' ORDER BY T.AuditTime"
           
'          sql = "set dateformat dmy select tr.*,gl.DocumentNo from transactions tr inner join Gltransactions gl on tr.transactionNo=gl.transactionNo  where tr.TransDate ='" & Format(dtpTransDate, "DD/MM/YYYY") _
'          & "' and GL.DocumentNo='" & txtDocumentno & "'  and  tr.status='Active' or GL.DocumentNo='" & txtDocumentno & "'  and  tr.status='Active'  and tr.transDescription not like'%Milk Advance%'  order by tr.audittime"
        Else
        sql = ""
'          sql = "set dateformat dmy SELECT T.*,GT.DocumentNo FROM Transactions T "
'           sql = sql & " RIGHT OUTER JOIN "
'           sql = sql & " (SELECT G.TransactionNo,G.DocumentNo FROM GLTRANSACTIONS G) GT ON GT.TransactionNo=T.TransactionNo "
'           sql = sql & "  WHERE  T.status='Active' AND T.transdate='" & Format(dtpTransDate, "DD/MM/YYYY") & "' AND T.transDescription not like'%Milk Advance%' ORDER BY T.AuditTime"

        End If
    Else
        sql = "set dateformat dmy SELECT T.*,GT.DocumentNo FROM Transactions T "
           sql = sql & " RIGHT OUTER JOIN "
           sql = sql & " (SELECT G.TransactionNo,G.DocumentNo FROM GLTRANSACTIONS G) GT ON GT.TransactionNo=T.TransactionNo "
           sql = sql & "  WHERE  T.status='Active' AND T.transdate='" & Format(DTPtransdate, "DD/MM/YYYY") & "' AND T.transDescription not like'%Milk Advance%' ORDER BY T.AuditTime"

'          sql = "set dateformat dmy select tr.*,isnull(doc.documentno,'Internal Trans')documentno  from transactions tr left outer join " _
'        & " (SELECT DISTINCT TransactionNo, DocumentNo FROM dbo.GLTRANSACTIONS GROUP BY TransactionNo, DocumentNo)Doc on Doc.transactionno=tr.transactionno where tr.transdate='" & _
'        Format(dtpTransDate, "DD/MM/YYYY") & "' AND tr.status='Active' and tr.transDescription not like'%Milk Advance%' order by tr.audittime desc"

    End If
    If sql <> "" Then
    Set Rst5 = oSaccoMaster.GetRecordset(sql)
    Else
    Exit Sub
    End If
    With Rst5
       If Not .EOF Then
           lvwTransactions.ListItems.Clear
           While Not .EOF
               Set li = lvwTransactions.ListItems.Add(, , !transactionNo)
               li.ListSubItems.Add , , !transdate
               li.ListSubItems.Add , , !Amount
               li.ListSubItems.Add , , !DocumentNo
               li.ListSubItems.Add , , !auditid
               li.ListSubItems.Add , , !audittime & ""
               li.ListSubItems.Add , , !TransDescription
               
               
           .MoveNext
           Wend
       Else
           lvwTransactions.ListItems.Clear
       End If
    End With
    
    Exit Sub
Capture:
    MsgBox ErrorMessage
End Sub
Sub loadWithdrawal()
    On Error GoTo Capture
    'Withdrawn Members
    lvwWithdrawn.ListItems.Clear
    sql = "select memberno,names,bookingdate,withdrawdate,transactionno from vwwithdrawnmembers"
    Set Rst = oSaccoMaster.GetRecordset(sql)
    With Rst
        While Not .EOF
            Set li = lvwWithdrawn.ListItems.Add(, , !memberno)
            li.ListSubItems.Add , , !NAMES
            li.ListSubItems.Add , , !bookingDate
            li.ListSubItems.Add , , !withdrawdate
            li.ListSubItems.Add , , !transactionNo
        .MoveNext
        Wend
    End With
    Exit Sub
    
Capture:
    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
End Sub





Private Sub Form_Load()
        
        'loadWithdrawal
        dtpCurrentDate = Get_Server_Date
        DTPtransdate = dtpCurrentDate
        Set rs = oSaccoMaster.GetRecordset("select distinct userloginid from useraccounts")
        While Not rs.EOF
            cboUser.AddItem rs(0)
            rs.MoveNext
        Wend
        'LoadTransactions
End Sub

Private Sub lvwTransactions_DblClick()
Dim rsrepay As New Recordset, Amnt As Double, Amnt1 As Double, tdate As Date
On Error GoTo UndoTrans

    If MsgBox("Do you Want To change Transaction Date?", vbQuestion + vbYesNo) = vbYes Then
      tdate = InputBox("Enter the Transaction", "Date", lvwTransactions.SelectedItem.ListSubItems(1))
        If Not IsDate(tdate) Then
            MsgBox "Enter Valid Date", vbInformation, Me.Caption
            Exit Sub
        End If
           lvwTransactions.SelectedItem.ListSubItems(1) = tdate
        Else
        tdate = lvwTransactions.SelectedItem.ListSubItems(1)
    End If
    
    If MsgBox("Do you want to Edit the Transaction Amount?", vbQuestion + vbYesNo) = vbYes Then
        Amnt1 = lvwTransactions.SelectedItem.ListSubItems(2)
        Amnt = InputBox("Enter the amount", "AMOUNT", lvwTransactions.SelectedItem.ListSubItems(2))
        lvwTransactions.SelectedItem.ListSubItems(2) = Amnt
    Else
    Amnt = lvwTransactions.SelectedItem.ListSubItems(2)
    End If
    
   If MsgBox("Are you sure you want to Edit this  Transaction", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    oSaccoMaster.ConnectDatabase
     With oSaccoMaster.goConn
       .BeginTrans
        With lvwTransactions
         If Not lvwTransactions.ListItems.Count <= 0 Then
            For I = 1 To .ListItems.Count
            If .ListItems(I).Checked = True Then
            TransNo = .ListItems(I).Text
            transactionNo = TransNo
            Amnt = .ListItems(I).SubItems(2)
            tdate = Format(tdate, "dd/mm/yyyy")
           ' ******/Update Pettycash  **********/
               sql = " set dateformat dmy select Distinct TransactionNo from PettyCash where transactionno='" & transactionNo & "'"
                Set rsrepay = oSaccoMaster.GetRecordset(sql)
                    If Not rsrepay.EOF Then
                     oSaccoMaster.ExecuteThis ("SET dateformat dmy Update PettyCash set transdate='" & tdate & "',Amount=" & Amnt & " where transactionno='" & transactionNo & "'")

                    End If
                    ' ******* Update Bank Accounts **********/
                     oSaccoMaster.ExecuteThis ("SET dateformat dmy Update BankAccount set transdate='" & tdate & "',Amount=" & Amnt & " where transactionno='" & transactionNo & "'")
                        

                    ' ******/Update Gltransaction  **********/
                    sql = " set dateformat dmy select Distinct TransactionNo from Gltransactions where transactionno='" & transactionNo & "'"
                   Set rs2 = oSaccoMaster.GetRecordset(sql)
                    If Not rs2.EOF Then
                     oSaccoMaster.ExecuteThis ("SET dateformat dmy Update Gltransactions set Transdate='" & tdate & "',Amount=" & Amnt & " where transactionno='" & transactionNo & "'")
                      
                    End If

                    ' ******/Update transaction Transdate **********/
                      sql = " set dateformat dmy select Distinct TransactionNo from Transactions where transactionno='" & transactionNo & "'"
                   Set rs = oSaccoMaster.GetRecordset(sql)
                    If Not rs.EOF Then
                     oSaccoMaster.ExecuteThis ("SET dateformat dmy Update transactions set Transdate='" & tdate & "',Amount=" & Amnt & ",Amount2=" & Amnt1 & ",ModifiedBy='" & User & "' where transactionno='" & transactionNo & "'")
                      
                    End If
                    
            End If
            Next I
         End If
        End With
        .CommitTrans
 MsgBox " Transactions Succefully Edited", vbInformation
Exit Sub
UndoTrans:
   .RollbackTrans
    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
 End With
End Sub

Private Sub mnushareimport_Click()
frmsharesupdate.Show vbModal
End Sub

Private Sub txtDocumentno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     LoadTransactions
   End If
End Sub
