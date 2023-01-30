VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBudgets 
   Caption         =   "Budgets"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10140
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBudgets.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdback 
      Caption         =   "&Back"
      Height          =   390
      Left            =   4200
      TabIndex        =   18
      Top             =   5880
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdview 
      Caption         =   "&View Ubudgetted Acc"
      Height          =   510
      Left            =   2760
      TabIndex        =   17
      Top             =   5760
      Width           =   1350
   End
   Begin MSComctlLib.ListView lvwaccounts 
      Height          =   6015
      Left            =   0
      TabIndex        =   16
      Top             =   -120
      Visible         =   0   'False
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   10610
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "AccNo"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "GlAccName"
         Object.Width           =   17639
      EndProperty
   End
   Begin VB.CheckBox Chkselect 
      Caption         =   "Select All"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   975
   End
   Begin MSComctlLib.ListView Lvwmonth 
      Height          =   3975
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7011
      View            =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComCtl2.DTPicker dtpBudget 
      Height          =   330
      Left            =   5865
      TabIndex        =   10
      Top             =   390
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   " dd/mm/yyyy"
      Format          =   147718145
      UpDown          =   -1  'True
      CurrentDate     =   39307
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   390
      Left            =   5640
      TabIndex        =   9
      Top             =   5835
      Width           =   1230
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   390
      Left            =   7200
      TabIndex        =   8
      Top             =   5835
      Width           =   1230
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
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
      Left            =   1860
      TabIndex        =   4
      Top             =   1035
      Width           =   1560
   End
   Begin VB.TextBox txtAccName 
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
      Left            =   1710
      TabIndex        =   3
      Top             =   390
      Width           =   4020
   End
   Begin VB.TextBox txtAccNo 
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
      Left            =   105
      TabIndex        =   2
      Top             =   390
      Width           =   1290
   End
   Begin VB.CommandButton cmdSearch 
      Height          =   330
      Left            =   1410
      Picture         =   "frmBudgets.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   405
      Width           =   300
   End
   Begin MSComctlLib.ListView lvwBudgets 
      Height          =   3975
      Left            =   4410
      TabIndex        =   0
      Top             =   1785
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "TransDate"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Amount"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label lblbudget 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4680
      TabIndex        =   22
      Top             =   1155
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Budget"
      Height          =   255
      Left            =   4935
      TabIndex        =   21
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblactuals 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6240
      TabIndex        =   20
      Top             =   1155
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Actuals"
      Height          =   255
      Left            =   6495
      TabIndex        =   19
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Variance"
      Height          =   255
      Left            =   8280
      TabIndex        =   13
      Top             =   885
      Width           =   735
   End
   Begin VB.Label lblvariance 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8025
      TabIndex        =   12
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Budget Year"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5820
      TabIndex        =   11
      Top             =   135
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Account Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1755
      TabIndex        =   7
      Top             =   120
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Account No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Budgetted Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   195
      TabIndex        =   5
      Top             =   1020
      Width           =   1515
   End
   Begin VB.Menu mnubudgetreporst 
      Caption         =   "Budget Reports"
      Begin VB.Menu mnubudgetvariance 
         Caption         =   "Budget Variance Income Statement"
      End
      Begin VB.Menu mnubudgetvariancebs 
         Caption         =   "Budget Variance Balance Sheet"
      End
   End
End
Attribute VB_Name = "frmBudgets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Chkselect_Click()
 lvwBudgets.ListItems.Clear
 
    If (txtAmount = "" Or CStr(txtAmount) = "0") Then
      MsgBox "Please Enter the Budget Amount", vbInformation, "Budget"
      txtAmount.SetFocus
      Exit Sub
    End If
     If Trim$(txtAccNo) = "" Then
        MsgBox "Please supply the Account No", vbInformation, Me.Caption
        txtAccNo.SetFocus
        Exit Sub
    End If
 For I = 1 To 12
 Set li = Lvwmonth.ListItems(I)
 If Chkselect.Value = vbChecked Then
 li.Checked = True
   Get_Budgets year(dtpBudget), txtAmount
    txtAmount = Format(txtAmount, Cfmt)
 Else
 li.Checked = False
 End If
Next I
End Sub


Private Sub cmdBack_Click()
lvwaccounts.Visible = False
Label5.Visible = True
cmdNew.Visible = True
cmdSave.Visible = True
cmdback.Visible = False
End Sub

Private Sub cmdnew_Click()
lvwBudgets.ListItems.Clear
txtAmount = 0
txtAccNo = ""
txtAccName = ""
End Sub

Private Sub cmdsave_Click()
  'SELECT     Accno, mmonth, yyear, Actual, Budgetted, Variance  FROM         budgets
    Save_My_Budget
    MsgBox "Budget Updated Successfully", vbInformation, Me.Caption
    txtAccNo = ""
    txtAccName = ""
    txtAmount = 0
    lvwBudgets.ListItems.Clear
    
    Exit Sub
   
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo SysError
    frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtAccNo = SearchValue
            SearchValue = ""
        End If
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub





Private Sub cmdview_Click()
lvwaccounts.Visible = True
Label5.Visible = False
cmdback.Visible = True
cmdNew.Visible = False
cmdSave.Visible = False
lvwaccounts.ListItems.Clear
Set Rst1 = Get_Records("set dateformat dmy select distinct  AccNo,GlAccName from GLSETUP" _
                & " where AccNo Not in ( select AccNo from budgets where yyear=year('" & dtpBudget & "'))order by AccNo", ErrorMessage)
  With Rst1
    While Not .EOF
    Set li = lvwaccounts.ListItems.Add(, , !AccNo)
         li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
      .MoveNext
    Wend
  End With
End Sub

Private Sub Form_Load()
    dtpBudget = Format(Get_Server_Date, " dd-MM-yyyy")
    
End Sub



Private Sub lvwAccounts_DblClick()
If lvwaccounts.ListItems.Count > 0 Then
txtAccNo = lvwaccounts.SelectedItem.Text
cmdBack_Click
End If
End Sub

Private Sub Lvwmonth_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    lvwBudgets.ListItems.Clear
    If (txtAmount = "" Or CStr(txtAmount) = "0") Then
      MsgBox "Please Enter the Budget Amount", vbInformation, "Budget"
      txtAmount.SetFocus
      Exit Sub
    End If
     If Trim$(txtAccNo) = "" Then
        MsgBox "Please supply the Account No", vbInformation, Me.Caption
        txtAccNo.SetFocus
        Exit Sub
    End If
    Get_Budgets year(dtpBudget), txtAmount
    txtAmount = Format(txtAmount, Cfmt)
    
End Sub

Private Sub mnubudgetvariance_Click()
Dim Variance As Double, remarks As String
transdate = Format(dtpBudget, "dd/mm/yyyy")
UsedAmount = 0
BudgettedAmount = 0
Variance = 0
Set Rst3 = Get_Records("Truncate table BudgetVariance", ErrorMessage)
Set Rst1 = Get_Records(" set dateformat dmy Select distinct Accno,isnull(SUM(Budgetted),0) as Budgetted from Budgets where Yyear= '" & year(transdate) & "' and mmonth<= '" & month(transdate) & "' Group by AccNo", ErrorMessage)
 With Rst1
  While Not .EOF
    Set Rst4 = oSaccoMaster.GetRecordset("select Normalbal,GlAccType from Glsetup where Accno='" & Rst1!AccNo & "'")
        If Rst4.Fields(0) = "Debit" Then
             sql = "set dateformat dmy Select(SELECT (SELECT isnull(SUM(Amount),0) FROM  GLTRANSACTIONS WHERE DrAccNo = '" & Rst1!AccNo & "' and MONTH(transdate)<=MONTH('" & transdate & "') and YEAR(transdate)=YEAR('" & transdate & "'))- (SELECT isnull(SUM(Amount),0) FROM  GLTRANSACTIONS WHERE CrAccNo = '" & Rst1!AccNo & "' and MONTH(transdate)<=MONTH('" & transdate & "') and YEAR(transdate)=YEAR('" & transdate & "'))) AS Amount"
            Set Rst5 = oSaccoMaster.GetRecordset(sql)
        ElseIf Rst4!NormalBal = "Credit" Then
           sql = "set dateformat dmy Select(SELECT (SELECT isnull(SUM(Amount),0) FROM  GLTRANSACTIONS WHERE CrAccNo = '" & Rst1!AccNo & "' and MONTH(transdate)<=MONTH('" & transdate & "') and YEAR(transdate)=YEAR('" & transdate & "'))- (SELECT isnull(SUM(Amount),0) FROM  GLTRANSACTIONS WHERE DrAccNo = '" & Rst1!AccNo & "' and MONTH(transdate)<=MONTH('" & transdate & "') and YEAR(transdate)=YEAR('" & transdate & "'))) AS Amount"
            Set Rst5 = oSaccoMaster.GetRecordset(sql)
        End If
         UsedAmount = IIf(IsNull(Rst5!Amount), 0, Rst5!Amount)
         BudgettedAmount = IIf(IsNull(Rst1!Budgetted), 0, Rst1!Budgetted)
         If UCase(Rst4!Glacctype) = "INCOME STATEMENT" And Rst4!NormalBal = "Credit" Then
         Variance = CDbl(UsedAmount - BudgettedAmount)
         Else
         Variance = CDbl(BudgettedAmount - UsedAmount)
         End If
         If Variance >= 0 Then
         remarks = "Favourable"
         Else
         remarks = "Adverse"
         End If
             If Not Execute_Command("insert into BudgetVariance(Accno, mmonth, yyear,Actual ,Budgetted, Variance,Remarks) values('" & Rst1!AccNo & "'," & month(transdate) & "," _
             & year(transdate) & "," & UsedAmount & "," & BudgettedAmount & "," & Variance & ",'" & remarks & "')", ErrorMessage) Then
                If ErrorMessage <> "" Then
                    MsgBox ErrorMessage, vbInformation, Me.Caption
                    ErrorMessage = ""
                End If
            End If
   .MoveNext
  Wend
 End With
         reportname = "Budgetvariance.rpt"
          STRFORMULA = ""
          Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
          
 
End Sub

Private Sub mnubudgetvariancebs_Click()
Set rsDebits = oSaccoMaster.GetRecordset("select * from Budgetvariance")
If Not rsDebits.EOF Then
         reportname = "BudgetvarianceBal.rpt"
          STRFORMULA = ""
          Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
 Else
 MsgBox "Process Income and Expense Budget variance first", vbInformation, "Budget variance"
 Exit Sub
 End If

End Sub

Private Sub txtAccNo_Change()
    On Error GoTo SysError
    Dim Variance As Double
    Lvwmonth.ListItems.Clear
    lvwBudgets.ListItems.Clear
    txtAmount = 0
    lblactuals = 0
    lblbudget = 0
    lblvariance = 0
    If Trim$(txtAccNo) <> "" Then
        Dim Account  As String
         Get_GL_AccDetails (txtAccNo)
        If GlAccName <> "" Then
            txtAccName = GlAccName
            For I = 1 To 12
            Set li = Lvwmonth.ListItems.Add(, , MonthName(I))
             Next I
        Else
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
    End If
    If GlAccName <> "" Then
      Set Rst3 = oSaccoMaster.GetRecordset("Set dateformat dmy select * from Budgets where AccNo='" & txtAccNo & "' and  yyear=YEAR('" & dtpBudget & "') order by mmonth")
      If Not Rst3.EOF Then
         lvwBudgets.ListItems.Clear
         With Rst3
         While Not .EOF
           Set li = lvwBudgets.ListItems.Add(, , IIf(IsNull(!mMonth), "", MonthName(!mMonth)))
               li.SubItems(1) = IIf(IsNull(!Budgetted), "", !Budgetted)
            .MoveNext
         Wend
        End With
      End If
      Set Rst1 = oSaccoMaster.GetRecordset("Set dateformat dmy select ISNULL(SUM(Budgetted),0) as AMOUNT from Budgets where AccNo='" & txtAccNo & "' and  yyear=YEAR('" & dtpBudget & "') ")
      If Not Rst1.EOF Then
      txtAmount = CDbl(Rst1.Fields(0))
      End If
      transdate = Format(dtpBudget, "dd/mm/yyyy")
        UsedAmount = 0
        BudgettedAmount = 0
        Variance = 0
    Set Rst1 = Get_Records(" set dateformat dmy Select  Accno,isnull(SUM(Budgetted),0) as Budgetted from Budgets where Accno='" & txtAccNo & "'and Yyear= '" & year(transdate) & "' and mmonth<= '" & month(transdate) & "' Group by AccNo", ErrorMessage)
     With Rst1
      If Not .EOF Then
        Set Rst4 = oSaccoMaster.GetRecordset("select Normalbal,GlAccType from Glsetup where Accno='" & Rst1!AccNo & "'")
           If Rst4.Fields(0) = "Debit" Then
             sql = "set dateformat dmy Select(SELECT (SELECT isnull(SUM(Amount),0) FROM  GLTRANSACTIONS WHERE DrAccNo = '" & Rst1!AccNo & "' and MONTH(transdate)<=MONTH('" & transdate & "') and YEAR(transdate)=YEAR('" & transdate & "'))- (SELECT isnull(SUM(Amount),0) FROM  GLTRANSACTIONS WHERE CrAccNo = '" & Rst1!AccNo & "' and MONTH(transdate)=MONTH('" & transdate & "') and YEAR(transdate)=YEAR('" & transdate & "'))) AS Amount"
            Set Rst5 = oSaccoMaster.GetRecordset(sql)
           
           ElseIf Rst4!NormalBal = "Credit" Then
            sql = "set dateformat dmy Select(SELECT (SELECT isnull(SUM(Amount),0) FROM  GLTRANSACTIONS WHERE CrAccNo = '" & Rst1!AccNo & "' and MONTH(transdate)<=MONTH('" & transdate & "') and YEAR(transdate)=YEAR('" & transdate & "'))- (SELECT isnull(SUM(Amount),0) FROM  GLTRANSACTIONS WHERE DrAccNo = '" & Rst1!AccNo & "' and MONTH(transdate)=MONTH('" & transdate & "') and YEAR(transdate)=YEAR('" & transdate & "'))) AS Amount"
            Set Rst5 = oSaccoMaster.GetRecordset(sql)
                
            End If
             UsedAmount = IIf(IsNull(Rst5!Amount), 0, Rst5!Amount)
             BudgettedAmount = IIf(IsNull(Rst1!Budgetted), 0, Rst1!Budgetted)
             If UCase(Rst4!Glacctype) = "INCOME STATEMENT" And Rst4!NormalBal = "Credit" Then
             Variance = CDbl(UsedAmount - BudgettedAmount)
             Else
             Variance = CDbl(BudgettedAmount - UsedAmount)
             End If
             lblvariance = CDbl(Variance)
             lblactuals = CDbl(UsedAmount)
             lblbudget = CDbl(BudgettedAmount)
      Else
      lblvariance = 0
      lblbudget = 0
      
      End If
     End With
      
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Get_Budgets(mYear As Long, Amount As Double)
    Dim I As Long, transdate As Date, j As Long, monthnam As String, Myrecs As Long
    On Error GoTo SysError
    lvwBudgets.ListItems.Clear
              Myrecs = 0
               For j = 1 To 12
                Set li = Lvwmonth.ListItems(j)
                
                If li.Checked Then
                Myrecs = Myrecs + 1
                End If
                Next j
        For I = 1 To 12
        If Lvwmonth.ListItems(I).Checked Then
        transdate = DateSerial(mYear, I + 1, 1 - 1)
        monthnam = MonthName(I)
        Set li = lvwBudgets.ListItems.Add(, , monthnam)
        li.SubItems(1) = Format(Amount / Myrecs, Cfmt)
        End If
    Next I
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = To_Upper_Case(KeyAscii)
End Sub

Private Sub txtAmount_LostFocus()
    On Error GoTo SysError
    If Trim$(txtAmount) <> "" Then
        Get_Budgets year(dtpBudget), txtAmount
        txtAmount = Format(txtAmount, Cfmt)
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub
Private Sub Save_My_Budget()
    On Error GoTo SysError
    Dim I As Long, rsBudget As New Recordset, j As Long, Myrecs As Long
    Set rsBudget = oSaccoMaster.GetRecordset("Select AccNo From GLSETUP where AccNo='" _
    & txtAccNo & "'")
    With rsBudget
        If .State = adStateOpen Then
            If Not .EOF Then
                Set rsBudget = oSaccoMaster.GetRecordset("Delete From BUDGETS where AccNo='" & txtAccNo & "' and yYear='" & year(dtpBudget) & "' ")
                Myrecs = 0
                For j = 1 To 12
                Set li = Lvwmonth.ListItems(j)
                
                If li.Checked Then
                Myrecs = Myrecs + 1
                End If
                Next j
                
                For I = 1 To 12
                    Set li = Lvwmonth.ListItems(I)
                    If li.Checked Then
                    If Not Save_The_Budget(txtAccNo, I, year(dtpBudget), _
                    CDbl(txtAmount) / Myrecs, ErrorMessage) Then
                        If ErrorMessage <> "" Then
                            MsgBox ErrorMessage, vbInformation, Me.Caption
                            ErrorMessage = ""
                        End If
                    End If
                    End If
                Next I
            Else
                MsgBox "Account No " & txtAccNo & " not found in the Chart Of Accounts", vbInformation, Me.Caption
                txtAccNo.SetFocus
                SendKeys "{Home}+{End}"
                Exit Sub
            End If
        End If
    End With
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

