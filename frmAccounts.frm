VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAccounts 
   BackColor       =   &H00C0C000&
   Caption         =   "Generate Trial Balance"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12435
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   12435
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdprintratio 
      Caption         =   "Print Financial Ratio"
      Height          =   375
      Left            =   6000
      TabIndex        =   27
      Top             =   2880
      Width           =   1785
   End
   Begin VB.CommandButton cmdfinancialR 
      Caption         =   "Process Financial Ratios"
      Height          =   375
      Left            =   3840
      TabIndex        =   26
      Top             =   2880
      Width           =   1905
   End
   Begin VB.CommandButton cmdprocessexp 
      Caption         =   "Process Expense"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton cmdprintexp 
      Caption         =   "Print Expense"
      Height          =   375
      Left            =   2160
      TabIndex        =   24
      Top             =   2880
      Width           =   1545
   End
   Begin VB.CommandButton cmdcashflow 
      Caption         =   "Cash Flow"
      Height          =   375
      Left            =   7440
      TabIndex        =   22
      Top             =   2160
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF00FF&
      Caption         =   "MONTHLY INCOME"
      Height          =   1815
      Left            =   7680
      TabIndex        =   14
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton cmdmonthlyincome 
         Caption         =   "Monthly Income Statement"
         Height          =   375
         Left            =   1920
         TabIndex        =   16
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox cbpdpt 
         Height          =   330
         ItemData        =   "frmAccounts.frx":0000
         Left            =   120
         List            =   "frmAccounts.frx":0013
         TabIndex        =   15
         Top             =   960
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker Startdate99 
         Height          =   375
         Left            =   2640
         TabIndex        =   17
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   132448257
         CurrentDate     =   42280
      End
      Begin MSComCtl2.DTPicker Enddate99 
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   132448257
         CurrentDate     =   42308
      End
      Begin VB.Label Label7 
         Caption         =   "TO  Date"
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "From Date"
         Height          =   375
         Left            =   1800
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Activity"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.ComboBox cbodepartment 
      Height          =   330
      ItemData        =   "frmAccounts.frx":0041
      Left            =   4680
      List            =   "frmAccounts.frx":0054
      TabIndex        =   12
      Top             =   480
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar prgStatus 
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdbalancesheet 
      Caption         =   "Print Balance Sheet"
      Height          =   375
      Left            =   5340
      TabIndex        =   8
      Top             =   2205
      Width           =   1935
   End
   Begin VB.CommandButton cmdincomestmnt 
      Caption         =   "Print Income Statement"
      Height          =   375
      Left            =   3300
      TabIndex        =   7
      Top             =   2205
      Width           =   1935
   End
   Begin VB.CommandButton cmdTrialbalance 
      Caption         =   "Print TB"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   2160
      Width           =   1545
   End
   Begin MSComCtl2.DTPicker dtpFinishDate 
      Height          =   345
      Left            =   2250
      TabIndex        =   2
      Top             =   510
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   609
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
      CustomFormat    =   " dd-MM-yyyy"
      Format          =   132448259
      CurrentDate     =   39705
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   360
      Left            =   690
      TabIndex        =   1
      Top             =   510
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   635
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
      CustomFormat    =   " dd-MM-yyyy"
      Format          =   132448259
      CurrentDate     =   39705
   End
   Begin VB.CommandButton cmdprocess 
      Caption         =   "Process"
      Height          =   375
      Left            =   195
      TabIndex        =   0
      Top             =   2205
      Width           =   1245
   End
   Begin VB.Label lblaccname 
      BackColor       =   &H00FFFF80&
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
      Left            =   2280
      TabIndex        =   23
      Top             =   1440
      Width           =   4935
   End
   Begin VB.Label Label4 
      Caption         =   "ACTIVITY"
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblAccount 
      BackColor       =   &H00FFFF80&
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00FFFF80&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Finish Date"
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
      Left            =   2280
      TabIndex        =   5
      Top             =   285
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Start Date"
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
      Left            =   720
      TabIndex        =   4
      Top             =   270
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5040
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   645
   End
End
Attribute VB_Name = "frmAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Long
Dim g As Boolean


Private Sub cmdEOY_Click()
Call EOY_Processing(dtpFinishDate)
End Sub

Private Sub cmdbalancesheet_Click()
  '//kimberbalancesheet
   ' reportname = "BalanceSheeet.rpt"
    reportname = "BalanceSheet.rpt"
    STRFORMULA = ""
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
Exit Sub
'IF YOU WANT TO DO A CSV FILE
cmdprocess_Click
End Sub

Private Sub cmdcashflow_Click()
  Dim rscflow As New Recordset
  Dim rsin2 As New Recordset
  Dim rsin3 As New Recordset
  Dim rsin4 As New Recordset
  Dim rsg As New Recordset
  Dim ACTYP As String, Amount2 As Double
  Dim scluster As String, AccNo As String
  Dim nb As String, Amount As Double
  Dim pYearEndDate As Date, CyearEndDate As Date
  Dim Depreciation As Double, Netincome As Double, NetOpbTax As Double, Amotisation, RunBal As Double
 If MsgBox("Requires Processing of Trialbalance first" & "Do You Wish To Continue?", vbQuestion + vbYesNo) = vbNo Then
    Exit Sub
 End If

 sql = ""
 sql = "select * from cashflows where yy1=" & year(dtpFinishDate) & ""
 Set rscflow = oSaccoMaster.GetRecordset(sql)
 If rscflow.EOF Then
    sql = ""
    sql = "SELECT DISTINCT csid, class From cashflows where (csid <= 50)ORDER BY csid"
    Set rsin2 = oSaccoMaster.GetRecordset(sql)
            While Not rsin2.EOF
               sql = ""
               sql = "select * from cashflows where class='" & rsin2!Class & "'and csid=" & rsin2!csid & ""
               Set rsin3 = oSaccoMaster.GetRecordset(sql)
               
               sql = ""
               sql = "INSERT INTO cashflows(cluster, class, descr, ddescr, amty2, amty1,yy1)VALUES     (" _
               & rsin3!Cluster & ", '" & rsin3!Class & "', N'" & rsin3!DESCR & "', N'" & rsin3!DDESCR & "', 0, 0," & year(dtpFinishDate) & ")"
               Set rsin4 = oSaccoMaster.GetRecordset(sql)
            rsin2.MoveNext
            Wend

 End If
 
     I = 1
    sql = ""
    sql = "update cashflows set amty2=0,amty1=0 where yy1=" & year(dtpFinishDate) & ""
    oSaccoMaster.GetRecordset (sql)
     Netincome = 0
     Depreciation = 0
     Amotisation = 0
     NetOpbTax = 0
    
     ' "" net income
     sql = ""
    sql = "SELECT (SELECT ISNULL(SUM(T.OBal),0)  FROM tbbalance T INNER JOIN GLSETUP G  ON G.AccNo=T.accno where G.NormalBal='Credit' )-"
    sql = sql & " (SELECT ISNULL(SUM(T.OBal),0)  FROM tbbalance T INNER JOIN GLSETUP G  ON G.AccNo=T.accno where G.NormalBal='DeBit' )"
   Set Rst = oSaccoMaster.GetRecordset(sql)
    
    If Not Rst.EOF Then
       Netincome = IIf(IsNull(Rst.Fields(0)), 0, Rst.Fields(0))
       sql = ""
    sql = "update cashflows set amty1=" & Rst.Fields(0) & " where yy1=" & year(dtpFinishDate) & " and ddescr='Net Income'"
    oSaccoMaster.GetRecordset (sql)
    End If
    
    '"" Depreciation
    sql = "select Obal from tbbalance where Accno='060302'"
    Set Rst = oSaccoMaster.GetRecordset(sql)
        If Not Rst.EOF Then
        Depreciation = IIf(IsNull(Rst.Fields(0)), 0, Rst.Fields(0))
        sql = ""
        sql = "update cashflows set amty1=" & Depreciation & " where yy1=" & year(dtpFinishDate) & " and ddescr='Depreciation'"
        oSaccoMaster.GetRecordset (sql)
        End If
        
        
'            '"" Amotisation
'    sql = "select Obal from tbbalance where Accno='060302'"
'    Set Rst = oSaccoMaster.GetRecordset(sql)
'        If Not Rst.EOF Then
 '          Amotisation = IIf(IsNull(Rst.Fields(0)), 0, Rst.Fields(0))
'        sql = ""
'        sql = "update cashflows set amty1=" & Amotisation & " where yy1=" & year(dtpFinishDate) & " and ddescr='Amotisation'"
'        oSaccoMaster.GetRecordset (sql)
'        End If


 '""" Net Operating Profit Before Working Capital
        NetOpbTax = Netincome + Depreciation + Amotisation
        
                sql = ""
        sql = "update cashflows set amty1=" & NetOpbTax & " where yy1=" & year(dtpFinishDate) & " and ddescr='Net Operating Profit Before Working Capital'"
        oSaccoMaster.GetRecordset (sql)
        
        ' xxx   Net Changes In Working Capital
        'xxx  (Increase)/Decrease in Trade and other receivables
        RunBal = 0
     sql = ""
    sql = "SELECT (SELECT ISNULL(SUM(T.OBal),0)  FROM tbbalance T INNER JOIN GLSETUP G  ON G.AccNo=T.accno where Scluster='Net Changes In Working Capital' and G.Fcluster='(Increase)/Decrease in Trade and other receivables' )"

   Set Rst = oSaccoMaster.GetRecordset(sql)
    
    If Not Rst.EOF Then
       RunBal = RunBal + IIf(IsNull(Rst.Fields(0)), 0, Rst.Fields(0))
       sql = ""
    sql = "update cashflows set amty1=" & Rst.Fields(0) & " where yy1=" & year(dtpFinishDate) & "  and descr='Net Changes In Working Capital'  and ddescr='(Increase)/Decrease in Trade and other receivables'"
    oSaccoMaster.GetRecordset (sql)
    
    End If
    
    'xx (Increase)/Decrease in Trade and other Payables
    sql = ""
    sql = "SELECT (SELECT ISNULL(SUM(T.OBal),0)  FROM tbbalance T INNER JOIN GLSETUP G  ON G.AccNo=T.accno where Scluster='Net Changes In Working Capital' and  G.Fcluster='(Increase)/Decrease in Trade and other Payables' )"

   Set Rst = oSaccoMaster.GetRecordset(sql)
    
    If Not Rst.EOF Then
       RunBal = RunBal + IIf(IsNull(Rst.Fields(0)), 0, Rst.Fields(0))
       sql = ""
    sql = "update cashflows set amty1=" & Rst.Fields(0) & " where yy1=" & year(dtpFinishDate) & " and descr='Net Changes In Working Capital' and ddescr='(Increase)/Decrease in Trade and other Payables'"
    oSaccoMaster.GetRecordset (sql)
    
    End If
    
        'xx (Increase)/(Increase)/Decrease in Provision
    sql = ""
    sql = "SELECT (SELECT ISNULL(SUM(T.OBal),0)  FROM tbbalance T INNER JOIN GLSETUP G  ON G.AccNo=T.accno where Scluster='Net Changes In Working Capital' and  G.Fcluster='(Increase)/Decrease in Provision ')"

   Set Rst = oSaccoMaster.GetRecordset(sql)
    
    If Not Rst.EOF Then
       RunBal = RunBal + IIf(IsNull(Rst.Fields(0)), 0, Rst.Fields(0))
       sql = ""
    sql = "update cashflows set amty1=" & Rst.Fields(0) & " where yy1=" & year(dtpFinishDate) & " and descr='Net Changes In Working Capital' and ddescr='(Increase)/Decrease in Provision'"
    oSaccoMaster.GetRecordset (sql)
    
    End If
    
    
            'xx (Increase)/Decrease in Inventories
    sql = ""
    sql = "SELECT (SELECT ISNULL(SUM(T.OBal),0)  FROM tbbalance T INNER JOIN GLSETUP G  ON G.AccNo=T.accno where Scluster='Net Changes In Working Capital' and  G.Fcluster='(Increase)/Decrease in Inventories ')"

   Set Rst = oSaccoMaster.GetRecordset(sql)
    
    If Not Rst.EOF Then
       RunBal = RunBal + IIf(IsNull(Rst.Fields(0)), 0, Rst.Fields(0))
       sql = ""
    sql = "update cashflows set amty1=" & Rst.Fields(0) & " where yy1=" & year(dtpFinishDate) & " and descr='Net Changes In Working Capital' and ddescr='(Increase)/Decrease in Inventories'"
    oSaccoMaster.GetRecordset (sql)
    
    End If
    
                'xx (Increase)/Decrease in Prepayment
    sql = ""
    sql = "SELECT (SELECT ISNULL(SUM(T.OBal),0)  FROM tbbalance T INNER JOIN GLSETUP G  ON G.AccNo=T.accno where Scluster='Net Changes In Working Capital' and  G.Fcluster='(Increase)/Decrease in Prepayment')"

   Set Rst = oSaccoMaster.GetRecordset(sql)
    
    If Not Rst.EOF Then
       RunBal = RunBal + IIf(IsNull(Rst.Fields(0)), 0, Rst.Fields(0))
       sql = ""
    sql = "update cashflows set amty1=" & Rst.Fields(0) & " where yy1=" & year(dtpFinishDate) & " and descr='Net Changes In Working Capital' and ddescr='(Increase)/Decrease in Prepayment'"
    oSaccoMaster.GetRecordset (sql)
    
    End If
    
    
      'xx Net Cashflow From Working Capital
      'Income tax paid
      
      
       'xx Interest Paid
          sql = ""
    sql = "SELECT (SELECT ISNULL(SUM(T.OBal),0)  FROM tbbalance T INNER JOIN GLSETUP G  ON G.AccNo=T.accno where Scluster='Net Changes In Working Capital' and  G.Fcluster='Interest Paid')"

   Set Rst = oSaccoMaster.GetRecordset(sql)
    
    If Not Rst.EOF Then
       RunBal = RunBal + IIf(IsNull(Rst.Fields(0)), 0, Rst.Fields(0))
       sql = ""
    sql = "update cashflows set amty1=" & Rst.Fields(0) & " where yy1=" & year(dtpFinishDate) & " and descr='Net Changes In Working Capital' and ddescr='Interest Paid'"
    oSaccoMaster.GetRecordset (sql)
    
    End If
    
    
           'xx Dividends Paid
          sql = ""
    sql = "SELECT (SELECT ISNULL(SUM(T.OBal),0)  FROM tbbalance T INNER JOIN GLSETUP G  ON G.AccNo=T.accno where Scluster='Net Changes In Working Capital' and  G.Fcluster='Dividends Paid')"

   Set Rst = oSaccoMaster.GetRecordset(sql)
    
    If Not Rst.EOF Then
       RunBal = RunBal + IIf(IsNull(Rst.Fields(0)), 0, Rst.Fields(0))
       sql = ""
    sql = "update cashflows set amty1=" & Rst.Fields(0) & " where yy1=" & year(dtpFinishDate) & " and descr='Net Changes In Working Capital' and ddescr='Dividends Paid'"
    oSaccoMaster.GetRecordset (sql)
    
    End If
    
    
    'xxx Net Cash Generated From Operating Activities
    
    
    
    'Net Changes In Investing Activities
    'xxx  Purchase of property and equipment
      RunBal = 0
              sql = ""
    sql = "SELECT (SELECT ISNULL(SUM(T.OBal),0)  FROM tbbalance T INNER JOIN GLSETUP G  ON G.AccNo=T.accno where Scluster='Net Changes In Investing Activities' and  G.Fcluster='Purchase of property and equipment')"

   Set Rst = oSaccoMaster.GetRecordset(sql)
    
    If Not Rst.EOF Then
       RunBal = RunBal + IIf(IsNull(Rst.Fields(0)), 0, Rst.Fields(0))
       sql = ""
    sql = "update cashflows set amty1=" & Rst.Fields(0) & " where yy1=" & year(dtpFinishDate) & " and descr='Net Changes In Investing Activities' and ddescr='Purchase of property and equipment'"
    oSaccoMaster.GetRecordset (sql)
    
    End If
    
    
    'xxx  Purchase of Intangible Assets
                  sql = ""
    sql = "SELECT (SELECT ISNULL(SUM(T.OBal),0)  FROM tbbalance T INNER JOIN GLSETUP G  ON G.AccNo=T.accno where Scluster='Net Changes In Investing Activities' and  G.Fcluster='Purchase of Intangible Assets')"

   Set Rst = oSaccoMaster.GetRecordset(sql)
    
    If Not Rst.EOF Then
       RunBal = RunBal + IIf(IsNull(Rst.Fields(0)), 0, Rst.Fields(0))
       sql = ""
    sql = "update cashflows set amty1=" & Rst.Fields(0) & " where yy1=" & year(dtpFinishDate) & " and descr='Net Changes In Investing Activities' and ddescr='Purchase of Intangible Assets'"
    oSaccoMaster.GetRecordset (sql)
    
    End If
    
    
    ' Net Cashflow From Investing Activities
           sql = ""
    sql = "update cashflows set amty1=" & RunBal & " where yy1=" & year(dtpFinishDate) & " and descr='Net Changes In Investing Activities' and ddescr='Net Cashflow From Investing Activities'"
    oSaccoMaster.GetRecordset (sql)
     RunBal = 0
    
    
    'xx Financing Activities
    'xx Proceeds From Issues Of Capital To Be Alloted
    
    sql = ""
    sql = "SELECT (SELECT ISNULL(SUM(T.OBal),0)  FROM tbbalance T INNER JOIN GLSETUP G  ON G.AccNo=T.accno where Scluster='Financing Activities' and  G.Fcluster='Proceeds From Issues Of Capital To Be Alloted')"

   Set Rst = oSaccoMaster.GetRecordset(sql)
    
    If Not Rst.EOF Then
       RunBal = RunBal + IIf(IsNull(Rst.Fields(0)), 0, Rst.Fields(0))
       sql = ""
    sql = "update cashflows set amty1=" & Rst.Fields(0) & " where yy1=" & year(dtpFinishDate) & " and descr='Financing Activities' and ddescr='Proceeds From Issues Of Capital To Be Alloted'"
    oSaccoMaster.GetRecordset (sql)
    
    End If
    
    
        'xx Proceeds From long-term Capital
    
    sql = ""
    sql = "SELECT (SELECT ISNULL(SUM(T.OBal),0)  FROM tbbalance T INNER JOIN GLSETUP G  ON G.AccNo=T.accno where Scluster='Financing Activities' and  G.Fcluster='Proceeds From long-term Capital')"

   Set Rst = oSaccoMaster.GetRecordset(sql)
    
    If Not Rst.EOF Then
       RunBal = RunBal + IIf(IsNull(Rst.Fields(0)), 0, Rst.Fields(0))
       sql = ""
    sql = "update cashflows set amty1=" & Rst.Fields(0) & " where yy1=" & year(dtpFinishDate) & " and descr='Financing Activities' and ddescr='Proceeds From long-term Capital'"
    oSaccoMaster.GetRecordset (sql)
    
    End If
    
            'xx Prior Year Adjustment
    
    sql = ""
    sql = "SELECT (SELECT ISNULL(SUM(T.OBal),0)  FROM tbbalance T INNER JOIN GLSETUP G  ON G.AccNo=T.accno where Scluster='Financing Activities' and  G.Fcluster='Prior Year Adjustment')"

   Set Rst = oSaccoMaster.GetRecordset(sql)
    
    If Not Rst.EOF Then
       RunBal = RunBal + IIf(IsNull(Rst.Fields(0)), 0, Rst.Fields(0))
       sql = ""
    sql = "update cashflows set amty1=" & Rst.Fields(0) & " where yy1=" & year(dtpFinishDate) & " and descr='Financing Activities' and ddescr='Prior Year Adjustment'"
    oSaccoMaster.GetRecordset (sql)
    
    End If
    
                'xx Other Loans
    
    sql = ""
    sql = "SELECT (SELECT ISNULL(SUM(T.OBal),0)  FROM tbbalance T INNER JOIN GLSETUP G  ON G.AccNo=T.accno where Scluster='Financing Activities' and  G.Fcluster='Other Loans')"

   Set Rst = oSaccoMaster.GetRecordset(sql)
    
    If Not Rst.EOF Then
       RunBal = RunBal + IIf(IsNull(Rst.Fields(0)), 0, Rst.Fields(0))
       sql = ""
    sql = "update cashflows set amty1=" & Rst.Fields(0) & " where yy1=" & year(dtpFinishDate) & " and descr='Financing Activities' and ddescr='Other Loans'"
    oSaccoMaster.GetRecordset (sql)
    
    End If
    
    'xx Net Financing Activities
           sql = ""
    sql = "update cashflows set amty1=" & RunBal & " where yy1=" & year(dtpFinishDate) & " and descr='Financing Activities' and ddescr='Net Financing Activities'"
    oSaccoMaster.GetRecordset (sql)
    
    reportname = "CashFlows.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
    
 
End Sub

Private Sub cmdfinancialR_Click()
    Dim AccNo As String
    Dim FAssets As Double, FromDate As Date, ToDate As Date
    Dim CAssets As Double, CLiabilities As Double, TAssets As Double, Equity As Double, LInt As Double
    Dim ACCBAL As Double, CStock As Double, GrossIncome As Double, TDebt As Double, LLiabilities As Double
    Dim TSales As Double, TPurchases As Double, EBIT As Double, Sharec As Double, AccName As String
    
    If MsgBox("Requires Processing of Trialbalance First ," & "Do You Wish To Continue?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    FromDate = Format(dtpStartDate, "dd/mm/yyyy")
    ToDate = Format(dtpFinishDate, "dd/mm/yyyy")
    
    oSaccoMaster.ExecuteThis ("Truncate table FinancialRatio")
    'Current Assets
    Set Rst = oSaccoMaster.GetRecordset("SELECT  (SELECT     isnull(SUM(Amount),0) FROM  tbbalance1 WHERE transtype = 'DR' and AcGroup='Current Assets') AS Debits, (SELECT     isnull(SUM(Amount),0) FROM  tbbalance1 WHERE transtype = 'CR' and AcGroup='Current Assets') AS Credits")
    If Not Rst.EOF Then
        CAssets = Rst("Debits") - Rst("Credits")
    Else
        CAssets = 0
    End If
    
    'Current Liabilities
    Set Rst = oSaccoMaster.GetRecordset("SELECT  (SELECT     isnull(SUM(Amount),0) FROM  tbbalance1 WHERE transtype = 'DR' and AcGroup='Current liabilities') AS Debits, (SELECT     isnull(SUM(Amount),0) FROM  tbbalance1 WHERE transtype = 'CR' and AcGroup='Current liabilities') AS Credits")
    If Not Rst.EOF Then
        CLiabilities = Rst("Credits") - Rst("Debits")
    Else
        CLiabilities = 0
    End If
    'Closing stock
    Set Rst = oSaccoMaster.GetRecordset("SELECT  (SELECT     isnull(SUM(Amount),0) FROM  tbbalance1 WHERE Accno = 'A074' or Accno = 'A003' or Accno = 'A010') AS Debits")
    If Not Rst.EOF Then
        CStock = Rst("Debits")
    Else
        CStock = 0
    End If
    
    'Gross income
    Set Rst = oSaccoMaster.GetRecordset("SELECT  (SELECT     isnull(SUM(Amount),0) FROM  tbbalance1 WHERE transtype = 'CR' and AcGroup='Sales') AS Debits, (SELECT isnull(SUM(Amount),0) FROM  tbbalance1 WHERE transtype = 'DR' and AcGroup='Purchases') AS Credits")
    If Not Rst.EOF Then
        TSales = Rst("Debits")
        GrossIncome = Rst("Debits") - Rst("Credits")
        TPurchases = Rst("Credits")
    Else
        GrossIncome = 0
    End If
    
    'EBIT OPERATING INCOME
    Set Rst = oSaccoMaster.GetRecordset("SELECT  (SELECT isnull(SUM(Amount),0) FROM  tbbalance1 WHERE Accno = 'R001') AS Debits")
    If Not Rst.EOF Then
        EBIT = Rst("Debits")
    Else
        EBIT = 0
    End If
    'Total Assets
    Set Rst = oSaccoMaster.GetRecordset("SELECT  (SELECT     isnull(SUM(Amount),0) FROM  tbbalance1 WHERE transtype = 'DR' and AccGroup='Assets') AS Debits, (SELECT     isnull(SUM(Amount),0) FROM  tbbalance1 WHERE transtype = 'CR' and AccGroup='Assets') AS Credits")
    If Not Rst.EOF Then
        TAssets = Rst("Debits") - Rst("Credits")
    Else
        TAssets = 0
    End If
    
    'Fixed Assets
    Set Rst = oSaccoMaster.GetRecordset("SELECT  (SELECT     isnull(SUM(Amount),0) FROM  tbbalance1 WHERE transtype = 'DR' and AcGroup='Fixed assets') AS Debits, (SELECT     isnull(SUM(Amount),0) FROM  tbbalance1 WHERE transtype = 'CR' and AcGroup='Fixed assets') AS Credits")
    If Not Rst.EOF Then
        FAssets = Rst("Debits") - Rst("Credits")
    Else
        FAssets = 0
    End If
    
    'Share Holders Equity
    Set Rst = oSaccoMaster.GetRecordset("SELECT  (SELECT     isnull(SUM(Amount),0) FROM  tbbalance1 WHERE transtype = 'DR' and AcGroup='Share Holders Equity') AS Debits, (SELECT     isnull(SUM(Amount),0) FROM  tbbalance1 WHERE transtype = 'CR' and AcGroup='Share Holders Equity') AS Credits")
    If Not Rst.EOF Then
        Equity = Rst("Credits") - Rst("Debits")
    Else
        Equity = 0
    End If
        'Loan interest)
    Set Rst = oSaccoMaster.GetRecordset("SELECT  (SELECT isnull(SUM(Amount),0) FROM  tbbalance1 WHERE Accno = 'E091') AS Debits")
    If Not Rst.EOF Then
        LInt = Rst("Debits")
    Else
        LInt = 0
    End If
            'Share holder's equity
    Set Rst = oSaccoMaster.GetRecordset("SELECT  (SELECT isnull(SUM(Amount),0) FROM  tbbalance1 WHERE Accno = 'L020') AS Debits")
    If Not Rst.EOF Then
        Sharec = Rst("Debits")
    Else
        Sharec = 0
    End If
    
    'Long Term Liabilities
    Set Rst = oSaccoMaster.GetRecordset("SELECT  (SELECT isnull(SUM(Amount),0) FROM  tbbalance1 WHERE Accno = 'L002') AS Debits")
    If Not Rst.EOF Then
        LLiabilities = Rst("Debits")
    Else
        LLiabilities = 0
    End If
    

       
        
        sql = "Set DateFormat DMY INSERT INTO [FinancialRatio] (CAssets,CLiabilities,CStock,TSales,TPurchases,GrossIncome,EBIT,TAssets,FAssets,Equity,LInt,Sharec,StartDate,EndDate)"
        
        sql = sql & " Values(" & CAssets & "," & CLiabilities & "," & CStock & "," & TSales & "," & TPurchases & "," & GrossIncome & "," & EBIT & _
        "," & TAssets & "," & FAssets & "," & Equity & "," & LInt & "," & Sharec & ",'" & FromDate & "','" & ToDate & "')"
            
        If Not oSaccoMaster.Execute(sql) Then
            GoTo sysError
        End If
        MsgBox "Financial Ratios Processed Successfully!", vbInformation, Me.Caption
    Exit Sub
sysError:
    cmdprocess.Enabled = True
    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage), vbInformation
        
End Sub

Private Sub cmdincomestmnt_Click()
   ' reportname = "incomeandexpenditure.rpt"
    reportname = "Incomestatement1.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
End Sub

Private Sub cmdmonthlyincome_Click()
    Dim transtype As String, DocumentNo As String, accType As String, accGroup As String, AcGroup As String, AccName As String
    Dim DataSource As String, errmsg As String, mMonth As Double, Date1 As Date, Date2 As Date, Date3 As Date
    Dim rsAccounts As New Recordset, AccNo As String, transdate As Date, ACCBAL As Double, _
    Amount As Double, Account As Account_Details, Mmonth1 As Long, yYear As Double, Mmonth2 As Long, Y As Long

    Set rsAccounts = New ADODB.Recordset
    Mmonth1 = month(Startdate99)
    mMonth = Mmonth1
    Mmonth2 = month(Enddate99)
    Date1 = Startdate99
    Date3 = Enddate99
    yYear = year(Enddate99)
    
    sql = ("TRUNCATE  TABLE Tbbalance2")
    oSaccoMaster.ExecuteThis (sql)
    Y = Mmonth1
    For Y = Mmonth1 To Mmonth2
        Date1 = DateSerial(year(Date1), month(Date1), 1)
        Date2 = DateSerial(year(Date1), month(Date1) + 1, 1 - 1)
        
        Set rsAccounts = oSaccoMaster.GetRecordset("SELECT Accno,NormalBal,glaccType,glaccname,GlAccMainGroup,GlAccGroup FROM GLSETUP" _
                & " WHERE AccCategory='" & cbpdpt & "' and (GlAccMainGroup='EXPENSES' or GlAccMainGroup='INCOME')  ORDER BY ACCNO")
       
               With rsAccounts
                If Not .EOF Then
                    prgStatus.Visible = True
                    lblStatus.Visible = True
                    lblAccount.Visible = True
                    prgStatus.Max = .RecordCount
                    I = 0
                    While Not .EOF
                        DoEvents
                        I = I + 1
                        lblStatus.Caption = CStr(Round((I / .RecordCount) * 100, 0)) & " %"
                        prgStatus.Value = .AbsolutePosition 'Round((I / .RecordCount) * 100, 0)
                        AccNo = !AccNo
                        lblAccount = AccNo
                        accType = IIf(IsNull(!Glacctype), "", Trim(!Glacctype))
                        accGroup = IIf(IsNull(!GlAccMainGroup), "", Trim(!GlAccMainGroup))
                        AcGroup = IIf(IsNull(!GLAccGroup), "", Trim(!GLAccGroup))
                        AccName = IIf(IsNull(!GlAccName), "", Trim(!GlAccName))
                        lblaccname = AccName
        
                         ACCBAL = getGlPeriodicTrans(AccNo, Date1, Date2)
                        
                        If !NormalBal = "Debit" Then
                            If ACCBAL >= 0 Then
                                transtype = "DR"
                            Else
                                transtype = "CR"
                            End If
                        Else
                            If ACCBAL >= 0 Then
                                transtype = "CR"
                            Else
                                transtype = "DR"
                            End If
                        End If
                                   
                          ACCBAL = Abs(ACCBAL)
                            If mMonth = Mmonth1 Then
                            sql = "Set DateFormat DMY INSERT INTO [Tbbalance2] ([AccNo],[AccName], [Amount],[Transtype],[StartDate], [EndDate], [AuditID], [AccType], [AccGroup],[AcGroup],[Jan],[NormalBal])"
                            sql = sql & " Values('" & AccNo & "','" & AccName & "'," & ACCBAL & ",'" & transtype & "','" & Startdate99 & "','" & Enddate99 & _
                            "','" & User & "','" & accType & "','" & accGroup & "','" & AcGroup & "'," & ACCBAL & ",'" & !NormalBal & "')"
                            If Not oSaccoMaster.Execute(sql) Then
                                GoTo sysError
                            End If
                            ElseIf Mmonth1 - mMonth = 1 Then
                            oSaccoMaster.ExecuteThis ("Update [Tbbalance2] set Feb=" & ACCBAL & " where Accno='" & AccNo & "' ")
                            ElseIf Mmonth1 - mMonth = 2 Then
                            oSaccoMaster.ExecuteThis ("Update [Tbbalance2] set Mar=" & ACCBAL & " where Accno='" & AccNo & "' ")
                            ElseIf Mmonth1 - mMonth = 3 Then
                            oSaccoMaster.ExecuteThis ("Update [Tbbalance2] set Apr=" & ACCBAL & " where Accno='" & AccNo & "' ")
                            End If
                    
                        
                        .MoveNext
                    Wend
                   
                End If
            End With
    
        Date1 = DateSerial(year(Date1), month(Date1) + 1, 1)
        Mmonth1 = Mmonth1 + 1
    Next Y
MsgBox "Process Done", vbInformation

    reportname = "Incomestatement2.rpt"
    STRFORMULA = ""
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
Exit Sub
sysError:
    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage), vbInformation
End Sub
Private Function Print_Trial_Balance99(Startdate As Date, FinishDate As Date, DataSource As String, _
errmsg As String) As Boolean
    On Error GoTo sysError
    Dim rsTBBalance As New Recordset, transtype As String, Account As Account_Details, _
    balance As Double, rsAccounts As New Recordset, AccNo As String, NormalBal As String, _
    OpTransType As String, AccName As String, cnn As New ADODB.Connection
    'XXXXXXXXXXXXXXXXXXXXX Open The Database XXXXXXXXXXXXXXXXXX
    
    Dim mysql As String
    Dim Provider As String
    Provider = "MAZIWA"
    cnn.Open Provider, "bi"
    
    sql = ("Delete From TBBALANCE")
    oSaccoMaster.ExecuteThis (sql)
    'XXXXXXXXXXXXXXXXXXXXX Get_The_Accounts_And_Balances from GLSetUp XXXXXXXXXXXXX
    Set rsAccounts = cnn.Execute("set dateformat dmy Select * From GLSETUP where NewGLOpeningBalDate >='" & Startdate & "' Order By AccNo")
    If Not rsAccounts.EOF Then
    With rsAccounts
    
        If .State = adStateOpen Then
            While Not .EOF
                DoEvents
                AccNo = IIf(IsNull(!AccNo), "", !AccNo)
                Account = Get_Account_Details(AccNo, DataSource, errmsg)
                If Account.AccountNo <> "" Then
                    'Balance = Account.OpeningBalance
                    balance = !NewGLOpeningBal
                    NormalBal = Account.NormalBalance
                    AccName = Account.AccountName
                    'XXXXXXXXXXXXXX Get Transactions From Temp TBBalance XXXXXXXXXXXXXXXX
                    mysql = "Select TransType,Sum(Amount) as Amount" _
                    & " From TEMTTBBALANCE where AccNo='" & AccNo & "' Group By TransType"
                    
                    Debug.Print mysql
                    
                    Set rsTBBalance = cnn.Execute(mysql)
                    
                    With rsTBBalance
                        If .State = adStateOpen Then
                            While Not .EOF
                                DoEvents
                                transtype = IIf(IsNull(!transtype), "", !transtype)
                                Select Case NormalBal
                                    Case "DR"
                                    Select Case transtype
                                        Case "DR"
                                        balance = balance + IIf(IsNull(!Amount), 0, !Amount)
                                        Case "CR"
                                        balance = balance - IIf(IsNull(!Amount), 0, !Amount)
                                    End Select
                                    Case "CR"
                                    Select Case transtype
                                        Case "DR"
                                        balance = balance - IIf(IsNull(!Amount), 0, !Amount)
                                        Case "CR"
                                        balance = balance + IIf(IsNull(!Amount), 0, !Amount)
                                    End Select
                                End Select
                                .MoveNext
                            Wend
                        End If
                    End With
                    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX Save To TBBalance XXXXXXXXXXXXXXXXXXXXXXXXXXXX
                    Select Case balance
                        Case Is > 0
                        Case Is < 0
                        balance = balance * (-1)
                        Select Case NormalBal
                            Case "DR"
                            NormalBal = "CR"
                            Case "CR"
                            NormalBal = "DR"
                        End Select
                    End Select
                    If Not Save_TrialBalance99(AccNo, AccName, balance, NormalBal, 0, dtpFinishDate, _
                    User, "", "", 0, DataSource, errmsg) Then
                        Print_Trial_Balance99 = False
                        Exit Function
                    End If
                End If
                .MoveNext
            Wend
            '// do the rest
    
       
        End If
        
        
    End With
    
    Else
    '// work on the items here only with o figures
    'MsgBox "here "
    Set rsAccounts = cnn.Execute("set dateformat dmy Select * From GLSETUP  Order By AccNo")
    With rsAccounts
    
        If .State = adStateOpen Then
            While Not .EOF
                DoEvents
                AccNo = IIf(IsNull(!AccNo), "", !AccNo)
                Account = Get_Account_Details(AccNo, DataSource, errmsg)
                If Account.AccountNo <> "" Then
                    'Balance = Account.OpeningBalance
                    balance = 0
                    NormalBal = Account.NormalBalance
                    AccName = Account.AccountName
                    'XXXXXXXXXXXXXX Get Transactions From Temp TBBalance XXXXXXXXXXXXXXXX
                    mysql = "Select TransType,Sum(Amount) as Amount" _
                    & " From TEMTTBBALANCE where AccNo='" & AccNo & "' Group By TransType"
                    
                    Debug.Print mysql
                    
                    Set rsTBBalance = cnn.Execute(mysql)
                    
                    With rsTBBalance
                        If .State = adStateOpen Then
                            While Not .EOF
                                DoEvents
                                transtype = IIf(IsNull(!transtype), "", !transtype)
                                Select Case NormalBal
                                    Case "DR"
                                    Select Case transtype
                                        Case "DR"
                                        balance = balance + IIf(IsNull(!Amount), 0, !Amount)
                                        Case "CR"
                                        balance = balance - IIf(IsNull(!Amount), 0, !Amount)
                                    End Select
                                    Case "CR"
                                    Select Case transtype
                                        Case "DR"
                                        balance = balance - IIf(IsNull(!Amount), 0, !Amount)
                                        Case "CR"
                                        balance = balance + IIf(IsNull(!Amount), 0, !Amount)
                                    End Select
                                End Select
                                .MoveNext
                            Wend
                        End If
                    End With
                    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX Save To TBBalance XXXXXXXXXXXXXXXXXXXXXXXXXXXX
                    Select Case balance
                        Case Is > 0
                        Case Is < 0
                        balance = balance * (-1)
                        Select Case NormalBal
                            Case "DR"
                            NormalBal = "CR"
                            Case "CR"
                            NormalBal = "DR"
                        End Select
                    End Select
                    If Not Save_TrialBalance99(AccNo, AccName, balance, NormalBal, 0, FinishDate, _
                    User, "", "", 0, DataSource, errmsg) Then
                        Print_Trial_Balance99 = False
                        Exit Function
                    End If
                End If
                .MoveNext
            Wend
            '// do the rest
    
       
        End If
        
        
    End With

    End If
    Exit Function
sysError:
    Print_Trial_Balance99 = False
End Function
Private Function Save_TrialBalance99(AccNo As String, AccName As String, Amount As _
Double, transtype As String, Closed As Long, transdate As Date, auditid As _
String, accType As String, accGroup As String, Budget As Double, DataSource As _
String, errmsg As String) As Boolean
    On Error GoTo sysError
    Set cn = New ADODB.Connection
    Provider = "MAZIWA"
    cn.Open Provider
    With cn
        If .State = adStateClosed Then
            .Open DataSource, "bi"
        End If
        .Execute ("Set DateFormat DMY Exec Save_TrialBalance '" & AccNo & "','" & _
        AccName & "'," & Amount & ",'" & transtype & "'," & Closed & ",'" & transdate & _
        "','" & dtpStartDate & "','" & dtpFinishDate & "','" & auditid & "','" & accType & "','" & accGroup & "'," & Budget)
        
    End With
    Save_TrialBalance99 = True
    Exit Function
sysError:
    Save_TrialBalance99 = False
    MsgBox err.description
End Function

Private Function Save_TEMTTBBALANCE99(AccNo As String, transdate As Date, Amount As Double, _
transtype As String, DocumentNo As String, DataSource As String, errmsg As String) As Boolean
    'On Error GoTo SysError
    Provider = "MAZIWA"
    Set cn = New ADODB.Connection
    cn.Open Provider
    With cn
        If .State = adStateOpen Then
            .Execute ("Set DateFormat DMY Exec Save_TEMTTBBALANCE '" & AccNo & "','" & _
            transdate & "'," & Amount & ",'" & transtype & "','" & DocumentNo & "'")
        End If
    End With
    Save_TEMTTBBALANCE99 = True
    Exit Function
sysError:
    Save_TEMTTBBALANCE99 = False
End Function
Private Sub monthlyincome()
  Dim rsAccounts As Recordset
  Dim rsBudgets As Recordset
  '// get the budget amount and then variances
  Set rsAccounts = oSaccoMaster.GetRecordset("Select * From GLSETUP where glacctype='Income Statement' order by AccNo")
         While Not rsAccounts.EOF
  
            Set rsBudgets = oSaccoMaster.GetRecordset("Set DateFormat DMY Select Budgetted" _
            & " As BudgetAmount From BUDGETS where AccNo='" & rsAccounts.Fields("accno") & "' and mmonth='" _
            & month(Startdate99) & "' and yyear='" & year(Startdate99) & "'")
            If Not rsBudgets.EOF Then
            Dim b As Currency
            '//updates on the tem
            b = rsBudgets.Fields(0)
            sql = ""
            sql = "UPDATE    TBBALANCE  SET budgetAMOUNT =" & b & " where accno='" & rsAccounts.Fields("accno") & "'"
            oSaccoMaster.ExecuteThis (sql)
            Else
                        sql = ""
            sql = "UPDATE    TBBALANCE  SET budgetAMOUNT=0 where accno='" & rsAccounts.Fields("accno") & "'"
            oSaccoMaster.ExecuteThis (sql)
            End If
            b = 0
            rsAccounts.MoveNext
        Wend
        
        '//GET THE NETINCOME
        Dim totexpenses As Currency, TotIncome  As Currency
        Dim totexpensesB As Currency, TotIncomeB  As Currency
       
    Set rs = oSaccoMaster.GetRecordset("SELECT     SUM(TBBALANCE.Amount) AS Expr1,sum(budgetamount) as AI  FROM         TBBALANCE TBBALANCE INNER JOIN  GLSETUP GLSETUP ON TBBALANCE.AccNo = GLSETUP.AccNo WHERE     (GLSETUP.GlAccMainGroup = 'INCOME')")
        If Not rs.EOF Then
            If Not IsNull(rs.Fields(0)) Then TotIncome = rs.Fields(0)
            If Not IsNull(rs.Fields(1)) Then TotIncomeB = rs.Fields(1)
        End If
    Set Rst = oSaccoMaster.GetRecordset("SELECT     SUM(TBBALANCE.Amount) AS Expr1,sum(budgetamount) as AE  FROM         TBBALANCE TBBALANCE INNER JOIN  GLSETUP GLSETUP ON TBBALANCE.AccNo = GLSETUP.AccNo WHERE     (GLSETUP.GlAccMainGroup = 'EXPENSES') or (glsetup.glaccmaingroup='PRODUCTION COST')")
        If Not Rst.EOF Then
            If Not IsNull(Rst.Fields(0)) Then totexpenses = Rst.Fields(0)
            If Not IsNull(Rst.Fields(1)) Then totexpensesB = Rst.Fields(1)
        End If
        
        If Not Save_TB("", "Net Income", TotIncome - totexpenses, "", 0, _
                    dtpFinishDate, dtpStartDate, dtpFinishDate, User, ErrorMessage, "2", "A", TotIncomeB - totexpensesB) Then
                        If ErrorMessage <> "" Then
                            MsgBox ErrorMessage, vbInformation, Me.Caption
                            ErrorMessage = ""
                        End If
        End If

  '//show the reports
  frmviewincomestatements.Show vbModal
   
 g = False
  

    cmdprocess.Enabled = True
    Exit Sub
sysError:
    cmdprocess.Enabled = True
    MsgBox err.description, vbInformation, Me.Caption

End Sub



Private Sub EOY_Processing(EOYDate As Date)
Dim AccNo As String, Amount As String, transdate As Date, Glacctype As String

Set rs = Nothing
Set rs = oSaccoMaster.GetRecordset("set dateformat dmy select * from TBBalance where TransDate='" & EOYDate & "'")
If rs.EOF Then
    MsgBox "Trial Balance has not been generated, Please generate it before proceeding.", vbCritical, Me.Caption
    Exit Sub
End If

Set rs = oSaccoMaster.GetRecordset("select AccNo,Glacctype from GlSetup order by accno")
With rs
    If Not .EOF Then
      While Not .EOF
      Me.Caption = !AccNo
        Set Rst = oSaccoMaster.GetRecordset("set dateformat dmy select AccNo,Amount,transdate From TBBalance" _
        & " where AccNO='" & !AccNo & "' and transdate='" & EOYDate & "' order by AccNO")
         If Not Rst.EOF Then
         
          If !Glacctype = "Income Statement" Then
          Amount = 0
          Set Rst1 = oSaccoMaster.GetRecordset("set dateformat dmy update GLSETUP set NewGLOpeningBal=0,NewGLOpeningBalDate='" & EOYDate & "',CurrentBal=" & Amount & " where AccNo='" & !AccNo & "'")

          Else
          Amount = Rst!Amount
          Set Rst1 = oSaccoMaster.GetRecordset("set dateformat dmy update GLSETUP set NewGLOpeningBal=" & Amount & ",NewGLOpeningBalDate='" & EOYDate & "',CurrentBal=" & Amount & " where AccNo='" & !AccNo & "'")

         End If

    End If
    .MoveNext
    Wend
End If
End With

End Sub

Private Sub cmdprintexp_Click()
    reportname = "Expense.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
End Sub

Private Sub cmdprintratio_Click()
    reportname = "FinancialRatios.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
End Sub

Private Sub cmdprocess_Click()
    On Error Resume Next
    Dim AccNo As String
    Dim suspense As Double
    Dim Debits As Double, Credits As Double
    Dim ACCBAL As Double, FromDate As Date, ToDate As Date
    Dim transtype As String, DocumentNo As String, accType As String, accGroup As String, AcGroup As String, AccName As String
    
    If cbodepartment = "" Then
      MsgBox "Select Activity option first", vbInformation, Me.Caption
       cbodepartment.SetFocus
      Exit Sub
    End If
    
     oSaccoMaster.ExecuteThis ("Truncate table TBBALANCE1")
    If Trim(UCase(cbodepartment)) = "ALL" Then
    sql = "SELECT Accno,NormalBal,glaccType,glaccname,GlAccMainGroup,GlAccGroup FROM GLSETUP  ORDER BY ACCNO"
    Else
    sql = "SELECT Accno,NormalBal,glaccType,glaccname,GlAccMainGroup,GlAccGroup FROM GLSETUP  WHERE AccCategory='" & cbodepartment & "'  ORDER BY ACCNO"
    End If
    Set Rst = oSaccoMaster.GetRecordset(sql)
    With Rst
        If Not .EOF Then
            prgStatus.Visible = True
            lblStatus.Visible = True
            lblAccount.Visible = True
            prgStatus.Max = .RecordCount
            'prgStatus.Min = 0
            I = 0
            While Not .EOF
                DoEvents
                I = I + 1
                lblStatus.Caption = CStr(Round((I / .RecordCount) * 100, 0)) & " %"
                prgStatus.Value = .AbsolutePosition 'Round((I / .RecordCount) * 100, 0)
                AccNo = !AccNo
            'If AccNo = "673-217" Then MsgBox ""
                lblAccount = AccNo
                accType = IIf(IsNull(!Glacctype), "", !Glacctype)
                accGroup = IIf(IsNull(!GlAccMainGroup), "", !GlAccMainGroup)
                AcGroup = IIf(IsNull(!GLAccGroup), "", !GLAccGroup)
                AccName = IIf(IsNull(!GlAccName), "", !GlAccName)
                lblaccname = AccName
                
                
                FromDate = dtpStartDate
                ToDate = dtpFinishDate
                 
'                If Branch = "ALL" Then
                 ACCBAL = getTBGlBalance(AccNo, FromDate, ToDate)
'                Else
'                ACCBAL = getTBGlBalance_branch(accno, FromDate, ToDate, Branch)
'                End If
                
                
                If Not success Then
                    GoTo sysError
                End If
                
                If !NormalBal = "Debit" Then
'                    If ACCBAL >= 0 Then
                        transtype = "DR"
'                    Else
'                        transtype = "CR"
'                    End If
                Else
'                    If ACCBAL >= 0 Then
                        transtype = "CR"
'                    Else
'                        transtype = "DR"
'                    End If
                End If
                           
'                ACCBAL = Abs(ACCBAL)
'            If ACCBAL <> 0 Then
                    sql = "Set DateFormat DMY INSERT INTO [tbbalance1] ([AccNo],[AccName], [Amount],[Transtype],[StartDate], [EndDate], [AuditID], [AccType], [AccGroup],[AcGroup],[BudgetAmount],OBAL,DR,CR)"
                    sql = sql & " Values('" & AccNo & "','" & AccName & "'," & ACCBAL & ",'" & transtype & "','" & dtpStartDate & "','" & dtpFinishDate.Value & _
                    "','" & User & "','" & accType & "','" & accGroup & "','" & AcGroup & "',0," & TBOpeningBal & "," & totaldr & "," & totalcr & ")"
                        
                    If Not oSaccoMaster.Execute(sql) Then
                        GoTo sysError
                      End If
'            End If
                
                .MoveNext
            Wend
        Else
            prgStatus.Visible = False
            lblStatus.Visible = False
            lblAccount.Visible = False
        End If
    End With
    
    Set Rst = oSaccoMaster.GetRecordset("SELECT(SELECT isnull(SUM(Amount),0) FROM  tbbalance1 WHERE transtype = 'DR') AS Debits, (SELECT     isnull(SUM(Amount),0) FROM  tbbalance1 WHERE transtype = 'CR') AS Credits")
    If Not Rst.EOF Then
        If Rst("Debits") > Rst("Credits") Then
            Credits = Rst("Debits") - Rst("Credits")
            ACCBAL = Rst("Debits") - Rst("Credits")
            transtype = "CR"
        Else
            Debits = Rst("Credits") - Rst("Debits")
            ACCBAL = Rst("Credits") - Rst("Debits")
            transtype = "DR"
        End If
        
        If ACCBAL > 0 Then
            sql = "Set DateFormat DMY INSERT INTO [tbbalance1] ([AccNo],[AccName], [Amount],[Transtype], [Closed],[StartDate], [EndDate], [AuditID], [AccType], [AccGroup], [BudgetAmount])"
            
            sql = sql & " Values('" & SuspenseAcc & "','SUSPENSE ACC'," & ACCBAL & ",'" & transtype & "',0,'" & dtpStartDate & "','" & dtpFinishDate.Value & _
            "','" & User & "','" & accType & "','" & accGroup & "',0)"
                
            If Not oSaccoMaster.Execute(sql) Then
                GoTo sysError
            End If
        End If
        
        'For BalanceSheet Items, insert net income
        Set Rst = oSaccoMaster.GetRecordset("SELECT  isnull((SELECT     SUM(Amount) FROM  tbbalance1 WHERE transtype = 'DR' and AccGroup='EXPENSES'),0) AS Debits, isnull((SELECT     SUM(Amount) FROM  tbbalance1 WHERE transtype = 'CR' and AccGroup='INCOME'),0) AS Credits")
        If Not Rst.EOF Then
            If Rst("Debits") > Rst("Credits") Then
                Credits = Rst("Debits") - Rst("Credits")
                ACCBAL = Rst("Debits") - Rst("Credits")
                transtype = "DR"
            Else
                Debits = Rst("Credits") - Rst("Debits")
                ACCBAL = Rst("Credits") - Rst("Debits")
                transtype = "CR"
            End If
            
        If ACCBAL <> 0 Then
            sql = "Set DateFormat DMY INSERT INTO [tbbalance1] ([AccNo],[AccName], [Amount],[Transtype], [Closed],[StartDate], [EndDate], [AuditID], [AccType], [AccGroup], [BudgetAmount])"
            
            sql = sql & " Values('" & REarningsAcc & "','" & UCase("Retained Earnings") & "'," & ACCBAL & ",'" & transtype & "',0,'" & dtpStartDate & "','" & dtpFinishDate.Value & _
            "','" & User & "','" & accType & "','CAPITAL',0)"
                
            If Not oSaccoMaster.Execute(sql) Then
                GoTo sysError
            End If
        End If
        
        End If
    End If
    MsgBox "Process Done", vbInformation
    Exit Sub
sysError:
    cmdprocess.Enabled = True
    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage), vbInformation

End Sub

Private Sub cmdprocessexp_Click()
On Error Resume Next
    Dim AccNo As String
    Dim suspense As Double
    Dim Debits As Double, Credits As Double
    Dim ACCBAL As Double, FromDate As Date, ToDate As Date
    Dim transtype As String, DocumentNo As String, accType As String, accGroup As String, AccName As String
    Dim Branch As String
    
    If cbodepartment = "" Then
      MsgBox "Select Activity option first", vbInformation, Me.Caption
       cbodepartment.SetFocus
      Exit Sub
    End If
    
    If Not oSaccoMaster.Execute("Truncate table TBBALANCE1") Then
        GoTo sysError
    End If
    
    If Trim(UCase(cbodepartment)) = "ALL" Then
    sql = "SELECT Accno,NormalBal,glaccType,glaccname,GlAccMainGroup,GlAccGroup FROM GLSETUP where GlAccMainGroup='EXPENSES'  ORDER BY ACCNO"
    Else
    sql = "SELECT Accno,NormalBal,glaccType,glaccname,GlAccMainGroup,GlAccGroup FROM GLSETUP  WHERE AccCategory='" & cbodepartment & "' and GlAccMainGroup='EXPENSES' ORDER BY ACCNO"
    End If
     Set Rst = oSaccoMaster.GetRecordset(sql)
    With Rst
        If Not .EOF Then
            prgStatus.Visible = True
            lblStatus.Visible = True
            lblAccount.Visible = True
            prgStatus.Max = .RecordCount
            'prgStatus.Min = 0
            I = 0
            While Not .EOF
                DoEvents
                I = I + 1
                lblStatus.Caption = CStr(Round((I / .RecordCount) * 100, 0)) & " %"
                prgStatus.Value = .AbsolutePosition 'Round((I / .RecordCount) * 100, 0)
                AccNo = !AccNo
                lblAccount = AccNo
                accType = !Glacctype
                accGroup = !GLAccGroup
                AccName = !GlAccName
                lblaccname = AccName
                
                FromDate = Format(dtpStartDate, "dd/mm/yyyy")
                ToDate = Format(dtpFinishDate, "dd/mm/yyyy")
                
                 totalcr = 0
                 totaldr = 0
                 TBOpeningBal = 0
                 sql = "set dateformat dmy (SELECT (SELECT isnull(SUM(Amount),0) FROM  GLTRANSACTIONS WHERE DrAccNo = '" & AccNo & "' and Transdate Between '" & FromDate & "' and '" & ToDate & "') AS Debits, (SELECT isnull(SUM(Amount),0) FROM  GLTRANSACTIONS WHERE CrAccNo = '" & AccNo & "' and Transdate Between '" & FromDate & "' and '" & ToDate & "') AS Credits)"
                 Set Rst = oSaccoMaster.GetRecordset(sql)
                If Not Rst.EOF Then
                    totalcr = Rst("Credits")
                    totaldr = Rst("Debits")
                    If Rst("Debits") > Rst("Credits") Then
                        Credits = Rst("Debits") - Rst("Credits")
                        ACCBAL = Rst("Debits") - Rst("Credits")
                        transtype = "DR"
                    Else
                        Debits = Rst("Credits") - Rst("Debits")
                        ACCBAL = Rst("Credits") - Rst("Debits")
                        transtype = "CR"
                    End If
                End If

                
                If !NormalBal = "Debit" Then
                    If ACCBAL >= 0 Then
                        transtype = "DR"
                    Else
                        transtype = "CR"
                    End If
                Else
                    If ACCBAL >= 0 Then
                        transtype = "CR"
                    Else
                        transtype = "DR"
                    End If
                End If
                           
                
            If ACCBAL <> 0 Then
                    sql = "Set DateFormat DMY INSERT INTO [tbbalance1] ([AccNo],[AccName], [Amount],[Transtype],[StartDate], [EndDate], [AuditID], [AccType], [AccGroup], [BudgetAmount],OBAL,DR,CR)"
                    sql = sql & " Values('" & AccNo & "','" & AccName & "'," & ACCBAL & ",'" & transtype & "','" & dtpStartDate & "','" & dtpFinishDate.Value & _
                    "','" & User & "','" & accType & "','" & accGroup & "',0," & TBOpeningBal & "," & totaldr & "," & totalcr & ")"
                        
                    If Not oSaccoMaster.Execute(sql) Then
                        GoTo sysError
                    End If
            End If
                
                .MoveNext
            Wend
        Else
            prgStatus.Visible = False
            lblStatus.Visible = False
            lblAccount.Visible = False
        End If
    End With
      MsgBox "Process Done", vbInformation
      Exit Sub
sysError:
    MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage), vbInformation
End Sub

Private Sub cmdTrialbalance_Click()
' reportname = "Trial Balance.rpt"
 reportname = "TrialBalance.rpt"
 Show_Sales_Crystal_Report "", reportname, CompanyName
End Sub

Private Sub Form_Load()
    dtpStartDate = Date
    dtpFinishDate = Date
    Startdate99 = Date
    Enddate99 = DateSerial(year(Date), month(Date) + 1, 1 - 1)
End Sub

