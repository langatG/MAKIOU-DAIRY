VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmagrovetinc 
   BackColor       =   &H00C000C0&
   Caption         =   "Agrovet Income Statement"
   ClientHeight    =   2745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbpdpt 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmagrovetinc.frx":0000
      Left            =   4440
      List            =   "frmagrovetinc.frx":0007
      TabIndex        =   9
      Text            =   "STORE"
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdprocess 
      Caption         =   "Process"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   1245
   End
   Begin VB.CommandButton cmdincome 
      Caption         =   "Income Statement"
      Height          =   375
      Left            =   2025
      TabIndex        =   7
      Top             =   1800
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker dtpFinishDate 
      Height          =   345
      Left            =   2040
      TabIndex        =   0
      Top             =   360
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
      Format          =   121044995
      CurrentDate     =   39705
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   360
      Left            =   360
      TabIndex        =   1
      Top             =   360
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
      Format          =   121044995
      CurrentDate     =   39705
   End
   Begin MSComctlLib.ProgressBar prgStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label6 
      Caption         =   "Department"
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblclassname 
      BackColor       =   &H00FFFF80&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label lblclass 
      BackColor       =   &H00FFFF80&
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   840
      Width           =   1095
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
      Left            =   390
      TabIndex        =   3
      Top             =   120
      Width           =   810
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
      Left            =   2070
      TabIndex        =   2
      Top             =   135
      Width           =   945
   End
End
Attribute VB_Name = "frmagrovetinc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cash As Double, Credt As Double, AIincome As Double, GAFS As Double, OpStock As Double, TotalSales As Double, accno As String
Dim Purchases As Double, Cin As Double, CGFS As Double, CStock As Double, COS As Double, Gprofit As Double, Exp As Double
Dim NetP As Double, TotExp As Double, OtherInc As Double

Private Sub cmdincome_Click()
 reportname = "AgrovetInc.rpt"
 Show_Sales_Crystal_Report "", reportname, ""
End Sub

Private Sub cmdprocess_Click()
    On Error GoTo ErrorHandler
    
   'Agrovet income statement format
   oSaccoMaster.ExecuteThis ("Truncate Table Ag_Income")
   dtpStartDate = Format(dtpStartDate, "dd/mm/yyyy")
   dtpFinishDate = Format(dtpFinishDate, "dd/mm/yyyy")
   prgStatus.Max = 100
   prgStatus.Value = 0
   NetP = 0
   TotExp = 0
   
   lblclassname.Caption = "Cash Sales"
   'Cash sales 1.0
   'sql = "set dateformat dmy select sum(isnull(Amount,0))Cash from Ag_Receipts where T_Date between '" & dtpStartDate & "' and '" & dtpFinishDate & "' and Sno='Cash' "
    sql = " Set DateFormat DMY select (Select isnull(sum(Amount),0) From GLTRANSACTIONS where TransDate >='" & dtpStartDate & "' and TransDate<='" & dtpFinishDate & "' and CrAccNo='I014')- "
    sql = sql & " (Select isnull(sum(Amount),0) From GLTRANSACTIONS where TransDate >='" & dtpStartDate & "' and TransDate<='" & dtpFinishDate & "' and DrAccNo='I014' )Cash "
 
   Set rss = oSaccoMaster.GetRecordset(sql)
   If Not rss.EOF Then
    cash = rss.Fields(0)
   Else
   cash = 0
   End If
   sql = "set dateformat dmy Insert into Ag_Income(ClassNo, ClassName, AccNo, AccName, Amount, StartDate, EndDate, AuditId)"
   sql = sql & " Values ('1.0','Cash Sales','1.3','SALES'," & cash & ",'" & dtpStartDate & "','" & dtpFinishDate & "','" & User & "')"
   oSaccoMaster.ExecuteThis (sql)
   
      'Credit sales 1.1
      lblclassname.Caption = "Credit Sales"
   'sql = "set dateformat dmy select sum(isnull(Amount,0))Credit from Ag_Receipts where T_Date between '" & dtpStartDate & "' and '" & dtpFinishDate & "' and Sno<>'cash' and Remarks<>'Dispatch to station' "
   sql = " Set DateFormat DMY select (Select isnull(sum(Amount),0) From GLTRANSACTIONS where TransDate >='" & dtpStartDate & "' and TransDate<='" & dtpFinishDate & "' and (CrAccNo='I013'))- "
    sql = sql & " (Select isnull(sum(Amount),0) From GLTRANSACTIONS where TransDate >='" & dtpStartDate & "' and TransDate<='" & dtpFinishDate & "' and (DrAccNo='I013') )Credit "
     Set rss = oSaccoMaster.GetRecordset(sql)
   If Not rss.EOF Then
    Credt = rss.Fields(0)
   Else
   Credt = 0
   End If
   sql = "set dateformat dmy Insert into Ag_Income(ClassNo, ClassName, AccNo, AccName, Amount, StartDate, EndDate, AuditId)"
   sql = sql & " Values ('1.1','Credit Sales','1.3','SALES'," & Credt & ",'" & dtpStartDate & "','" & dtpFinishDate & "','" & User & "')"
   oSaccoMaster.ExecuteThis (sql)
   
    'A.I /Vet  income 1.2
    lblclassname.Caption = "A.I /Vet Income"
    sql = " Set DateFormat DMY select (Select isnull(sum(Amount),0) From GLTRANSACTIONS where TransDate >='" & dtpStartDate & "' and TransDate<='" & dtpFinishDate & "' and CrAccNo='I203')- "
    sql = sql & " (Select isnull(sum(Amount),0) From GLTRANSACTIONS where TransDate >='" & dtpStartDate & "' and TransDate<='" & dtpFinishDate & "' and DrAccNo='I203' )AiIncome "
    Set rss = oSaccoMaster.GetRecordset(sql)
    If Not rss.EOF Then
    AIincome = rss.Fields(0)
    Else
    AIincome = 0
    End If
    
    sql = "set dateformat dmy Insert into Ag_Income(ClassNo, ClassName, AccNo, AccName, Amount, StartDate, EndDate, AuditId)"
    sql = sql & " Values ('1.2','A.I/Vet Income','1.3','SALES'," & AIincome & ",'" & dtpStartDate & "','" & dtpFinishDate & "','" & User & "')"
    oSaccoMaster.ExecuteThis (sql)
    
        'Other  income 1.21
    lblclassname.Caption = "Other Income"
    sql = " Set DateFormat DMY select (Select isnull(sum(Amount),0) From GLTRANSACTIONS where TransDate >='" & dtpStartDate & "' and TransDate<='" & dtpFinishDate & "' and CrAccNo='I204')- "
    sql = sql & " (Select isnull(sum(Amount),0) From GLTRANSACTIONS where TransDate >='" & dtpStartDate & "' and TransDate<='" & dtpFinishDate & "' and DrAccNo='I204' )OtherInc "
    Set rss = oSaccoMaster.GetRecordset(sql)
    If Not rss.EOF Then
    OtherInc = rss.Fields(0)
    Else
    OtherInc = 0
    End If
    
    sql = "set dateformat dmy Insert into Ag_Income(ClassNo, ClassName, AccNo, AccName, Amount, StartDate, EndDate, AuditId)"
    sql = sql & " Values ('1.21','KCEP Income','1.3','SALES'," & OtherInc & ",'" & dtpStartDate & "','" & dtpFinishDate & "','" & User & "')"
    oSaccoMaster.ExecuteThis (sql)
   
    TotalSales = cash + Credt + AIincome + OtherInc
    
      'Total sales 1.3
      lblclassname.Caption = "Total Sales"
    sql = "set dateformat dmy Insert into Ag_Income(ClassNo, ClassName, AccNo, AccName, Amount, StartDate, EndDate, AuditId)"
    sql = sql & " Values ('1.3','Total Sales','1.3','SALES'," & TotalSales & ",'" & dtpStartDate & "','" & dtpFinishDate & "','" & User & "')"
    oSaccoMaster.ExecuteThis (sql)
    
    prgStatus.Value = 30
    
      'Opening Stock 1.4
      lblclassname.Caption = "Opening stock"
    Startdate = DateSerial(year(dtpStartDate), month(dtpStartDate) - 1, 1)
    Enddate = DateSerial(year(dtpStartDate), month(dtpStartDate) - 1 + 1, 1 - 1)
   sql = "set dateformat dmy select isnull(sum(ClosingStockV),0)Cstock from  Ag_StationCStock where Transdate between '" & Startdate & "' and '" & Enddate & "' "
   Set rss = oSaccoMaster.GetRecordset(sql)
   If Not rss.EOF Then
    OpStock = rss.Fields(0)
   Else
   OpStock = 0
   End If
   sql = "set dateformat dmy Insert into Ag_Income(ClassNo, ClassName, AccNo, AccName, Amount, StartDate, EndDate, AuditId)"
   sql = sql & " Values ('1.4','Opening Stock','1.7','GOODS AVAILABLE FOR SALE'," & OpStock & ",'" & Startdate & "','" & Enddate & "','" & User & "')"
   oSaccoMaster.ExecuteThis (sql)
    
    'Purchases 1.5
    lblclassname.Caption = "Purchases"
   'sql = "set dateformat dmy select isnull(sum(Amount),0)Cstock from    Ag_Received where Transdate between '" & dtpStartDate & "' and '" & dtpFinishDate & "' "
    sql = " Set DateFormat DMY select (Select isnull(sum(Amount),0) From GLTRANSACTIONS where TransDate >='" & dtpStartDate & "' and TransDate<='" & dtpFinishDate & "' and DrAccNo='466-022')- "
    sql = sql & " (Select isnull(sum(Amount),0) From GLTRANSACTIONS where TransDate >='" & dtpStartDate & "' and TransDate<='" & dtpFinishDate & "' and CrAccNo='466-022' )AiIncome "
   Set rss = oSaccoMaster.GetRecordset(sql)
   If Not rss.EOF Then
    Purchases = rss.Fields(0)
   Else
   Purchases = 0
   End If
   
   prgStatus.Value = 45
   
   sql = "set dateformat dmy Insert into Ag_Income(ClassNo, ClassName, AccNo, AccName, Amount, StartDate, EndDate, AuditId)"
   sql = sql & " Values ('1.5','Purchases','1.7','GOODS AVAILABLE FOR SALE'," & Purchases & ",'" & dtpStartDate & "','" & dtpFinishDate & "','" & User & "')"
   oSaccoMaster.ExecuteThis (sql)
   
 'Carriage Inwards 1.6
' lblclassname.Caption = "Carriage Inwards"
  sql = " Set DateFormat DMY select (Select isnull(sum(Amount),0) From GLTRANSACTIONS where TransDate >='" & dtpStartDate & "' and TransDate<='" & dtpFinishDate & "' and DrAccNo='E227')- "
 sql = sql & " (Select isnull(sum(Amount),0) From GLTRANSACTIONS where TransDate >='" & dtpStartDate & "' and TransDate<='" & dtpFinishDate & "' and CrAccNo='E227' )AiIncome "
 Set rss = oSaccoMaster.GetRecordset(sql)
 If Not rss.EOF Then
 Cin = rss.Fields(0)
 Else
 Cin = 0
 End If
 sql = "set dateformat dmy Insert into Ag_Income(ClassNo, ClassName, AccNo, AccName, Amount, StartDate, EndDate, AuditId)"
 sql = sql & " Values ('1.6','Carriage Inwards','1.7','GOODS AVAILABLE FOR SALE'," & Cin & ",'" & dtpStartDate & "','" & dtpFinishDate & "','" & User & "')"
 oSaccoMaster.ExecuteThis (sql)
 
 prgStatus.Value = 60
  'Cost Of Goods Availble for sales 1.7
  lblclassname.Caption = "Cost Of Goods Availble for sales"
  CGFS = OpStock + Purchases + Cin
  sql = "set dateformat dmy Insert into Ag_Income(ClassNo, ClassName, AccNo, AccName, Amount, StartDate, EndDate, AuditId)"
 sql = sql & " Values ('1.7','C.G.A.F.S','1.7','COST OF GOODS AVAILABLE FOR SALE'," & CGFS & ",'" & dtpStartDate & "','" & dtpFinishDate & "','" & User & "')"
 oSaccoMaster.ExecuteThis (sql)

 
       'Closing Stock 1.8
       lblclassname.Caption = "Closing Stock"
   sql = "set dateformat dmy select isnull(sum(ClosingStockV),0)Cstock from  Ag_StationCStock where Transdate between '" & dtpStartDate & "' and '" & dtpFinishDate & "' "
   Set rss = oSaccoMaster.GetRecordset(sql)
   If Not rss.EOF Then
    CStock = rss.Fields(0)
   Else
   CStock = 0
   End If
   sql = "set dateformat dmy Insert into Ag_Income(ClassNo, ClassName, AccNo, AccName, Amount, StartDate, EndDate, AuditId)"
   sql = sql & " Values ('1.8','Closing Stock','1.8','CLOSING STOCK'," & CStock & ",'" & dtpStartDate & "','" & dtpFinishDate & "','" & User & "')"
   oSaccoMaster.ExecuteThis (sql)
   
    'Cost Of Sales 1.9
    lblclassname.Caption = "Cost of Sales"
 COS = CGFS - CStock
 sql = "set dateformat dmy Insert into Ag_Income(ClassNo, ClassName, AccNo, AccName, Amount, StartDate, EndDate, AuditId)"
 sql = sql & " Values ('1.9','Cost Of Sales','1.9','COST OF SALES'," & COS & ",'" & dtpStartDate & "','" & dtpFinishDate & "','" & User & "')"
 oSaccoMaster.ExecuteThis (sql)
 
 'Gross Profit 2.0
 lblclassname.Caption = "Gross Profit"
 Gprofit = TotalSales - COS
  sql = "set dateformat dmy Insert into Ag_Income(ClassNo, ClassName, AccNo, AccName, Amount, StartDate, EndDate, AuditId)"
 sql = sql & " Values ('2.0','Gross Profit','2.0','GROSS PROFIT'," & Gprofit & ",'" & dtpStartDate & "','" & dtpFinishDate & "','" & User & "')"
 oSaccoMaster.ExecuteThis (sql)
 
  prgStatus.Value = 80
  'Expenses 2.1
  sql = "SELECT Accno,NormalBal,glaccType,glaccname,glaccGroup FROM GLSETUP where AccNo<>'466-022' And AccNo<>'E227' And GlAccMainGroup='EXPENSES' and AccCategory='" & cbpdpt & "' ORDER BY ACCNO "
  Set rs2 = oSaccoMaster.GetRecordset(sql)
  With rs2
  While Not .EOF
        accno = rs2.Fields("Accno")
        GlAccName = rs2.Fields("glaccname")
        lblclassname.Caption = GlAccName
        lblclass = "2.1"
         sql = " Set DateFormat DMY select (Select isnull(sum(Amount),0) From GLTRANSACTIONS where TransDate >='" & dtpStartDate & "' and TransDate<='" & dtpFinishDate & "' and DrAccNo='" & accno & "')- "
         sql = sql & " (Select isnull(sum(Amount),0) From GLTRANSACTIONS where TransDate >='" & dtpStartDate & "' and TransDate<='" & dtpFinishDate & "' and CrAccNo='" & accno & "' )AiIncome "
        Set rss = oSaccoMaster.GetRecordset(sql)
            If Not rss.EOF Then
            Exp = rss.Fields(0)
            Else
            Exp = 0
            End If
          
            sql = "set dateformat dmy Insert into Ag_Income(ClassNo, ClassName, AccNo, AccName, Amount, StartDate, EndDate, AuditId)"
            sql = sql & " Values ('2.1','" & GlAccName & "','2.1','EXPENSES'," & Exp & ",'" & dtpStartDate & "','" & dtpFinishDate & "','" & User & "')"
            oSaccoMaster.ExecuteThis (sql)
            TotExp = TotExp + Exp
     .MoveNext
   Wend
 End With
 
 'Total Expenses 2.2
 sql = "set dateformat dmy Insert into Ag_Income(ClassNo, ClassName, AccNo, AccName, Amount, StartDate, EndDate, AuditId)"
sql = sql & " Values ('2.2','TOTAl EXPENSES','2.2','TOTAl EXPENSES'," & TotExp & ",'" & dtpStartDate & "','" & dtpFinishDate & "','" & User & "')"
oSaccoMaster.ExecuteThis (sql)

  'Net Profit 2.3
 lblclassname.Caption = "Net Profit"
 NetP = Gprofit - TotExp
  sql = "set dateformat dmy Insert into Ag_Income(ClassNo, ClassName, AccNo, AccName, Amount, StartDate, EndDate, AuditId)"
 sql = sql & " Values ('2.3','Net Profit','2.3','NET PROFIT'," & NetP & ",'" & dtpStartDate & "','" & dtpFinishDate & "','" & User & "')"
 oSaccoMaster.ExecuteThis (sql)
 
 prgStatus.Value = 100
  MsgBox "Comprehensive Income Statement Processed Successfully", vbInformation, Me.Caption
  Exit Sub
ErrorHandler:
  MsgBox err.description
End Sub

Private Sub Form_Load()
dtpStartDate = DateSerial(year(Get_Server_Date), month(Get_Server_Date), 1)
dtpFinishDate = DateSerial(year(Get_Server_Date), month(Get_Server_Date) + 1, 1 - 1)
End Sub
