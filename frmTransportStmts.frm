VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTransportStmts 
   Caption         =   "Transporters Statement"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7410
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox ports 
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      ItemData        =   "frmTransportStmts.frx":0000
      Left            =   5280
      List            =   "frmTransportStmts.frx":000A
      TabIndex        =   13
      Text            =   "LPT1"
      Top             =   720
      Width           =   2175
   End
   Begin VB.CheckBox ckprint 
      BackColor       =   &H00FFFF80&
      Caption         =   "Use LPT1 Printer"
      Height          =   345
      Left            =   5880
      TabIndex        =   12
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtTCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Transporter's Statements"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   5775
      Begin VB.OptionButton optAdvanceSlip 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Print Advance Slip"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1440
         Width           =   4095
      End
      Begin VB.OptionButton OptNormalStmt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Normal Statement (Use POS Printer)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   4095
      End
      Begin VB.OptionButton OptDetailedStmt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Detailed Statement (Use POS Printer)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   720
         Width           =   4095
      End
      Begin VB.OptionButton OptNormalA4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Normal Statement (Use Normal Printer (A4))"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   4095
      End
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      Caption         =   "Print"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1920
      TabIndex        =   1
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmClose 
      Appearance      =   0  'Flat
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      TabIndex        =   0
      Top             =   3240
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DTPStmts 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Format          =   165412865
      CurrentDate     =   40109
   End
   Begin VB.Label Label6 
      Caption         =   "Printer Port"
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Transport Code :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "End of Period :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Enter Transporter code and select end of period to print statement"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5970
   End
End
Attribute VB_Name = "frmTransportStmts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Enddate As Date
Dim Startdate As Date

Private Sub cmClose_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
Dim fso, chkPrinter, txtfile
    Dim ttt
     Dim escFeedAndCut As String
     Dim escNewLine As String
     Dim escUnerLineON As String
     Dim escUnerLineOnX2 As String
     Dim escUnerLineOff As String
     Dim escBoldOn As String
     Dim escBoldOff As String
     Dim escNegativeOn As String
     Dim escNegativeOff As String
     Dim esc8CpiOn As String
     Dim esc8CPiOff As String
     Dim esc16Cpi As String
     Dim esc20Cpi As String
     Dim escAlignLeft As String
     Dim escAlignCenter As String
     Dim escAlignRight As String
    Dim CummulKgs As Double
     
        escNewLine = Chr(10) '//New Line (LF Line Feed)
        escUnerLineON = Chr(27) + Chr(45) + Chr(1) '//Unerline on
        escUnerLineOnX2 = Chr(27) + Chr(45) + Chr(1) '//Unerline on X2
        escUnerLineOff = Chr(27) + Chr(45) + Chr(0) '//unerline off
        escBoldOn = Chr(27) + Chr(69) + Chr(1) '//Bold on
        escBoldOff = Chr(27) + Chr(69) + Chr(0) '//Bold off
        escNegativeOn = Chr(29) + Chr(66) + Chr(1) '//White on Black on
        escNegativeOff = Chr(29) + Chr(66) + Chr(0) '//white on
        esc8CpiOn = Chr(29) + Chr(33) + Chr(16) '//Font Size X2 on
        esc8CPiOff = Chr(29) + Chr(33) + Chr(0) '//Font size X2 off
        esc16Cpi = Chr(27) + Chr(77) + Chr(48) '//Font A - Normal Size
        esc20Cpi = Chr(27) + Chr(77) + Chr(49) '//Font B - Small Font
        escAlignLeft = Chr(27) + Chr(97) + Chr(48) '//Align text to the left
        escAlignCenter = Chr(27) + Chr(97) + Chr(49) '//Align text to the center
        escAlignRight = Chr(27) + Chr(97) + Chr(50) '//Align text to the right
        escFeedAndCut = Chr(29) + Chr(86) + Chr(65) '//Partial cut and feed
       
If txtTCode = "" Then
    MsgBox "Please enter Transporter code.", vbCritical
        txtTCode.SetFocus
    Exit Sub
End If

Startdate = DateSerial(year(DTPStmts), month(DTPStmts), 1)
Enddate = DateSerial(year(DTPStmts), month(DTPStmts) + 1, 1 - 1)

If optAdvanceSlip.value = True Then
'--Net amount as at date
    Set Rst1 = New ADODB.Recordset
        sql = "d_sp_PrintTransStmt '" & txtTCode & "','" & DTPStmts & "'"
    Set Rst1 = oSaccoMaster.GetRecordset(sql)
    
    If Rst1.EOF Then
    MsgBox "There is no current record."
        txtTCode.SetFocus
        Exit Sub
    End If
    
        ttt = ports 'LPT1,LPT2....
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error GoTo err
        
        'Set chkPrinter = fso.GetFile(ttt)
       
        
    Set txtfile = fso.CreateTextFile(ttt, True)
    txtfile.WriteLine escAlignCenter
    txtfile.WriteLine "Advance Slip"
    txtfile.WriteLine "" & cname & ""
    txtfile.WriteLine "........................................"
    txtfile.WriteLine escAlignLeft
    txtfile.WriteLine "Trans Code : " & txtTCode
    txtfile.WriteLine "Names : " & Rst1.Fields("TransName")
    txtfile.WriteLine "Issue Items/Services worth not more than"
    Dim rsg As New ADODB.Recordset
    Set rsg = oSaccoMaster.GetRecordset("SET  dateformat dmy  SELECT     SUM(Amount) + SUM(Subsidy) AS Gross   FROM  d_TransDetailed  WHERE     (Trans_Code = '" & txtTCode & "') AND (EndPeriod = '" & DTPStmts & "')")
    Dim Gro As Currency
    If Not IsNull(rsg.Fields(0)) Then Gro = CCur(rsg.Fields(0))
    Set rsg = oSaccoMaster.GetRecordset("SET dateformat dmy SELECT SUM(amount) FROM d_Transport_Deduc   WHERE     tdate_deduc BETWEEN '" & Startdate & "' AND '" & Enddate & "' AND transcode = '" & txtTCode & "'")
    Dim Ded As Currency
    If Not IsNull(rsg.Fields(0)) Then Ded = CCur(rsg.Fields(0))
    
    txtfile.WriteLine "Kshs. : " & Format((Gro - Ded), "#,##0.00") & ""
    txtfile.WriteLine "Sign"
    txtfile.WriteLine "___________________________"
    txtfile.WriteLine UCase(username)
    txtfile.Write "Date " & Format(Get_Server_Date, "dd/mm/yyyy")
    txtfile.WriteLine ", Time : " & Format(Get_Server_Date, "HH:MM:SS")
    txtfile.WriteLine "........................................"
    txtfile.WriteLine escFeedAndCut
    
    
    
End If

If OptNormalStmt.value = True Then
    Set Rst = New ADODB.Recordset
        sql = "d_sp_PrintTransStmt '" & txtTCode & "','" & Enddate & "'"
    Set Rst = oSaccoMaster.GetRecordset(sql)
If Rst.EOF Then
    MsgBox "There is no record in the payroll for Transporter with code " & txtTCode, vbInformation
        txtTCode.SetFocus
    Exit Sub
End If



        ttt = ports 'LPT1,LPT2....
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error GoTo err
        
'        Set chkPrinter = fso.GetFile(ttt)
       
        
        Set txtfile = fso.CreateTextFile(ttt, True)
    txtfile.WriteLine escAlignCenter
    txtfile.WriteLine "" & cname & ""
    txtfile.WriteLine "" & paddress & ""
    txtfile.WriteLine "" & town & ""
    txtfile.WriteLine "TRANSPORTER STATEMENT FOR " & UCase(Format(DTPStmts, "MMMM/YYYY"))
    txtfile.WriteLine escAlignLeft
    txtfile.WriteLine "........................................"
    txtfile.WriteLine "Trans Code :" & txtTCode
    txtfile.WriteLine "Tans Name :" & Rst!TransName
    txtfile.WriteLine "........................................"
    txtfile.WriteLine "Total Kgs Delivered :" & Format(Rst!QntySup, "#,##0.00" & " Kgs")
    txtfile.WriteLine "Gross Amount               Kshs: " & Format(Rst!Amnt, "#,##0.00") & ""
    txtfile.WriteLine "Subsidy                    Kshs: " & Format(Rst!subsidy, "#,##0.00") & ""
    txtfile.WriteLine "Net Gross                  Kshs: " & Format(Rst!GrossPay, "#,##0.00") & ""
   
    txtfile.Write escBoldOn
    txtfile.WriteLine "DEDUCTIONS"
    txtfile.WriteLine "-------------"
    txtfile.Write escBoldOff
    
    txtfile.WriteLine "Agrovet          Kshs: " & Format(Rst!agrovet, "#,##0.00") & ""
    txtfile.WriteLine "TM Shares        Kshs: " & Format(Rst!TMShares, "#,##0.00") & ""
    txtfile.WriteLine "H Shares         Kshs: " & Format(Rst!HShares, "#,##0.00") & ""
    txtfile.WriteLine "Advances         Kshs: " & Format(Rst!Advance, "#,##0.00") & ""
    txtfile.WriteLine "FSA              Kshs: " & Format(Rst!FSA, "#,##0.00") & ""
    txtfile.WriteLine "AI               Kshs: " & Format(Rst!AI, "#,##0.00") & ""
    txtfile.WriteLine "Others           Kshs: " & Format(Rst!Others, "#,##0.00") & ""
    txtfile.WriteLine "Total Deductions Kshs: " & Format(Rst!Totaldeductions, "#,##0.00") & ""
    txtfile.WriteLine "........................................"
    txtfile.WriteLine "NET PAY                    Kshs: " & Format((Rst!netpay), "#,##0.00") & ""
    txtfile.WriteLine "........................................"
    txtfile.WriteLine "BANK DETAILS"
    txtfile.WriteLine "-------------"
    txtfile.WriteLine "Bank Name :" & Rst!BankName & ""
    txtfile.WriteLine "Bank Branch :" & Rst!Branch
    txtfile.WriteLine "Account Number :" & Rst!accno

    txtfile.WriteLine "---------------------------------------"
    txtfile.WriteLine "        Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtfile.WriteLine "         " & motto & ""
    txtfile.WriteLine "---------------------------------------"
    txtfile.WriteLine escFeedAndCut
    txtfile.Close
'End If
Exit Sub
err: MsgBox err.description & " or There is no printer connected."
End If

'--Print detailed statement
If OptDetailedStmt.value = True Then

    
  ttt = ports 'LPT1,LPT2....
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error GoTo err
        
 '       Set chkPrinter = fso.GetFile(ttt)
       
        
        Set txtfile = fso.CreateTextFile(ttt, True)
    txtfile.WriteLine escAlignCenter
    txtfile.WriteLine "" & cname & ""
    txtfile.WriteLine "" & paddress & ""
    txtfile.WriteLine "" & town & ""
    txtfile.WriteLine "TRANSPORT DETAILED STATEMENT FOR " & UCase(Format(DTPStmts, "MMMM/YYYY"))
    txtfile.WriteLine escAlignLeft
    txtfile.WriteLine "........................................"
    txtfile.WriteLine "Trans Code :" & UCase(txtTCode)
    '---d_sp_TransDetStmt @Code varchar(35), @EPeriod varchar(12)
    Set rs = New ADODB.Recordset
    'sql = "d_sp_TransDetStmt '" & txtTCode & "','" & Enddate & "'"
    sql = "  set dateformat dmy  SELECT  dbo.d_TransDetailed.Trans_Code, dbo.d_Transporters.TransName, dbo.d_TransDetailed.SNo, dbo.d_TransDetailed.QNTY, dbo.d_TransDetailed.Amount , dbo.d_TransDetailed.subsidy From dbo.d_TransDetailed, dbo.d_Transporters WHERE     dbo.d_TransDetailed.Trans_Code = dbo.d_Transporters.TransCode AND   (dbo.d_TransDetailed.Trans_Code ='" & txtTCode & "' ) AND (dbo.d_TransDetailed.EndPeriod ='" & Enddate & "')"
    Set rs = oSaccoMaster.GetRecordset(sql)
    
    txtfile.WriteLine "Name :" & rs.Fields(1)
    txtfile.WriteLine "........................................"
    txtfile.WriteLine "SNo " & vbTab & "SUBSIDY" & vbTab & "QNTY" & vbTab & "PAYABLE"
    txtfile.WriteLine "........................................"
    
    
         
        While Not rs.EOF
        'MsgBox rs!QSupplied
        txtfile.WriteLine rs.Fields(2) & " " & vbTab & " " & Format(rs.Fields(5), "#,##0.00") & " " & vbTab & " " & Format(rs.Fields(3), "#,##0.0#") & " " & vbTab & " " & Format(rs.Fields(4), "#,##0.00") & " "
        'txtfile.WriteLine rs!PPU
         rs.MoveNext
        
        Wend
        
        
    Set Rst = New ADODB.Recordset
        'sql = "d_sp_PrintTransStmt '" & txtTCode & "','" & Enddate & "'"
        sql = "SET              dateformat dmy SELECT     dbo.d_TransportersPayRoll.Code, dbo.d_Transporters.TransName, dbo.d_TransportersPayRoll.QntySup, dbo.d_TransportersPayRoll.Amnt, dbo.d_TransportersPayRoll.Subsidy, dbo.d_TransportersPayRoll.GrossPay,  dbo.d_TransportersPayRoll.Agrovet, dbo.d_TransportersPayRoll.TMShares, dbo.d_TransportersPayRoll.FSA, dbo.d_TransportersPayRoll.HShares, dbo.d_TransportersPayRoll.Advance, dbo.d_TransportersPayRoll.AI,dbo.d_TransportersPayRoll.Others, dbo.d_TransportersPayRoll.Totaldeductions, dbo.d_TransportersPayRoll.NetPay,dbo.d_TransportersPayRoll.BankName , dbo.d_TransportersPayRoll.Branch, dbo.d_TransportersPayRoll.accno   FROM dbo.d_Transporters INNER JOIN dbo.d_TransportersPayRoll ON dbo.d_Transporters.TransCode = dbo.d_TransportersPayRoll.Code   WHERE     (dbo.d_TransportersPayRoll.Code = '" & txtTCode & "') AND (dbo.d_TransportersPayRoll.EndPeriod = '" & Enddate & "')"
    Set Rst = oSaccoMaster.GetRecordset(sql)
   txtfile.WriteLine "........................................"
   txtfile.WriteLine "TOTAL Kgs : " & Format(Rst!QntySup, "#,##0.00 ") & vbNewLine & "Amount " & vbTab & "" & vbTab & "Kshs " & Format(Rst!Amnt, "#,##0.00") & " " & vbNewLine & "Subsidy " & vbTab & "Kshs " & Format(Rst!subsidy, "#,##0.00") & vbNewLine & "Gross Amount " & vbTab & "Kshs " & Format((Rst!subsidy + Rst!Amnt), "#,##0.00")
   txtfile.WriteLine "........................................"
    txtfile.WriteLine escBoldOn
    txtfile.WriteLine "DEDUCTIONS"
    txtfile.WriteLine "........................................"
    txtfile.WriteLine escBoldOff
    Dim Gross As Currency
    Gross = Rst!subsidy + Rst!Amnt
    
  '--d_sp_PrintTransDeduc @Code varchar(35), @SDate varchar(12), @EDate varchar(12) AS
Set Rst = New ADODB.Recordset
sql = "d_sp_PrintTransDeduc " & txtTCode & ",'" & Startdate & "','" & Enddate & "'"
Set Rst = oSaccoMaster.GetRecordset(sql)
    

    txtfile.WriteLine "........................................"
    txtfile.WriteLine "DATE " & vbTab & "" & vbTab & "AMOUNT" & vbTab & "DESCRIPTION"
    txtfile.WriteLine "........................................"

    While Not Rst.EOF
        'MsgBox rs!QSupplied
        txtfile.WriteLine Rst.Fields(0) & " " & vbTab & " " & Format(Rst.Fields(2), "#,##0.00" & vbTab & " " & Rst.Fields(1) & " ")
        'txtfile.WriteLine rs!PPU
         Rst.MoveNext
        
        Wend
    Set Rst1 = New ADODB.Recordset
        sql = "SET dateformat dmy SELECT SUM(Amount) AS TOTAL From d_Transport_Deduc WHERE     (TransCode = '" & txtTCode & "') AND (TDate_Deduc BETWEEN '" & Startdate & "' AND '" & Enddate & "')"
    Set Rst1 = oSaccoMaster.GetRecordset(sql)
    txtfile.WriteLine "........................................"
    If Not IsNull(Rst1!Total) Then
    txtfile.WriteLine "Total Deductions Kshs: " & Format(Rst1!Total, "#,##0.00") & ""
    Else
    txtfile.WriteLine "Total Deductions Kshs: 0.00"
    End If
    txtfile.WriteLine "........................................"
    If Not IsNull(Rst1!Total) Then
    txtfile.WriteLine "NET PAY                   Kshs :" & Format(Gross - Rst1!Total, "#,##0.00") & ""
    Else
    txtfile.WriteLine "NET PAY                   Kshs :" & Format(Gross, "#,##0.00") & ""
    End If
    txtfile.WriteLine "-----------------------------------------"
    txtfile.WriteLine "BANK DETAILS"
    txtfile.WriteLine "-------------"
'    txtfile.WriteLine "Bank Name :" & Rst1!Bank & ""
'    txtfile.WriteLine "Bank Branch :" & Rst1!BBranch
'    txtfile.WriteLine "Account Number :" & Rst1!accountnumber
    txtfile.WriteLine "---------------------------------------"
    txtfile.WriteLine "        Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    txtfile.WriteLine "         " & motto & ""
    txtfile.WriteLine "---------------------------------------"
    txtfile.WriteLine escFeedAndCut
    txtfile.Close
    End If
    
    If OptNormalA4.value = True Then
    reportname = "d_StmtA4.rpt"
    STRFORMULA = "{d_Payroll.SNo}= " & txtTCode & " and {d_Payroll.MMonth }=" & month(DTPStmts) & " AND {d_Payroll.YYear} =" & year(DTPStmts)
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""

'    d_StmtA4
    End If
    txtTCode = ""
End Sub

Private Sub DTPStmts_Validate(Cancel As Boolean)
DTPStmts = DateSerial(year(DTPStmts), month(DTPStmts) + 1, 1 - 1)
End Sub

Private Sub Form_Load()
DTPStmts = Format(Get_Server_Date, "dd/mm/yyyy")
DTPStmts = DateSerial(year(DTPStmts), month(DTPStmts) + 1, 1 - 1)
End Sub
