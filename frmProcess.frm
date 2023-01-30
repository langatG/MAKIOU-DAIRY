VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProcess 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Process Payroll"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   8655
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdmidmonth 
      Caption         =   "MidMonth"
      Height          =   375
      Left            =   2160
      TabIndex        =   33
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Txtcreditedac 
      Height          =   285
      Left            =   2520
      TabIndex        =   30
      ToolTipText     =   "a/c to be credited"
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox lblcreditedac 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   29
      Top             =   5520
      Width           =   2775
   End
   Begin VB.TextBox Txtdebitedac 
      Height          =   285
      Left            =   2520
      TabIndex        =   28
      ToolTipText     =   "a/c to be debited"
      Top             =   6000
      Width           =   975
   End
   Begin VB.TextBox lbldebitedac 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   27
      Top             =   6000
      Width           =   2775
   End
   Begin VB.CommandButton Cmds1 
      Height          =   255
      Left            =   3480
      TabIndex        =   26
      Top             =   5550
      Width           =   255
   End
   Begin VB.CommandButton Cmds2 
      Height          =   255
      Left            =   3480
      TabIndex        =   25
      Top             =   6030
      Width           =   255
   End
   Begin VB.CommandButton CMDCFN 
      Caption         =   "Carry Forward Transport Deductions"
      Height          =   375
      Left            =   3120
      TabIndex        =   24
      Top             =   3600
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Include Compulsory Deductions"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6240
      TabIndex        =   23
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton cmdupdatebr 
      Caption         =   "Payroll Update"
      Height          =   375
      Left            =   7080
      TabIndex        =   22
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtsubsidy 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7440
      TabIndex        =   21
      Text            =   "1.25"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CheckBox chksubsidyc 
      Caption         =   "Add Subsidy Based on Current Month on Self Only "
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   1200
      Width           =   5175
   End
   Begin VB.CheckBox chksubsidyprev 
      Caption         =   "Add Subsidy Based on Previous Month on Self Only "
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   720
      Width           =   5175
   End
   Begin VB.TextBox txttotal 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1080
      TabIndex        =   17
      Top             =   8280
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker DTPto 
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   7320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   173670401
      CurrentDate     =   40555
   End
   Begin MSComCtl2.DTPicker DTPfrom 
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   7320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   173670401
      CurrentDate     =   40555
   End
   Begin VB.CommandButton cmdtotalmonthlyq 
      Caption         =   "Get The Kgs Periods Total"
      Height          =   375
      Left            =   960
      TabIndex        =   14
      Top             =   7800
      Width           =   2775
   End
   Begin VB.CommandButton cmdcompare 
      Caption         =   "Compare"
      Height          =   375
      Left            =   8760
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   8760
      TabIndex        =   12
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DTPEOD 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   173670401
      CurrentDate     =   40440
   End
   Begin VB.CommandButton cmdendofday 
      Caption         =   "End Of Day"
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton cmdCarry 
      Caption         =   "Carry Forward Suppliers Deductions"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtpProcess 
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   173604865
      CurrentDate     =   40214
   End
   Begin VB.CheckBox chkStop 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Stop Further Updates"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   1560
      UseMaskColor    =   -1  'True
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   0
      TabIndex        =   4
      Top             =   2040
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComCtl2.DTPicker dtpCarry 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   3000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16384
      Format          =   323354625
      CurrentDate     =   40214
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   135
      Left            =   0
      TabIndex        =   8
      Top             =   4200
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComCtl2.DTPicker previousp 
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   323354625
      CurrentDate     =   40214
   End
   Begin VB.Label Label51 
      Caption         =   "CR:"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   32
      Top             =   5550
      Width           =   735
   End
   Begin VB.Label Label101 
      Caption         =   "DR:"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   31
      Top             =   6030
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF80&
      Caption         =   "Process the total kilo for the day for seliing to the processor or any debtor."
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   4440
      Width           =   8895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "Carry Forward Deductions For Negative Net Pay For Period Ending"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   2760
      Width           =   8895
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Process Payrolls For the Period ending :"
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
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3630
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim midMonthDate As Date
    Dim EndMonth As Date
Private Sub cmdCarry_Click()
Dim Startdate1 As String, Enddate1 As String
Dim sno As String, Npay  As Currency, Totalnetpay As Currency, CFoward As Currency, DedAmount As Currency
Dim RsLessAmount As New ADODB.Recordset, RsDescription As New ADODB.Recordset
Dim desc As String, id As Double, Amnt As Currency, Flag As Double, TotalDed As Currency
Dim TDeductions As Currency, NetPay As Currency
Dim DeductCusor  As New ADODB.Recordset, RsTotalDed As New ADODB.Recordset

dtpCarry_Validate True

If dtpCarry > Get_Server_Date Then
    MsgBox "The records for the period ending " & dtpCarry & " has not been processed."
        dtpCarry.SetFocus
    Exit Sub
End If
   
Startdate = DateSerial(year(dtpCarry), month(dtpCarry), 1)
Enddate = DateSerial(year(dtpCarry), month(dtpCarry) + 1, 1 - 1)
Startdate1 = DateSerial(year(dtpCarry), month(dtpCarry) + 1, 1)
Enddate1 = DateSerial(year(dtpCarry), month(dtpCarry) + 2, 1 - 1)

ProgressBar1.Value = 0
sql = ""

Set RsLessAmount = oSaccoMaster.GetRecordset("set dateformat dmy SELECT     sno, Npay   From d_Payroll Where (NPay < 0) And endofperiod = '" & Enddate & "' order by npay ")
Do Until RsLessAmount.EOF
DoEvents
ProgressBar2.Max = RsLessAmount.RecordCount
ProgressBar2.Value = RsLessAmount.AbsolutePosition
frmProcess.Caption = RsLessAmount.Fields("sno")

NetPay = IIf(IsNull(RsLessAmount!Npay), 0, RsLessAmount!Npay)
Totalnetpay = NetPay

Set DeductCusor = oSaccoMaster.GetRecordset("Set Dateformat dmy SELECT S.[Description], S.Amount,S.[Id] From d_Supplier_deduc S INNER JOIN d_DCodes D on  D.[Description]=S.[Description] WHERE      (S.Date_Deduc  BETWEEN '" & Startdate & "' AND '" & Enddate & "' AND S.SNo = '" & RsLessAmount!sno & "' AND S.[Description] <> 'Transport' AND S.[Description] <> 'TCHP' AND S.[Description] <> 'HShares'  AND S.[Description] <> 'TMShares') AND S.Amount > 0  order by D.[DCode] DESC")
Do Until DeductCusor.EOF
 TotalDed = 0
desc = DeductCusor!description
Set RsTotalDed = oSaccoMaster.GetRecordset("Set Dateformat dmy select isnull(SUM(Amount),0) from d_supplier_deduc  where (Date_Deduc  BETWEEN '" & Startdate & "'  AND '" & Enddate & "' AND SNo ='" & RsLessAmount!sno & "' AND [Description]='" & DeductCusor!description & "' and id =" & DeductCusor!id & ") ")
 While Not RsTotalDed.EOF
             
             CFoward = IIf(Totalnetpay + DeductCusor!Amount >= 0, Totalnetpay * -1, DeductCusor!Amount)
             Totalnetpay = Totalnetpay + CFoward
             DedAmount = DeductCusor!Amount - CFoward
             
             Npay = Npay + RsTotalDed.Fields(0)
            NetPay = NetPay + DeductCusor!Amount
            NetPay = Totalnetpay
            
            If CFoward > 0 Then
            
            oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_supplier_Deduc SET Amount=" & DedAmount & ",Remarks='C/F '+CONVERT(VARCHAR(150), (" & CFoward & "))  WHERE SNo='" & RsLessAmount!sno & "' AND [Id]=" & DeductCusor!id & "")
            oSaccoMaster.ExecuteThis ("Set Dateformat dmy INSERT INTO d_Supplier_Deduc (SNo, Date_Deduc,[Description],Amount,StartDate,enddate, AuditID,Remarks) values ('" & RsLessAmount!sno & "', '" & Startdate1 & "','" & DeductCusor!description & "'," & ((CFoward)) & ",'" & Startdate1 & "','" & Enddate1 & "','" & User & "','Brought Forward')")
            
            If UCase(Trim(desc)) = "AGROVET" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_Payroll SET NPay=" & NetPay & ",TDeductions= TDeductions - (" & CFoward & "),Agrovet=Agrovet - (" & CFoward & ") WHERE SNo='" & RsLessAmount!sno & "' AND endofperiod = '" & Enddate & "'")
            End If '"AGROVET"
            
            If UCase(Trim(desc)) = "LEPESA LOAN" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_Payroll SET NPay=" & NetPay & ",TDeductions=TDeductions - (" & CFoward & "),FSA=FSA - (" & CFoward & ") WHERE SNo='" & RsLessAmount!sno & "' AND endofperiod ='" & Enddate & "'")
            End If '"LOANS"
            
            If UCase(Trim(desc)) = "AI" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_Payroll SET NPay=" & NetPay & ",TDeductions= TDeductions -(" & CFoward & "),AI=AI - (" & CFoward & ") WHERE SNo='" & RsLessAmount!sno & "' AND endofperiod ='" & Enddate & "'")
            End If '"AI"
            
            If UCase(Trim(desc)) = "OTHERS" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_Payroll SET NPay=" & NetPay & ",TDeductions= TDeductions - (" & CFoward & "),Others=Others -  (" & CFoward & ") WHERE SNo='" & RsLessAmount!sno & "' AND endofperiod = '" & Enddate & "'")
            End If '"OTHERS"
            
            If UCase(Trim(desc)) = "ADVANCE" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_Payroll SET NPay=" & NetPay & ",TDeductions=TDeductions -(" & CFoward & "),Advance=Advance - (" & CFoward & ") WHERE SNo='" & RsLessAmount!sno & "'  AND endofperiod ='" & Enddate & "'")
            End If ' "ADVANCE"
            
            If UCase(Trim(desc)) = "NHIF" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_Payroll SET NPay=" & NetPay & ",TDeductions=TDeductions -(" & CFoward & "),NHIF=NHIF - (" & CFoward & ") WHERE SNo='" & RsLessAmount!sno & "'  AND endofperiod ='" & Enddate & "'")
            End If ' "NHIF"
            
            If UCase(Trim(desc)) = "ECF" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_Payroll SET NPay=" & NetPay & ",TDeductions=TDeductions -(" & CFoward & "),ECF=ECF - (" & CFoward & ") WHERE SNo='" & RsLessAmount!sno & "'  AND endofperiod ='" & Enddate & "'")
            End If ' "ECF"
            
            If UCase(Trim(desc)) = "LEPESA SHARES" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_Payroll SET NPay=" & NetPay & ",TDeductions=TDeductions -(" & CFoward & "),LSHARES=LSHARES - (" & CFoward & ") WHERE SNo='" & RsLessAmount!sno & "'  AND endofperiod ='" & Enddate & "'")
            End If ' "LEPESA SHARES"
            
            If UCase(Trim(desc)) = "LOAN SAVINGS" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_Payroll SET NPay=" & NetPay & ",TDeductions=TDeductions -(" & CFoward & "),LSAV=LSAV - (" & CFoward & ") WHERE SNo='" & RsLessAmount!sno & "'  AND endofperiod ='" & Enddate & "'")
            End If ' "LOAN SAVINGS"
            
            If UCase(Trim(desc)) = "PREPAYMENTS" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_Payroll SET NPay=" & NetPay & ",TDeductions=TDeductions -(" & CFoward & "),Prepay=Prepay - (" & CFoward & ") WHERE SNo='" & RsLessAmount!sno & "'  AND endofperiod ='" & Enddate & "'")
            End If ' "PREPAYMENTS"
            
            If UCase(Trim(desc)) = "WATER BILL" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_Payroll SET NPay=" & NetPay & ",TDeductions=TDeductions -(" & CFoward & "),Water=water - (" & CFoward & ") WHERE SNo='" & RsLessAmount!sno & "'  AND endofperiod ='" & Enddate & "'")
            End If ' "WATER BILL"
            
            If UCase(Trim(desc)) = "SILAGE" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_Payroll SET NPay=" & NetPay & ",TDeductions=TDeductions -(" & CFoward & "),Silage=Silage - (" & CFoward & ") WHERE SNo='" & RsLessAmount!sno & "'  AND endofperiod ='" & Enddate & "'")
            End If ' "SILAGE"
            If UCase(Trim(desc)) = "INSURANCE" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_Payroll SET NPay=" & NetPay & ",TDeductions=TDeductions -(" & CFoward & "),Insurance=Insurance - (" & CFoward & ") WHERE SNo='" & RsLessAmount!sno & "'  AND endofperiod ='" & Enddate & "'")
            End If ' "INSURANCE"


GoTo maritim

Else

End If
maritim:

RsTotalDed.MoveNext
            Wend

 DeductCusor.MoveNext
 Loop
 
 '//loop next description
' RsDescription.MoveNext
 
 'Loop
RsLessAmount.MoveNext
Loop


oSaccoMaster.ExecuteThis ("DELETE FROM D_Supplier_Deduc WHERE  [Description]='' AND Amount=0")

'''*************** Brought forwards and supplier didnt supplier milk this month
''
''Set RsTotalDed = oSaccoMaster.GetRecordset("set dateformat dmy select sno,[Description], Amount,[Id] from d_supplier_deduc  where (Date_Deduc  BETWEEN '" & Startdate & "'  AND '" & Enddate & "') " _
''    & " and Amount<>0 and sno not in  (SELECT sno   From d_Payroll where endofperiod = '" & Enddate & "') order by sno")
'' With RsTotalDed
''     While Not RsTotalDed.EOF
''      frmProcess.Caption = !sno
''
''        CFoward = IIf(IsNull(.Fields(2)), 0, .Fields(2))
''        oSaccoMaster.ExecuteThis ("UPDATE d_supplier_Deduc SET Amount=0,Remarks='C/F '+CONVERT(VARCHAR(150), (" & CFoward & "))  WHERE SNo='" & !sno & "' AND [Id]=" & !id & "")
''        oSaccoMaster.ExecuteThis ("INSERT INTO d_Supplier_Deduc (SNo, Date_Deduc,[Description],Amount,StartDate,enddate, AuditID,Remarks) values ('" & !sno & "', '" & Startdate1 & "','" & !description & "'," & ((CFoward)) & ",'" & Startdate1 & "','" & Enddate1 & "','" & user & "','Brought Forward')")
''
''        .MoveNext
''     Wend
'' End With
 
 sql = ""
sql = "set dateformat dmy select distinct sno from d_supplier_deduc where Date_Deduc  BETWEEN '" & Startdate & "'  AND '" & Enddate & "'  order by sno"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
ProgressBar2.Max = rs.RecordCount
While Not rs.EOF
DoEvents

ProgressBar2.Value = rs.AbsolutePosition

sno = rs.Fields(0)
sql = "select sno from d_payroll where  sno='" & sno & "' and mmonth=" & month(Enddate) & " and yyear=" & year(Enddate) & "  order by sno"
Set Rst = oSaccoMaster.GetRecordset(sql)
    If Rst.EOF Then
        Set RsTotalDed = oSaccoMaster.GetRecordset("set dateformat dmy select sno,[Description], Amount,[Id] from d_supplier_deduc  where sno='" & sno & "' and (Date_Deduc  BETWEEN '" & Startdate & "'  AND '" & Enddate & "') " _
    & " and Amount<>0 order by Date_Deduc,id")
        With RsTotalDed
            While Not RsTotalDed.EOF
             frmProcess.Caption = !sno
             
               CFoward = IIf(IsNull(.Fields(2)), 0, .Fields(2))
               oSaccoMaster.ExecuteThis ("UPDATE d_supplier_Deduc SET Amount=0,Remarks='C/F '+CONVERT(VARCHAR(150), (" & CFoward & "))  WHERE SNo='" & !sno & "' AND [Id]=" & !id & "")
               oSaccoMaster.ExecuteThis ("INSERT INTO d_Supplier_Deduc (SNo, Date_Deduc,[Description],Amount,StartDate,enddate, AuditID,Remarks) values ('" & !sno & "', '" & Startdate1 & "','" & !description & "'," & ((CFoward)) & ",'" & Startdate1 & "','" & Enddate1 & "','" & User & "','Brought Forward')")
              
               .MoveNext
            Wend
        End With
        
    End If
    frmProcess.Caption = sno
    rs.MoveNext
Wend
End If
MsgBox "Records saved successful!"




End Sub



Private Sub CMDCFN_Click()
dtpCarry_Validate True

If dtpCarry > Get_Server_Date Then
    MsgBox "The records for the period ending " & dtpCarry & " has not been processed."
        dtpCarry.SetFocus
    Exit Sub
End If
   
Startdate = DateSerial(year(dtpCarry), month(dtpCarry), 1)
Enddate = DateSerial(year(dtpCarry), month(dtpCarry) + 1, 1 - 1)

ProgressBar1.Value = 0
Dim Startdate1 As String
Dim Enddate1 As String
Dim sno As String
Dim Npay  As Currency
Dim RsLessAmount As New ADODB.Recordset
Startdate1 = DateSerial(year(dtpCarry), month(dtpCarry) + 1, 1)
Enddate1 = DateSerial(year(dtpCarry), month(dtpCarry) + 1, 28)
sql = ""
sql = "SET dateformat DMY SELECT     distinct d_TransportersPayRoll.code, d_TransportersPayRoll.NetPay      "
sql = sql & " FROM         d_TransportersPayRoll   inner join d_transport on "
 sql = sql & " d_TransportersPayRoll.code=d_transport.trans_code WHERE     (d_TransportersPayRoll.NetPay < 0)"
sql = sql & " AND d_TransportersPayRoll.endperiod = '" & Enddate & "'    and  d_transport.active=1 "
sql = sql & "  order by code"

Set RsLessAmount = oSaccoMaster.GetRecordset(sql)

sql = ""
Do Until RsLessAmount.EOF
DoEvents
ProgressBar2.Max = RsLessAmount.RecordCount
ProgressBar2.Value = RsLessAmount.AbsolutePosition
frmProcess.Caption = RsLessAmount.Fields("code")


Dim desc As String
Dim id  As Double
Dim Amnt As Currency
Dim Flag As Double
Dim TotalDed As Currency
Dim Totalnetpay, CFoward, DedAmount As Currency
Dim TDeductions As Currency
Dim NetPay As Currency
Dim RsDescription As New ADODB.Recordset
'--SET Flag = 1

NetPay = IIf(IsNull(RsLessAmount!NetPay), 0, RsLessAmount!NetPay)
Totalnetpay = NetPay

Dim DeductCusor  As New ADODB.Recordset

Set DeductCusor = oSaccoMaster.GetRecordset("Set Dateformat dmy SELECT T.[Description], T.Amount,T.[Id] From d_Transport_Deduc T INNER JOIN d_DCodes D on D.[Description]=T.[Description]  WHERE(T.TDate_Deduc   BETWEEN '" & Startdate & "' AND '" & Enddate & "' AND T.transcode  = '" & RsLessAmount!code & "'  AND T.[Description] <> 'TCHP' AND T.[Description] <> 'HShares'  AND T.[Description] <> 'TMShares') AND T.Amount > 0  order by D.[DCode] DESC")
Do Until DeductCusor.EOF


 TotalDed = 0
 Dim RsTotalDed As New ADODB.Recordset
 
Set RsTotalDed = oSaccoMaster.GetRecordset("Set Dateformat dmy select isnull(SUM(Amount),0) from d_Transport_Deduc  where (TDate_Deduc   BETWEEN '" & Startdate & "'  AND '" & Enddate & "' AND transcode ='" & RsLessAmount!code & "' AND [Description]='" & DeductCusor!description & "' and id =" & DeductCusor!id & ") ")
 While Not RsTotalDed.EOF
             
             CFoward = IIf(Totalnetpay + DeductCusor!Amount >= 0, Totalnetpay * -1, DeductCusor!Amount)
             Totalnetpay = Totalnetpay + CFoward
             DedAmount = DeductCusor!Amount - CFoward
             
             Npay = Npay + RsTotalDed.Fields(0)
            NetPay = NetPay + DeductCusor!Amount
            NetPay = Totalnetpay
            
            If CFoward > 0 Then
            'description='C/F '+CONVERT(VARCHAR(150), (" & CDbl(Rs1.Fields(1)) - NetPay & "))
            
            oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_Transport_Deduc SET Amount=" & DedAmount & ",Remarks='C/F '+CONVERT(VARCHAR(150), (" & CFoward & "))  WHERE transcode='" & RsLessAmount!code & "' AND [Id]=" & DeductCusor!id & "")
            oSaccoMaster.ExecuteThis ("Set Dateformat dmy INSERT INTO d_Transport_Deduc (transcode,  tdate_deduc,[Description],Amount,StartDate,enddate,AuditID,Remarks) values ('" & RsLessAmount!code & "', '" & Startdate1 & "','" & DeductCusor!description & "'," & ((CFoward)) & ",'" & Startdate1 & "','" & Enddate1 & "','" & User & "','Brought Forward')")
            
            If UCase(Trim(DeductCusor!description)) = "AGROVET" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_TransportersPayRoll  SET NetPay=" & NetPay & ",Totaldeductions= Totaldeductions - (" & CFoward & "),Agrovet=Agrovet - (" & CFoward & ") WHERE code='" & RsLessAmount!code & "' AND endperiod = '" & Enddate & "'")
            End If '"AGROVET"
            
            If UCase(Trim(DeductCusor!description)) = "ADVANCE" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_TransportersPayRoll SET NetPay=" & NetPay & ",Totaldeductions=Totaldeductions -(" & CFoward & "),Advance=Advance - (" & CFoward & ") WHERE code='" & RsLessAmount!code & "'  AND endperiod ='" & Enddate & "'")
            End If ' "ADVANCE"
            
            If UCase(Trim(DeductCusor!description)) = "LEPESA LOAN" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_TransportersPayRoll SET NetPay=" & NetPay & ",Totaldeductions=Totaldeductions - (" & CFoward & "),FSA=FSA - (" & CFoward & ") WHERE code='" & RsLessAmount!code & "' AND endperiod ='" & Enddate & "'")
            End If '"LEPESA LOAN"
            
            If UCase(Trim(DeductCusor!description)) = "AI" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_TransportersPayRoll SET NetPay=" & NetPay & ",Totaldeductions= Totaldeductions -(" & CFoward & "),AI=AI - (" & CFoward & ") WHERE code='" & RsLessAmount!code & "' AND endperiod ='" & Enddate & "'")
            End If '"AI"
            
            If UCase(Trim(DeductCusor!description)) = "NHIF" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_TransportersPayRoll SET NetPay=" & NetPay & ",Totaldeductions= Totaldeductions -(" & CFoward & "),NHIF=NHIF - (" & CFoward & ") WHERE code='" & RsLessAmount!code & "' AND endperiod ='" & Enddate & "'")
            End If '"NHIF"
            
            If UCase(Trim(DeductCusor!description)) = "ECF" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_TransportersPayRoll SET NetPay=" & NetPay & ",Totaldeductions= Totaldeductions -(" & CFoward & "),ECF=ECF - (" & CFoward & ") WHERE code='" & RsLessAmount!code & "' AND endperiod ='" & Enddate & "'")
            End If '"ECF"
            
            If UCase(Trim(DeductCusor!description)) = "LEPESA SHARES" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_TransportersPayRoll SET NetPay=" & NetPay & ",Totaldeductions= Totaldeductions -(" & CFoward & "),LSHARES=LSHARES - (" & CFoward & ") WHERE code='" & RsLessAmount!code & "' AND endperiod ='" & Enddate & "'")
            End If '"LEPESA SHARES"
            
            If UCase(Trim(DeductCusor!description)) = "LOAN SAVINGS" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_TransportersPayRoll SET NetPay=" & NetPay & ",Totaldeductions= Totaldeductions -(" & CFoward & "),LSAV=LSAV - (" & CFoward & ") WHERE code='" & RsLessAmount!code & "' AND endperiod ='" & Enddate & "'")
            End If '"LOAN SAVINGS"
            
            If UCase(Trim(DeductCusor!description)) = "PREPAYMENTS" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_TransportersPayRoll SET NetPay=" & NetPay & ",Totaldeductions= Totaldeductions -(" & CFoward & "),Prepay=Prepay - (" & CFoward & ") WHERE code='" & RsLessAmount!code & "' AND endperiod ='" & Enddate & "'")
            End If '"PREPAYMENTS"
            
            If UCase(Trim(DeductCusor!description)) = "WATER BILL" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_TransportersPayRoll SET NetPay=" & NetPay & ",Totaldeductions= Totaldeductions -(" & CFoward & "),Water=Water - (" & CFoward & ") WHERE code='" & RsLessAmount!code & "' AND endperiod ='" & Enddate & "'")
            End If '"WATER BILL"
            
            If UCase(Trim(DeductCusor!description)) = "SILAGE" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_TransportersPayRoll SET NetPay=" & NetPay & ",Totaldeductions= Totaldeductions -(" & CFoward & "),Silage=Silage - (" & CFoward & ") WHERE code='" & RsLessAmount!code & "' AND endperiod ='" & Enddate & "'")
            End If '"SILAGE"
            
            If UCase(Trim(DeductCusor!description)) = "FUEL" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_TransportersPayRoll SET NetPay=" & NetPay & ",Totaldeductions= Totaldeductions -(" & CFoward & "),Fuel=Fuel - (" & CFoward & ") WHERE code='" & RsLessAmount!code & "' AND endperiod ='" & Enddate & "'")
            End If '"FUEL"
            
            If UCase(Trim(DeductCusor!description)) = "MILK REJETCS" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_TransportersPayRoll SET NetPay=" & NetPay & ",Totaldeductions= Totaldeductions -(" & CFoward & "),MilkR=MilkR - (" & CFoward & ") WHERE code='" & RsLessAmount!code & "' AND endperiod ='" & Enddate & "'")
            End If '"MILK REJETCS"
            
            If UCase(Trim(DeductCusor!description)) = "MILK VARIANCE" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_TransportersPayRoll SET NetPay=" & NetPay & ",Totaldeductions= Totaldeductions -(" & CFoward & "),MilkV=MilkV - (" & CFoward & ") WHERE code='" & RsLessAmount!code & "' AND endperiod ='" & Enddate & "'")
            End If '"MILK VARIANCE"
            If UCase(Trim(DeductCusor!description)) = "INSURANCE" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_TransportersPayRoll SET NetPay=" & NetPay & ",Totaldeductions= Totaldeductions -(" & CFoward & "),Insurance=Insurance - (" & CFoward & ") WHERE code='" & RsLessAmount!code & "' AND endperiod ='" & Enddate & "'")
            End If '"INSURANCE"
            
            If UCase(Trim(DeductCusor!description)) = "OTHERS" Then
                oSaccoMaster.ExecuteThis ("Set Dateformat dmy UPDATE d_TransportersPayRoll SET NetPay=" & NetPay & ",Totaldeductions= Totaldeductions - (" & CFoward & "),Others=Others -  (" & CFoward & ") WHERE code='" & RsLessAmount!code & "' AND endperiod = '" & Enddate & "'")
            End If '"OTHERS"

GoTo maritim

Else


End If
maritim:
RsTotalDed.MoveNext
            Wend

 DeductCusor.MoveNext
 Loop

RsLessAmount.MoveNext
Loop

oSaccoMaster.ExecuteThis ("DELETE FROM d_Transport_Deduc  WHERE  [Description]='' AND Amount=0")
'*************** Brought forwards and Transporter didnt supplier milk this month
          
Set RsTotalDed = oSaccoMaster.GetRecordset("set dateformat dmy select transcode,[Description], Amount,[Id] from d_Transport_Deduc  where (TDate_Deduc  BETWEEN '" & Startdate & "'  AND '" & Enddate & "') " _
    & " and Amount<>0 and  transcode not in  (SELECT code   From d_TransportersPayRoll where endperiod = '" & Enddate & "') order by transcode")
 With RsTotalDed
     While Not RsTotalDed.EOF
        frmProcess.Caption = !Transcode
        CFoward = IIf(IsNull(.Fields(2)), 0, .Fields(2))
        
        oSaccoMaster.ExecuteThis ("UPDATE d_Transport_Deduc SET Amount=0,Remarks='C/F '+CONVERT(VARCHAR(150), (" & CFoward & "))  WHERE transcode='" & !Transcode & "' AND [Id]=" & !id & "")
        oSaccoMaster.ExecuteThis ("INSERT INTO d_Transport_Deduc (transcode,  tdate_deduc,[Description],Amount,StartDate,enddate,AuditID,Remarks) values ('" & !Transcode & "', '" & Startdate1 & "','" & !description & "'," & ((CFoward)) & ",'" & Startdate1 & "','" & Enddate1 & "','" & User & "','Brought Forward')")
    
        .MoveNext
     Wend
 End With

MsgBox "Records saved successful!"
End Sub

Private Sub cmdcompare_Click()
On Error GoTo ErrorHandler

Set rs = oSaccoMaster.GetRecordset("SELECT     SNo, AccNo, Bcode, BBranch  FROM         d_Suppliers  where sno>=4242 and sno<=4469 ORDER BY SNo")
While Not rs.EOF
DoEvents

Set Rst = oSaccoMaster.GetRecordset("SELECT     sno,accno,bank,branch,idno  FROM         Sheet11 where sno=" & rs.Fields(0) & "")
If Not Rst.EOF Then

If Trim(rs.Fields(1)) <> Trim(Rst.Fields(1)) Then
sql = ""
sql = "update d_suppliers set ACCNO='" & Rst.Fields(1) & "',BCODE='" & Rst.Fields(2) & "',BBRANCH='" & Rst.Fields(3) & "',IDNO='" & Rst.Fields(4) & "' where sno=" & rs.Fields(0) & ""
oSaccoMaster.ExecuteThis (sql)
End If
End If
rs.MoveNext
Wend


Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdendofday_Click()
On Error GoTo ErrorHandler
Dim totalkilo As Double
Dim dipping As Double

'get the total kilo for the day
  Set rs = New ADODB.Recordset
    sql = "d_sp_DailyTotal '" & DTPEOD & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    totalkilo = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
    Else
    totalkilo = 0
    End If
    '//check if milk is available
    If totalkilo = 0 Then
    MsgBox ("No milk has been received for this day; kindly choose another date"), vbInformation, "EASYMA=END OF DAY"
    Exit Sub
    End If
        If Txtdebitedac = "" Then
            MsgBox "please input the account to be debited"
            Txtdebitedac.SetFocus
        Exit Sub
        End If
        If Txtcreditedac = "" Then
            MsgBox "please input the account to be credited"
            Txtcreditedac.SetFocus
        Exit Sub
        End If
    
    Dim dipss, Intake As Double, Price As Double
    dipss = 0
    Intake = 0
    Price = 0
        
        sql = "select price from d_price"
        Set rs = oSaccoMaster.GetRecordset(sql)
        If Not rs.EOF Then
        Price = rs!Price
        End If
        
        dipping = totalkilo
    '//Check whether there was close of the same day before
    sql = "set dateformat dmy  SELECT   dipping  From d_dispatch where  descrip='Intake' and transdate = '" & DTPEOD & "' and dipping>0 and Intake>0 "
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
         If MsgBox(DTPEOD & " Has Already Been Closed", vbInformation + vbYesNo, "Are You Sure You Want To Save The Changes ") = vbYes Then
        
           sql = "set dateformat dmy select isnull(Sum(intake),0),isnull(sum(dipping),0) from  d_dispatch where transdate='" & DTPEOD & "'"
           Set rs = oSaccoMaster.GetRecordset(sql)
            If Not rs.EOF Then
                Intake = (rs.Fields(0))
                dipss = rs.Fields(1)
              If totalkilo > Intake Then
                dipping = totalkilo - Intake
                Else
                dipping = totalkilo - Intake
              End If
            End If
         Else
         Exit Sub
         End If
    End If
    'validate the milk available for intake
        sql = ""
        sql = "set dateformat dmy INSERT INTO d_dispatch (Transdate, descrip, Intake, dipping, dispatch, auditid, auditdate)values ('" & DTPEOD.Value & "','Intake'," & dipping & "," & dipping & ",0,'" & User & "','" & Get_Server_Date & "') "
        'sql = "set dateformat dmy insert_d_dispatch '" & DTPEOD.value & "','Intake'," & totalkilo & "," & dipping + totalkilo & ",0,'" & User & "','" & Get_Server_Date & "'"
        oSaccoMaster.ExecuteThis (sql)
      
         NewTransaction dipping * Price, transdate, "milk purchase"
     If dipping > 0 Then
        If Not Save_GLTRANSACTION(DTPEOD, dipping * Price, Txtdebitedac, Txtcreditedac, "milk purchase", "eod" & dipping, User, ErrorMessage, "close of day ", 1, 1, "intake" & Get_Server_Date, transactionNo, "") Then
            If ErrorMessage <> "" Then
            MsgBox err.description, vbInformation, "end of day"
            End If
        End If
     Else
        If Not Save_GLTRANSACTION(DTPEOD, dipping * Price * -1, Txtcreditedac, Txtdebitedac, "milk purchase", "eod" & dipping, User, ErrorMessage, "close of day ", 1, 1, "intake" & Get_Server_Date, transactionNo, "") Then
            If ErrorMessage <> "" Then
            MsgBox err.description, vbInformation, "end of day"
            End If
        End If
     End If
MsgBox "Close of day sucessfully updated"
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmdmidmonth_Click()
Dim rsdeduction As New Recordset
Dim Yr As Integer, ym As Integer
Dim currdate As Date
currdate = Format(Get_Server_Date, "dd/mm/yyyy")
Startdate = DateSerial(year(dtpProcess), month(dtpProcess), 1)
Enddate = DateSerial(year(dtpProcess), month(dtpProcess) + 1, 1 - 1)
EndMonth = Enddate
midMonthDate = Format(dtpProcess, "DD/MM/YYYY") 'DateSerial(year(Enddate), month(Enddate), 15)
ProgressBar1.Value = 0
Yr = year(dtpProcess)
ym = DateDiff("d", Startdate, currdate)

'====== Check Reprocessed midMonth ==================
Set rsdeduction = oSaccoMaster.GetRecordset("set dateformat dmy Select * from d_supplier_deduc where  Date_Deduc='" & midMonthDate & "' and Description='WEEKLY' and Remarks='payment'")
 If Not rsdeduction.EOF Then
  If MsgBox("Weekly Payroll for the period has already been Processed,Are You Sure You want To Reprocess", vbQuestion + vbYesNo) = vbNo Then
   Exit Sub
 End If
 End If
 
  If MsgBox("Are you sure you want to Process Weekly Payroll", vbQuestion + vbYesNo) = vbNo Then
   Exit Sub
 End If

 
    sql = "set dateformat dmy Delete from d_supplier_deduc where Date_Deduc='" & midMonthDate & "' and Description='WEEKLY' and Remarks='payment'"
    oSaccoMaster.ExecuteThis (sql)
 
 addomittedentried

'//update deduction before anything else before running the payroll start here
Dim rshast1 As New ADODB.Recordset, descrip As String, remark As String, amt As Double, sno As Long
 Set rshast1 = oSaccoMaster.GetRecordset("set dateformat dmy select * from d_supplier_standingorder  where enddate>='" & Enddate & "' order by sno")
While Not rshast1.EOF
DoEvents
sno = rshast1.Fields("sno")
remark = Trim(rshast1.Fields("description"))
'remark = rshast.Fields("remarks")
amt = rshast1.Fields("amount")
sql = ""
sql = "select description,remarks from d_supplier_deduc where sno='" & sno & "' and description ='Others' and remarks='" & remark & "' and month(date_deduc)=" & month(Enddate) & " and year(date_deduc)=" & year(Enddate) & ""
Set Rst = oSaccoMaster.GetRecordset(sql)
        If Rst.EOF Then
        frmProcess.Caption = sno
            If amt > 0 Then
            sql = "d_sp_SupplierDeduct " & sno & ",'" & Enddate & "','Others'," & amt & ",'" & Startdate & "','" & Enddate & "'," & year(Enddate) & ",'" & User & "','" & remark & "'"
            oSaccoMaster.ExecuteThis (sql)
            End If
            Else
            
        End If
        amt = 0
rshast1.MoveNext
Wend
'**************end here
ProgressBar1.Value = 20
sql = ("d_sp_PresetDeductAssignWeekly '" & Startdate & "','" & midMonthDate & "'," & Yr & ",'" & User & "'")
oSaccoMaster.ExecuteThis (sql)
ProgressBar1.Value = 30
'Payroll update
sql = " SET DATEFORMAT DMY SELECT Sno  From d_Payroll WHERE     YYEAR = '" & year(Enddate) & "' AND MMONTH='" & month(Enddate) & "' AND SNO IN(SELECT SNO FROM D_SUPPLIERS WHERE TRADER=1) order by sno"
Set rs = oSaccoMaster.GetRecordset(sql)
With rs
    If Not .EOF Then
    While Not .EOF
    DoEvents
    sno = !sno
    frmProcess.Caption = sno
   
        Dim RsShares As New Recordset
        Dim rspu As New Recordset
   
        DoEvents
            Dim amt1 As Double
            Dim amt2 As Double
            Dim sp As Double
   
            sql = "SET dateformat dmy SELECT SUM(QSupplied) AS QNTY, SUM(pAmount) AS GrossPay " _
            & "From d_Milkintake WHERE (TransDate BETWEEN '" & Startdate & "'  AND '" & midMonthDate & "' AND SNo ='" & sno & "')"
                   
            Set rs2 = oSaccoMaster.GetRecordset(sql)
            Dim GPay As Double
            GPay = 0
            
            If Not rs2.EOF Then
                GPay = IIf(IsNull(rs2!GrossPay), 0, rs2!GrossPay)
                sql = "SET dateformat DMY UPDATE d_Payroll SET GPay=" & GPay & ", KgsSupplied = " & IIf(IsNull(rs2!qnty), 0, rs2!qnty) & " WHERE YYEAR = '" & year(Enddate) & "' AND MMONTH='" & month(Enddate) & "' AND SNo= '" & sno & "'"
                oSaccoMaster.GetRecordset (sql)
            Else
                sql = "SET dateformat DMY UPDATE d_Payroll SET GPay=0, KgsSupplied =0 WHERE YYEAR = '" & year(Enddate) & "' AND MMONTH='" & month(Enddate) & "' AND SNo= '" & !sno & "'"
                oSaccoMaster.GetRecordset (sql)
            End If

        Dim agrovet As Double
        Dim FSA As Double
        Dim AI As Double
        Dim Others As Double
        Dim TMShares As Double
        Dim HShares As Double
        Dim Advance As Double
        Dim Transport As Double
        Dim TCHP As Double
        Dim MidMonth As Double
        Dim KDBC As Double
            agrovet = 0
            FSA = 0
            AI = 0
            Others = 0
            TMShares = 0
            HShares = 0
            Advance = 0
            Transport = 0
            TCHP = 0
            KDBC = 0
            MidMonth = 0
           
          sql = " set dateformat dmy SELECT  [Description], SUM(Amount) AS Amount " _
          & "From d_Supplier_deduc WHERE  (Date_Deduc >= '" & Startdate & "' AND Date_Deduc <= '" & midMonthDate & "') AND SNo='" & sno & "'" _
                           & "GROUP BY [Description]"
         Set rsdeduction = oSaccoMaster.GetRecordset(sql)
         With rsdeduction
            If Not .EOF Then
                While Not .EOF
                DoEvents
                Dim description  As String
                Dim deduction, TotalDed As Double
                        description = ""
                        deduction = 0
                        TotalDed = 0
                    description = IIf(IsNull(rsdeduction!description), "others", rsdeduction!description)
                    deduction = IIf(IsNull(rsdeduction!Amount), 0, rsdeduction!Amount)
                    If UCase(description) = "AGROVET" Then
                     agrovet = agrovet + deduction
                    End If
                    If UCase(description) = "FSA" Then
                     FSA = FSA + deduction
                     End If
                    If UCase(description) = "AI" Then
                     AI = AI + deduction
                     End If
                    If UCase(description) = "REGISTRATION" Then
                     TMShares = TMShares + deduction
                     End If
                    If UCase(description) = "OTHERS" Then
                     Others = Others + deduction
                     End If
                    If UCase(description) = "SHARES" Then
                     HShares = HShares + deduction
                     End If
                    If UCase(description) = "ADVANCE" Then
                     Advance = Advance + deduction
                     End If
                    If UCase(description) = "TRANSPORT" Then
                     Transport = Transport + deduction
                    End If
                    If UCase(description) = "TCHP" Then
                     TCHP = TCHP + deduction
                    End If
                    If UCase(description) = "KDBC" Then
                     KDBC = KDBC + deduction
                    End If
                    If UCase(description) = "WEEKLY" Then
                     MidMonth = MidMonth + deduction
                    End If
                .MoveNext
                 frmProcess.Caption = sno
                Wend
             End If
            TotalDed = agrovet + FSA + AI + Others + TMShares + HShares + Advance + Transport + TCHP + MidMonth + KDBC
                
                sql = "SET DATEFORMAT DMY UPDATE    d_Payroll " _
                 & " SET  Transport = " & Transport & ", Agrovet = " & agrovet & ", AI = " & AI & ", TMShares = " & TMShares & ", FSA = " & FSA & ", HShares =" & HShares & ", Advance = " & Advance & ", " _
                                      & " Others = " & Others & ",MidMonth=" & MidMonth & " ,TDeductions =" & TotalDed & ", NPay = " & GPay - TotalDed & " ,TCHP=" & TCHP & " " _
                & "Where sno = '" & sno & "' AND  YYEAR = '" & year(Enddate) & "' AND MMONTH='" & month(Enddate) & "' AND SNO IN(SELECT SNO FROM D_SUPPLIERS WHERE TRADER=1)"
                oSaccoMaster.GetRecordset (sql)
        End With
    .MoveNext
    frmProcess.Caption = sno
    Wend
    End If
End With

oSaccoMaster.ExecuteThis ("d_sp_RegSharesweekly '" & EndMonth & "' ,'" & User & "'")


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
ProgressBar1.Value = 50
''oSaccoMaster.ExecuteThis ("set dateformat dmy UPDATE d_Payroll SET Transport=0,TDeductions=(TDeductions -Transport),NPay=(NPay + Transport) WHERE YYEAR = '" & year(Enddate) & "' AND MMONTH='" & month(Enddate) & "' ")
''oSaccoMaster.ExecuteThis ("set dateformat dmy     UPDATE    d_supplier_deduc   SET  amount=0 where [Description] = 'Transport' and YYEAR = '" & year(Enddate) & "' AND MMONTH='" & month(Enddate) & "'")

'****************** COMMENTED COZ TRANSPORTERS ARE PAID BY LELCHEGO COMPANY LTD
''''*********** Refresh Transporters Payroll First ****************
'''Refresh_TranportersPayroll
'''
'''Set Rst = oSaccoMaster.GetRecordset("select transcode from d_transporters  order by transcode asc")
'''While Not Rst.EOF
'''DoEvents
'''frmProcess.Caption = Rst.Fields(0)
'''oSaccoMaster.ExecuteThis ("d_sp_TransUpdate '" & Format(Startdate, "dd/mm/yyyy") & "','" & Format(midMonthDate, "dd/mm/yyyy") & "','" & User & "','" & Trim(Rst.Fields(0)) & "'")
'''Rst.MoveNext
'''Wend
'''Set Rst = Nothing
'''ProgressBar1.value = 70
'''
'''    updatepayroll_deduc
'''    transproll
'''    TRIPtransproll

    
    '===================== MID MONTH ADVANCE========================================
     
         'midMonthDate = DateSerial(year(Enddate), month(Enddate), 15)
        sql = "SET DATEFORMAT DMY INSERT d_supplier_deduc   (SNo, Date_Deduc, Description, Amount, Period, StartDate, EndDate, auditid, auditdatetime, yyear, Remarks," _
             & " status, status2, status3, status4,status5, status6, BranchCode) " _
         & "  SELECT   SNo,'" & midMonthDate & "','Weekly',NPay,'" & Format(Enddate, "mm/yyyy") & "','" & Startdate & "','" & Enddate & "' , '" & User & "',GETDATE(),'" & year(Enddate) & "','Payment',0,0,0,0,0,0,0 " _
           & "  FROM d_Payroll  where YYear='" & year(Enddate) & "' and MMonth='" & month(Enddate) & "'  and npay>0 AND SNO IN(SELECT SNO FROM D_SUPPLIERS WHERE TRADER=1)"
       oSaccoMaster.ExecuteThis (sql)
  
    
Dim Va As Integer


ProgressBar1.Value = 100

MsgBox "Completed Payroll"

End Sub
Private Sub cmdprocess_Click()
'addomittedentried

Dim Yr As Integer, ym As Integer
Dim currdate As Date, Shares As Double, MaxShares As Double, SharesAcc As Double, AmountCur As Double
Dim desc As String, remark As String, rate As Double, rsdeduction As New Recordset
currdate = Format(Get_Server_Date, "dd/mm/yyyy")
Startdate = DateSerial(year(dtpProcess), month(dtpProcess), 1)
Enddate = DateSerial(year(dtpProcess), month(dtpProcess) + 1, 1 - 1)
ProgressBar1.Value = 0
Yr = year(dtpProcess)
ym = DateDiff("d", Startdate, currdate)

updatepayroll
GetBonus
'//update deduction before anything else before running the payroll start here
Dim rshast1 As New ADODB.Recordset, descrip As String, amt As Double, sno As Long

ProgressBar1.Value = 20

sql = ("d_sp_PresetDeductAssign '" & Startdate & "','" & Enddate & "'," & Yr & ",'" & User & "'")
oSaccoMaster.ExecuteThis (sql)

'presetdeductassign
sql = ("d_sp_PresetDeductFlatRate '" & Startdate & "','" & Enddate & "'," & Yr & ",'" & User & "'")
oSaccoMaster.ExecuteThis (sql)

ProgressBar1.Value = 40
'Payroll update
sql = " SET DATEFORMAT DMY SELECT  Sno From d_Payroll WHERE     EndofPeriod = '" & Enddate & "'  order by sno"
Set rs = oSaccoMaster.GetRecordset(sql)
With rs
    If Not .EOF Then
    While Not .EOF
    DoEvents
    sno = !sno
    frmProcess.Caption = sno
   
        Dim RsShares As New Recordset
        Dim rspu As New Recordset
        Dim GPay As Double

            
             sql = "SET dateformat dmy SELECT     SUM(QSupplied) AS QNTY, SUM(QSupplied * (PPU)) AS GrossPay From d_Milkintake " _
             & " WHERE     (TransDate BETWEEN '" & Startdate & "'  AND '" & Enddate & "' AND SNo ='" & sno & "')"
             
             Set rs2 = oSaccoMaster.GetRecordset(sql)
            
             GPay = 0
             
             If Not rs2.EOF Then
             GPay = IIf(IsNull(rs2!GrossPay), 0, rs2!GrossPay)
             sql = "SET dateformat DMY UPDATE d_Payroll SET GPay=" & GPay & ",NPay=" & GPay & ", KgsSupplied = " & IIf(IsNull(rs2!qnty), 0, rs2!qnty) & " WHERE EndofPeriod = '" & Enddate & "' AND SNo= '" & sno & "'"
             oSaccoMaster.GetRecordset (sql)
             Else
             sql = "SET dateformat DMY UPDATE d_Payroll SET GPay=0, KgsSupplied =0 WHERE EndofPeriod = '" & Enddate & "' AND SNo= '" & sno & "'"
             oSaccoMaster.GetRecordset (sql)
             End If
                        
          'Add Bonus
       
        sql = " set dateformat dmy SELECT  S.SNo,isnull(SUM(S.Amount),0) AS Amount From d_Supplier_deduc S WHERE  S.SNo='" & sno & "' and S.Amount<10 and S.[Description]='BONUS' and (S.Date_Deduc >= '" & Startdate & "' AND S.Date_Deduc <= '" & Enddate & "') GROUP BY S.SNo"
        Set rs2 = oSaccoMaster.GetRecordset(sql)
        If Not rs2.EOF Then
        GPay = GPay + (rs2!Amount * -1)
        sql = "SET dateformat DMY UPDATE d_Payroll SET GPay=" & GPay & ",NPay=" & GPay & "  WHERE EndofPeriod = '" & Enddate & "' AND SNo= '" & sno & "'"
        oSaccoMaster.GetRecordset (sql)
        End If
        
        Dim agrovet As Double, NHIF As Double, Water As Double, Silage As Double, Prepay As Double
        Dim FSA As Double, AI As Double, Others As Double, TMShares As Double, HShares As Double
        Dim Advance As Double, Transport As Double, TCHP As Double, KDBC As Double, MidMonth As Double
        Dim Sshares As Double, CashBf As Double, Ecf As Double, TotalDed As Double, LoanSav As Double
        Dim description  As String, deduction, Insurance As Double, Npay As Double
            agrovet = 0
            FSA = 0
            AI = 0
            Others = 0
            TMShares = 0
            HShares = 0
            Advance = 0
            Transport = 0
            TCHP = 0
            KDBC = 0
            MidMonth = 0
            NHIF = 0
            Silage = 0
            Water = 0
            Prepay = 0
            loan = 0
            CashBf = 0
            Ecf = 0
            Sshares = 0
            LoanSav = 0
            Insurance = 0
            Npay = 0
            
          sql = " set dateformat dmy SELECT  S.[Description], SUM(S.Amount) AS Amount " _
          & "From d_Supplier_deduc S WHERE  (S.Date_Deduc >= '" & Startdate & "' AND S.Date_Deduc <= '" & Enddate & "') AND S.SNo='" & sno & "'" _
            & "GROUP BY S.[Description]"
         Set rsdeduction = oSaccoMaster.GetRecordset(sql)
         With rsdeduction
            If Not .EOF Then
                While Not .EOF
                DoEvents
                        description = ""
                        deduction = 0
                        TotalDed = 0
                        
                    description = IIf(IsNull(rsdeduction!description), "others", rsdeduction!description)
                    deduction = IIf(IsNull(rsdeduction!Amount), 0, rsdeduction!Amount)
                    If UCase(Trim(description)) = "AGROVET" Then
                     agrovet = agrovet + deduction
                    End If
                    If UCase(Trim(description)) = "LEPESA LOAN" Then
                     loan = loan + deduction
                     End If
                    If UCase(Trim(description)) = "AI" Then
                     AI = AI + deduction
                     End If
                    If UCase(Trim(description)) = "CASH BF" Then
                     CashBf = CashBf + deduction
                     End If
                    If UCase(Trim(description)) = "SHARES" Then
                     HShares = HShares + deduction
                     End If
                    If UCase(Trim(description)) = "ADVANCE" Then
                     Advance = Advance + deduction
                     End If
                    If UCase(Trim(description)) = "TRANSPORT" Then
                     Transport = Transport + deduction
                    End If
                    If UCase(Trim(description)) = "NHIF" Then
                     NHIF = NHIF + deduction
                    End If
                    If UCase(Trim(description)) = "ECF" Then
                     Ecf = Ecf + deduction
                    End If
                    If UCase(Trim(description)) = "LOAN SAVINGS" Then
                     LoanSav = LoanSav + deduction
                     End If
                    If UCase(Trim(description)) = "LEPESA SHARES" Then
                     Sshares = Sshares + deduction
                    End If
                    If UCase(Trim(description)) = "PREPAYMENTS" Then
                     Prepay = Prepay + deduction
                    End If
                    If UCase(Trim(description)) = "WATER BILL" Then
                     Water = Water + deduction
                    End If
                    If UCase(Trim(description)) = "SILAGE" Then
                     Silage = Silage + deduction
                    End If
                    If UCase(Trim(description)) = "INSURANCE" Then
                     Insurance = Insurance + deduction
                    End If
                    If UCase(Trim(description)) = "OTHERS" Then
                    Others = Others + deduction
                    End If
                .MoveNext
                 frmProcess.Caption = sno
                Wend
             End If

            ' Transport = 0
             TotalDed = 0
            TotalDed = agrovet + loan + AI + Others + Sshares + HShares + Advance + Transport + NHIF + CashBf + Ecf + Silage + Water + Prepay + LoanSav + Insurance
            Npay = GPay - TotalDed
            sql = "SET DATEFORMAT DMY UPDATE    d_Payroll " _
             & " SET  Transport = " & Transport & ", Agrovet = " & agrovet & ", AI = " & AI & ",HShares =" & HShares & ", Advance = " & Advance & ",Water=" & Water & " ," _
             & " FSA = " & loan & ",LSHARES=" & Sshares & ",LSav=" & LoanSav & " ,NHIF =" & NHIF & ",ECF =" & Ecf & ",CASHBF =" & CashBf & ",PREPAY =" & Prepay & ",Silage =" & Silage & "," _
            & " Insurance = " & Insurance & ",Others = " & Others & ",TDeductions =" & TotalDed & ",GPay = " & GPay & " ,NPay = " & Npay & " " _
            & "Where sno = '" & sno & "' And EndofPeriod =  '" & Enddate & "'"
            oSaccoMaster.GetRecordset (sql)
                
        End With
    .MoveNext
    frmProcess.Caption = sno
    Wend
    End If
End With

oSaccoMaster.ExecuteThis ("SET DATEFORMAT DMY UPDATE P SET P.AccountNumber=S.AccNo,P.Bank=S.Bcode,P.BBranch=S.BBranch FROM d_Suppliers S INNER JOIN d_Payroll P ON P.SNo=S.SNo WHERE P.EndofPeriod='" & Enddate & "'")

'**** update Deducted shares*************
oSaccoMaster.ExecuteThis ("d_sp_RegShares '" & Enddate & "','Payroll'")

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''ProgressBar1.Value = 50
''oSaccoMaster.ExecuteThis ("set dateformat dmy UPDATE d_Payroll SET Transport=0,TDeductions=(TDeductions -Transport),NPay=(NPay + Transport) WHERE Endofperiod='" & Enddate & "'")
''oSaccoMaster.ExecuteThis ("set dateformat dmy  UPDATE d_supplier_deduc   SET  amount=0 where [Description] = 'Transport'and  EndDate ='" & Enddate & "'")
''
''Refresh_TranportersPayroll
''
''Startdate = DateSerial(year(dtpProcess), month(dtpProcess), 1)
''Enddate = DateSerial(year(dtpProcess), month(dtpProcess) + 1, 1 - 1)
''
''Set Rst = oSaccoMaster.GetRecordset("select transcode from d_transporters  order by transcode asc")
''While Not Rst.EOF
''DoEvents
''frmProcess.Caption = Rst.Fields(0)
''oSaccoMaster.ExecuteThis ("d_sp_TransUpdate '" & Format(Startdate, "dd/mm/yyyy") & "','" & Format(Enddate, "dd/mm/yyyy") & "','" & User & "','" & Trim(Rst.Fields(0)) & "'")
''Rst.MoveNext
''Wend
Set Rst = Nothing

ProgressBar1.Value = 70

Startdate = DateSerial(year(dtpProcess), month(dtpProcess), 1)
Enddate = DateSerial(year(dtpProcess), month(dtpProcess) + 1, 1 - 1)

''    updatepayroll_deduc
    
Startdate = DateSerial(year(dtpProcess), month(dtpProcess), 1)
Enddate = DateSerial(year(dtpProcess), month(dtpProcess) + 1, 1 - 1)

    transproll
    TRIPtransproll
    Dim Va As Integer

oSaccoMaster.ExecuteThis ("d_sp_Periods '" & Enddate & "'," & Va & ",'" & User & "'")
ProgressBar1.Value = 100

MsgBox "Completed Payroll"
'vbDefault

End Sub
Sub updatepayroll()
Startdate = DateSerial(year(dtpProcess), month(dtpProcess), 1)
Enddate = DateSerial(year(dtpProcess), month(dtpProcess) + 1, 1 - 1)
Dim sno As String, speriod As Date, eperiod As Date
sql = ""
sql = "set dateformat dmy select distinct sno from d_milkintake where transdate<='" & Enddate & "' and transdate>='" & Startdate & "'  order by sno"
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
sno = rs.Fields(0)

sql = "select sno from d_payroll where mmonth=" & month(Enddate) & " and yyear=" & year(Enddate) & " and sno='" & sno & "' order by sno"
Set Rst = oSaccoMaster.GetRecordset(sql)
    If Rst.EOF Then
        sql = ""
        sql = "insert into d_Payroll (SNo,EndofPeriod,auditid ) "
        sql = sql & " values ('" & sno & "','" & Enddate & "','" & User & "' )"
        oSaccoMaster.ExecuteThis (sql)
        
    End If
    frmProcess.Caption = sno
    rs.MoveNext
Wend
oSaccoMaster.GetRecordset ("SET DATEFORMAT DMY UPDATE  d_payroll SET EndofPeriod= '" & Enddate & "' where mmonth=" & month(Enddate) & " and yyear=" & year(Enddate) & "")

Exit Sub
End Sub
Sub GetBonus()
Dim rsdeduction As New ADODB.Recordset, sno As Double
Startdate = DateSerial(year(dtpProcess), month(dtpProcess), 1)
Enddate = DateSerial(year(dtpProcess), month(dtpProcess) + 1, 1 - 1)

sql = " set dateformat dmy SELECT  S.SNo,SUM(S.Amount) AS Amount From d_Supplier_deduc S WHERE S.Amount<10 and S.[Description]='BONUS' and (S.Date_Deduc >= '" & Startdate & "' AND S.Date_Deduc <= '" & Enddate & "') GROUP BY S.SNo"
Set rsdeduction = oSaccoMaster.GetRecordset(sql)
With rsdeduction
   If Not .EOF Then
       While Not .EOF
         DoEvents
          sno = rsdeduction!sno
           sql = ""
           sql = "select sno from d_payroll where mmonth=" & month(Enddate) & " and yyear=" & year(Enddate) & " and sno='" & sno & "' order by sno"
           Set Rst = oSaccoMaster.GetRecordset(sql)
           If Rst.EOF Then
               sql = ""
               sql = "insert into d_Payroll (SNo,EndofPeriod,auditid ) "
               sql = sql & " values ('" & sno & "','" & Enddate & "','" & User & "' )"
               oSaccoMaster.ExecuteThis (sql)
           End If
           frmProcess.Caption = sno
         rsdeduction.MoveNext
       Wend
   End If
End With
End Sub
Public Sub TRIPtransproll()
'On Error Resume Next
On Error GoTo milgo

Dim GPay As Double
Dim qnty As Double
Dim tcode As String
Dim Amnt As Double
Dim subsidy As Double

Set Rst = oSaccoMaster.GetRecordset(" SELECT SUM(dbo.d_TripTransDetailed.qnty) AS QNTY, dbo.d_TripTransDetailed.Trans_Code AS Code, SUM(dbo.d_TripTransDetailed.Amount) AS Amount,SUM(d_TripTransDetailed.Subsidy) As Subsidy From d_TripTransDetailed WHERE     Transdate between  '" & Startdate & "' and '" & Enddate & "'   GROUP BY d_TripTransDetailed.Trans_Code")
While Not Rst.EOF
'TCode = "T308"
    DoEvents
    tcode = Rst.Fields("code")
    frmProcess.Caption = tcode
    subsidy = IIf(IsNull(Rst.Fields("subsidy")), 0, Rst.Fields("subsidy"))
    Amnt = IIf(IsNull(Rst.Fields("amount")), 0, Rst.Fields("amount"))
    GPay = IIf(IsNull(Amnt + subsidy), 0, (Amnt + subsidy))
    qnty = IIf(IsNull(Rst.Fields("qnty")), 0, Rst.Fields("qnty"))
    tcode = IIf(IsNull(Rst.Fields("Code")), 0, Rst.Fields("Code"))
    
    oSaccoMaster.ExecuteThis ("exec d_sp_UpdateTransPay '" & tcode & "', '" & qnty & "'," & Amnt & "," & subsidy & "," & GPay & ", '" & Enddate & "','" & User & "'")
    
Dim agrovet As Double
Dim FSA  As Double
Dim AI As Double
Dim Others As Double
Dim TMShares As Double
Dim HShares As Double
Dim Advance As Double
Dim TotalDed As Double

agrovet = 0
FSA = 0
AI = 0
Others = 0
TMShares = 0
HShares = 0
Advance = 0

Dim desc As String
Dim deduction As Double

    Set rst2 = oSaccoMaster.GetRecordset(" set dateformat dmy SELECT  [Description], SUM(Amount) AS Amount From d_Transport_Deduc WHERE  (startdate>='" & Startdate & "'and enddate<='" & Enddate & "') and TransCode='" & tcode & "'  GROUP BY [Description]")

 
    While Not rst2.EOF
    DoEvents
    'frmProcess.Caption = rst2.Fields("TransCode")
        desc = IIf(IsNull(rst2.Fields("description")), "others", rst2.Fields("description"))
        deduction = IIf(IsNull(rst2.Fields("amount")), 0, rst2.Fields("amount"))
        
        If UCase(desc) = "AGROVET" Then
        agrovet = deduction
        
        ElseIf UCase(desc) = "FSA" Then
        FSA = deduction
        
        ElseIf UCase(desc) = "AI" Then
        AI = deduction
        
        ElseIf UCase(desc) = "TMSHARES" Then
        TMShares = deduction
        
        ElseIf UCase(desc) = "OTHERS" Then
        Others = deduction
        
        ElseIf UCase(desc) = "HSHARES" Then
        HShares = deduction
        
        ElseIf UCase(desc) = "ADVANCE" Then
        Advance = deduction
        
        End If
        
        TotalDed = agrovet + FSA + AI + TMShares + Others + HShares + Advance
        
        oSaccoMaster.ExecuteThis ("exec d_sp_UpdateTransDed  '" & tcode & "',' " & Enddate & "'," & TotalDed & "," & agrovet & "," & AI & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & "")
 
    
    rst2.MoveNext
    Wend
    

    Rst.MoveNext
    Wend


Exit Sub

milgo:
    MsgBox err.description, vbInformation, tcode
    Exit Sub
    
End Sub




Private Sub Cmds1_Click()
frmSearchGLAccounts.Show vbModal
If SearchValue <> "" Then
Txtcreditedac = SearchValue
sql = ""
sql = "select * from glsetup where accno='" & SearchValue & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
lblcreditedac = rs!GlAccName
End If
End Sub

Private Sub Cmds2_Click()
frmSearchGLAccounts.Show vbModal
If SearchValue <> "" Then
Txtdebitedac = SearchValue
sql = ""
sql = "select * from glsetup where accno='" & SearchValue & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
'lblcreditedac = rs!GlAccName
End If
End Sub

Private Sub cmdtotalmonthlyq_Click()

sql = ""
sql = "SET              dateformat dmy SELECT     SUM(qsupplied) From d_Milkintake WHERE     transdate BETWEEN '" & dtpFrom & "' AND '" & dtpTo & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
txttotal = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
End If
End Sub

Private Sub cmdupdatebr_Click()
'Startdate = DateSerial(year(dtpProcess), month(dtpProcess), 1)
'Enddate = DateSerial(year(dtpProcess), month(dtpProcess) + 1, 1 - 1)
'Dim sno As String, speriod As Date, eperiod As Date
'sql = ""
'sql = "set dateformat dmy select distinct sno from d_milkintake where transdate<='" & Enddate & "' and transdate>='" & Startdate & "'  order by sno"
'Set rs = oSaccoMaster.GetRecordset(sql)
'While Not rs.EOF
'DoEvents
'sno = rs.Fields(0)
'
'sql = "select sno from d_payroll where mmonth=" & month(dtpProcess) & " and yyear=" & year(dtpProcess) & " and sno=" & sno & " order by sno"
'Set Rst = oSaccoMaster.GetRecordset(sql)
'    If Rst.EOF Then
'        sql = ""
'        sql = "insert into d_Payroll (SNo,EndofPeriod,auditid ) "
'        sql = sql & " values (" & sno & ",'" & dtpProcess & "','" & User & "' )"
'        oSaccoMaster.ExecuteThis (sql)
'
'    End If
'    frmProcess.Caption = sno
'    rs.MoveNext
'Wend
  
End Sub

Private Sub Command1_Click()
On Error GoTo ErrorHandler
Dim sno As String
sno = 1
'process_payroll (sno)
'UPDATE THE ACCOUNTS FOR LELCHEGO FSA
sql = "SELECT     *  FROM         FSA_ACCS  WHERE     (Payrollno <> 'NA')  ORDER BY 4"
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
DoEvents
sql = "UPDATE    d_Suppliers   SET              accno='" & rs.Fields(0) & "' where sno='" & Trim(rs.Fields(3)) & "'"
oSaccoMaster.ExecuteThis (sql)
frmProcess.Caption = rs.Fields(0)
rs.MoveNext
Wend
 MsgBox "Done"
 Exit Sub
ErrorHandler:
 MsgBox err.description
End Sub

Private Sub Command2_Click()
Startdate = DateSerial(year(dtpProcess), month(dtpProcess), 1)
Enddate = DateSerial(year(dtpProcess), month(dtpProcess) + 1, 1 - 1)
Dim sno As String, speriod As Date, eperiod As Date
Dim Others As String, REMARKS As String, rate As Double, Rated As Integer
Dim deduction As String
Dim Stopped As Integer
sql = ""
sql = "set dateformat dmy select distinct sno from d_milkintake where transdate<='" & Enddate & "' and transdate>='" & Startdate & "'  order by sno"
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
DoEvents
sno = rs.Fields(0)

sql = "select * from d_PreSets where  sno=" & sno & " order by Deduction"
Set Rst = oSaccoMaster.GetRecordset(sql)
    If Rst.EOF Then
    'SELECT     SNo, Deduction, Remark, StartDate, Rate, Stopped, Auditdatetime, AuditId, Rated
'From d_PreSets
'//OPERATIONS ON OTHERS
        Stopped = 0
        'sno = Rst.Fields("sno")
        deduction = "Others"
        REMARKS = "OPERATIONS"
        Rated = 1
        rate = 1.7
        Dim Rst1 As New ADODB.Recordset
        sql = "select sno from d_PreSets where  sno=" & sno & " and deduction='Others' and remark='OPERATIONS' order by Deduction"
        Set Rst1 = oSaccoMaster.GetRecordset(sql)
        If Rst1.EOF Then
            sql = ""
            sql = "set dateformat dmy insert into d_PreSets (SNo, Deduction, Remark, StartDate, Rate, Stopped, Auditdatetime, AuditId, Rated) "
            sql = sql & " values (" & sno & ", '" & deduction & "','" & REMARKS & "', '" & Startdate & "', " & rate & ", " & Stopped & ", '" & Get_Server_Date & "', '" & User & "', " & Rated & ")"
            oSaccoMaster.ExecuteThis (sql)
        End If
        
        'HSARES
        Stopped = 0
        deduction = "HShares"
        REMARKS = ""
        Rated = 1
        rate = 0.3
        
        sql = "select sno from d_PreSets where sno=" & sno & " and deduction='HShares' order by Deduction"
        Set Rst1 = oSaccoMaster.GetRecordset(sql)
        If Rst1.EOF Then
            sql = ""
            sql = "set dateformat dmy insert into d_PreSets (SNo, Deduction, Remark, StartDate, Rate, Stopped, Auditdatetime, AuditId, Rated) "
            sql = sql & " values (" & sno & ", '" & deduction & "','" & REMARKS & "', '" & Startdate & "', " & rate & ", " & Stopped & ", '" & Get_Server_Date & "', '" & User & "', " & Rated & ")"
            oSaccoMaster.ExecuteThis (sql)
        End If
        
        
    End If
    '//cbo fees
    
            sql = "select description from d_supplier_standingorder where sno='" & sno & "' and description ='CBO'"
            Set Rst = oSaccoMaster.GetRecordset(sql)
            If Rst.EOF Then
            '//Update deductions
                Set cn = New ADODB.Connection
                sql = "set dateformat dmy INSERT INTO d_supplier_standingorder"
                sql = sql & "           (SNo, Date_Deduc, Description, Amount, MaxAmount, Period, StartDate, EndDate, auditid,  yyear, Remarks)"
                sql = sql & "  VALUES     (" & sno & ",'" & Startdate & "','CBO',50,0,'" & Format(Startdate, "mmm-YYYY") & "','" & Startdate & "','31/05/2015','" & User & "'," & year(Enddate) & ",'" & REMARKS & "')"
                oSaccoMaster.ExecuteThis (sql)
                
                
                
            End If
    
    
    Stopped = 0
        deduction = ""
        REMARKS = ""
        Rated = 1
        rate = 0
    
    frmProcess.Caption = sno
    rs.MoveNext
Wend

End Sub

Private Sub dtpCarry_Validate(Cancel As Boolean)
dtpCarry = DateSerial(year(dtpCarry), month(dtpCarry) + 1, 1 - 1)

End Sub

Private Sub Form_Load()
dtpProcess = DateSerial(year(Get_Server_Date), month(Get_Server_Date) + 1, 1 - 1)
dtpCarry = DateSerial(year(Get_Server_Date), month(Get_Server_Date) + 1, 1 - 1)
DTPEOD = Format(Get_Server_Date, "dd/mm/yyyy")
dtpFrom = DTPEOD
dtpTo = DTPEOD
previousp = dtpFrom

Txtcreditedac = "L041"
Txtdebitedac = "E039"

End Sub

Public Sub transproll()
'On Error Resume Next
On Error GoTo milgo
Dim GPay As Double, qnty As Double, tcode As String, Amnt As Double, subsidy As Double
Dim desc As String, deduction As Double, description As String

Set Rst = oSaccoMaster.GetRecordset(" SELECT SUM(dbo.d_TransDetailed.qnty) AS QNTY, dbo.d_TransDetailed.Trans_Code AS Code, SUM(dbo.d_TransDetailed.Amount) AS Amount,SUM(d_TransDetailed.Subsidy) As Subsidy From d_TransDetailed WHERE     EndPeriod = '" & Enddate & "'  GROUP BY d_TransDetailed.Trans_Code")
While Not Rst.EOF
    DoEvents
    tcode = Rst.Fields("code")
    frmProcess.Caption = tcode
    subsidy = IIf(IsNull(Rst.Fields("subsidy")), 0, Rst.Fields("subsidy"))
    Amnt = IIf(IsNull(Rst.Fields("amount")), 0, Rst.Fields("amount"))
    GPay = IIf(IsNull(Amnt + subsidy), 0, (Amnt + subsidy))
    qnty = IIf(IsNull(Rst.Fields("qnty")), 0, Rst.Fields("qnty"))
    tcode = IIf(IsNull(Rst.Fields("Code")), 0, Rst.Fields("Code"))
    
    oSaccoMaster.ExecuteThis ("exec d_sp_UpdateTransPay '" & tcode & "', '" & qnty & "'," & Amnt & "," & subsidy & "," & GPay & ", '" & Enddate & "','" & User & "'")
    
        Dim agrovet As Double, NHIF As Double, Water As Double, Silage As Double, Prepay As Double, Welfare As Double
        Dim FSA As Double, AI As Double, Others As Double, TMShares As Double, HShares As Double, Fuel As Double, Insurance As Double
        Dim Advance As Double, Transport As Double, TCHP As Double, KDBC As Double, MidMonth As Double
        Dim Sshares As Double, CashBf As Double, Ecf As Double, TotalDed As Double, MilkR As Double, MilkV As Double
            agrovet = 0
            FSA = 0
            AI = 0
            Others = 0
            TMShares = 0
            HShares = 0
            Advance = 0
            Transport = 0
            TCHP = 0
            KDBC = 0
            MidMonth = 0
            NHIF = 0
            Silage = 0
            Water = 0
            Prepay = 0
            loan = 0
            CashBf = 0
            Ecf = 0
            Sshares = 0
            MilkR = 0
            MilkV = 0
            Fuel = 0
            Welfare = 0
            Insurance = 0
            
    Set rst2 = oSaccoMaster.GetRecordset("set dateformat dmy SELECT  T.[Description], SUM(T.Amount) AS Amount From d_Transport_Deduc T  WHERE  (T.startdate>='" & Startdate & "'and T.enddate<='" & Enddate & "') and T.TransCode='" & tcode & "' GROUP BY T.[Description]")
    While Not rst2.EOF
    DoEvents
    frmProcess.Caption = tcode
        desc = IIf(IsNull(rst2.Fields("description")), "others", rst2.Fields("description"))
        description = desc
        deduction = IIf(IsNull(rst2.Fields("amount")), 0, rst2.Fields("amount"))
        
                    If UCase(description) = "AGROVET" Then
                     agrovet = agrovet + deduction
                    End If
                    If UCase(description) = "LEPESA LOAN" Then
                     loan = loan + deduction
                     End If
                    If UCase(description) = "AI" Then
                     AI = AI + deduction
                     End If
                    If UCase(description) = "CASH BF" Then
                     CashBf = CashBf + deduction
                     End If
                    If UCase(description) = "SHARES" Then
                     HShares = HShares + deduction
                     End If
                    If UCase(description) = "ADVANCE" Then
                     Advance = Advance + deduction
                     End If
                    If UCase(description) = "TRANSPORT" Then
                     Transport = Transport + deduction
                    End If
                    If UCase(description) = "NHIF" Then
                     NHIF = NHIF + deduction
                    End If
                    If UCase(description) = "ECF" Then
                     Ecf = Ecf + deduction
                    End If
                    If UCase(description) = "LEPESA SHARES" Then
                     Sshares = Sshares + deduction
                    End If
                    If UCase(description) = "PREPAYMENTS" Then
                     Prepay = Prepay + deduction
                    End If
                    If UCase(description) = "WATER BILL" Then
                     Water = Water + deduction
                    End If
                    If UCase(description) = "SILAGE" Then
                     Silage = Silage + deduction
                    End If
                    If UCase(description) = "FUEL" Then
                     Fuel = Fuel + deduction
                    End If
                    If UCase(description) = "MILK REJETCS" Then
                     MilkR = MilkR + deduction
                    End If
                    If UCase(description) = "MILK VARIANCE" Then
                     MilkV = MilkV + deduction
                    End If
                    If UCase(description) = "WELFARE" Then
                     Welfare = Welfare + deduction
                    End If
                    If UCase(description) = "INSURANCE" Then
                     Insurance = Insurance + deduction
                    End If
                    If UCase(description) = "OTHERS" Then
                    Others = Others + deduction
                    End If
        
                    TotalDed = agrovet + loan + AI + Others + Sshares + HShares + Advance + Transport + NHIF + CashBf + Ecf + Silage + Water + Prepay + MilkR + MilkV + Fuel + Welfare + Insurance
                
            sql = "SET DATEFORMAT DMY UPDATE    d_TransportersPayRoll " _
             & " SET  Agrovet = " & agrovet & ", AI = " & AI & ",HShares =" & HShares & ", Advance = " & Advance & ",Water=" & Water & " ,Welfare=" & Welfare & " ," _
             & " FSA = " & loan & ",LSHARES=" & Sshares & " ,NHIF =" & NHIF & ",ECF =" & Ecf & ",CASHBF =" & CashBf & ",PREPAY =" & Prepay & ",Silage =" & Silage & "," _
            & " MilkR = " & MilkR & ",MilkV=" & MilkV & " ,fuel =" & Fuel & ",Insurance =" & Insurance & "," _
            & " Others = " & Others & ",Totaldeductions =" & TotalDed & ", NetPay = " & GPay - TotalDed & "" _
            & "Where Code = '" & tcode & "' And EndPeriod =  '" & Enddate & "'"
            oSaccoMaster.GetRecordset (sql)
        'oSaccoMaster.ExecuteThis ("exec d_sp_UpdateTransDed  '" & tcode & "',' " & Enddate & "'," & TotalDed & "," & agrovet & "," & AI & "," & TMShares & "," & FSA & "," & HShares & "," & Advance & "," & Others & "")
     
    rst2.MoveNext
    Wend

    Rst.MoveNext
    Wend
Exit Sub

milgo:
    MsgBox err.description, vbInformation, tcode
    Exit Sub
    
    
End Sub

Public Sub Transportertransporter()
Dim ptrans As String
Dim tt As String
Dim rate As Integer

Set Rst = oSaccoMaster.GetRecordset("select transcode,ptransporter,tt from d_transporters where tt=1")
While Not Rst.EOF
DoEvents
Rst.MoveNext
Wend

2
End Sub



Public Sub updatepayroll_deduc()

Set Rst = oSaccoMaster.GetRecordset("select sno,sum(amount)as amount from d_transdetailed where endperiod='" & dtpProcess & "' and amount>0 group by sno order by sno asc")
While Not Rst.EOF
DoEvents
frmProcess.Caption = Rst.Fields("sno")
    oSaccoMaster.ExecuteThis ("set dateformat dmy UPDATE d_Payroll SET Transport=" & Rst.Fields("amount") & " WHERE SNo='" & Rst.Fields("sno") & "' AND Endofperiod='" & dtpProcess & "' ")
    oSaccoMaster.ExecuteThis ("set dateformat dmy update d_payroll set TDeductions=TDeductions +" & Rst.Fields("amount") & ",NPay=gpay - tdeductions  WHERE SNo='" & Rst.Fields("sno") & "' AND Endofperiod='" & dtpProcess & "'  ")
    oSaccoMaster.ExecuteThis ("set dateformat dmy update d_payroll set NPay=gpay - tdeductions  WHERE SNo='" & Rst.Fields("sno") & "' AND Endofperiod='" & dtpProcess & "'  ")
    sql = ""
   sql = " set dateformat dmy select sno from d_supplier_deduc where sno='" & Rst.Fields("sno") & "' and [Description]='Transport' and Date_Deduc between '" & Startdate & "'  and '" & Enddate & "' "
   Set rst6 = oSaccoMaster.GetRecordset(sql)
   If rst6.EOF Then
    oSaccoMaster.ExecuteThis ("set dateformat dmy   insert into  d_supplier_deduc  (sno,Date_Deduc,[Description],Amount,StartDate,EndDate,auditid) values('" & Rst.Fields("sno") & "','" & Enddate & "','Transport'," & Rst.Fields("amount") & ",'" & Startdate & "','" & Enddate & "' ,'" & User & "')  ")
    Else
    sql = ""
    sql = "set dateformat dmy update d_supplier_deduc set amount=" & Rst.Fields("amount") & " where sno='" & Rst.Fields("sno") & "' and [Description]='Transport' and Date_Deduc between '" & Startdate & "'  and '" & Enddate & "' "
     oSaccoMaster.ExecuteThis (sql)
    End If
    Rst.MoveNext
Wend
End Sub

Public Sub Refresh_TranportersPayroll()
Dim year As Integer
Dim month As Integer


oSaccoMaster.ExecuteThis (" set dateformat dmy    delete from d_TransportersPayRoll where  yyear=year('" & dtpProcess & "') and mmonth =month('" & dtpProcess & "')")

Set Rst = oSaccoMaster.GetRecordset("set dateformat dmy select distinct Trans_Code from d_Transport where sno in (select sno from d_milkintake where (month(transdate)=month('" & dtpProcess & "')) and (year(transdate)=year('" & dtpProcess & "')))" _
  & "  and active=1 and Trans_Code not in(select code from d_TransportersPayRoll where yyear=year('" & dtpProcess & "') and mmonth =month('" & dtpProcess & "')) order by Trans_Code asc")
While Not Rst.EOF
DoEvents
frmProcess.Caption = Rst.Fields("Trans_Code")
    oSaccoMaster.ExecuteThis ("insert into d_TransportersPayRoll (code,EndPeriod) values('" & Trim$(Rst.Fields("Trans_Code")) & "','" & dtpProcess & "')")
Rst.MoveNext
Wend
End Sub
Public Sub addomittedentried()
Dim year As Integer
Dim month As Integer

Set Rst = oSaccoMaster.GetRecordset("set dateformat dmy select distinct sno from d_milkintake where (month(transdate)=month('" & dtpProcess & "')) and (year(transdate)=year('" & dtpProcess & "')) and (sno not in(select sno from d_payroll where yyear=year('" & dtpProcess & "') and mmonth =month('" & dtpProcess & "'))) order by sno asc")
While Not Rst.EOF
DoEvents
frmProcess.Caption = Rst.Fields("sno")

    oSaccoMaster.ExecuteThis ("insert into d_payroll (sno,endofperiod) values('" & Rst.Fields("sno") & "','" & dtpProcess & "')")
Rst.MoveNext
Wend

'Suppliers who didnt take last month cash and didnt supply this month milk.
Set Rst = oSaccoMaster.GetRecordset("set dateformat dmy select distinct sno from d_supplier_deduc where (month(Date_Deduc)=month('" & dtpProcess & "')) and (year(Date_Deduc)=year('" & dtpProcess & "')) and Description='CASH BF' and (sno not in(select sno from d_payroll where yyear=year('" & dtpProcess & "') and mmonth =month('" & dtpProcess & "'))) order by sno asc")
While Not Rst.EOF
DoEvents
frmProcess.Caption = Rst.Fields("sno")

    oSaccoMaster.ExecuteThis ("insert into d_payroll (sno,endofperiod) values('" & Rst.Fields("sno") & "','" & dtpProcess & "')")
Rst.MoveNext
Wend
End Sub

Private Sub Txtcreditedac_Change()
Set Rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & Txtcreditedac & "'")
If Not Rst.EOF Then
    lblcreditedac = Rst.Fields(0)
End If
End Sub

Private Sub Txtdebitedac_Change()
Set Rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & Txtdebitedac & "'")
If Not Rst.EOF Then
    lbldebitedac = Rst.Fields(0)
End If



End Sub
