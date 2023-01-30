VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTransAssign 
   BackColor       =   &H00FFFF80&
   Caption         =   "Transport Assignment"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8280
   DrawStyle       =   3  'Dash-Dot
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   8280
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdreload 
      Caption         =   "Reload All Farmers"
      Height          =   375
      Left            =   5400
      TabIndex        =   23
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "Show  Farmers"
      Height          =   375
      Left            =   3960
      TabIndex        =   22
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtsubtransname 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   19
      Top             =   1560
      Width           =   5535
   End
   Begin VB.ComboBox CboSubTrans 
      Height          =   315
      Left            =   360
      TabIndex        =   18
      Top             =   1560
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   1800
      Picture         =   "frmTransAssign.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   17
      Top             =   2520
      Width           =   255
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   1800
      Picture         =   "frmTransAssign.frx":02C2
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   16
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton cmdActive 
      Caption         =   "In Activate"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   3240
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvwTransportassign 
      Height          =   3855
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6800
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   6480
      TabIndex        =   13
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdAssign 
      Caption         =   "Assign"
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox txtTCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtTNames 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   840
      Width           =   5535
   End
   Begin VB.TextBox txtSNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtAmount 
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00;(#,##0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox txtSNames 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   2520
      Width           =   5535
   End
   Begin MSComCtl2.DTPicker DTPDRemoved 
      Height          =   255
      Left            =   5880
      TabIndex        =   1
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   146735105
      CurrentDate     =   40096
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Sub_Trans Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   21
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Sub_Transporter Name"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2160
      TabIndex        =   20
      Top             =   1200
      Width           =   1650
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Transporter Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   480
      Width           =   1230
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Transporter Name"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2160
      TabIndex        =   10
      Top             =   480
      Width           =   1275
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Rate Per Kg"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   870
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Supplier Numer"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   1080
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Date of Assignment/Removal"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5400
      TabIndex        =   5
      Top             =   120
      Width           =   2085
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      Caption         =   "Supplier Name"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2160
      TabIndex        =   4
      Top             =   2160
      Width           =   1035
   End
End
Attribute VB_Name = "frmTransAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CboSubTrans_Change()
Set Rst = New ADODB.Recordset
sql = "d_sp_SelectSubTrans '" & CboSubTrans & "'"
Set Rst = oSaccoMaster.GetRecordset(sql)
If Not Rst.EOF Then
If Not IsNull(Rst.Fields(0)) Then txtsubtransname = Rst.Fields(0)
Rst.MoveNext
End If
End Sub

Private Sub CboSubTrans_Click()
CboSubTrans_Change
End Sub

Private Sub cmdActive_Click()
On Error GoTo ErrorHandler

'check if already assigned
sql = "set dateformat dmy select trans_code,sno,active,startdate,dateinactivate from d_transport where active=0 and sno='" & txtSNo & "' AND DateInactivate >= '" & DTPDRemoved & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
MsgBox "This supplier Number had been assigned to another transporter  "
Exit Sub
End If
Set rs = New ADODB.Recordset
Set cn = New ADODB.Connection
sql = "d_sp_CheckDate '" & txtSNo & "','" & txtTCode & "','" & DTPDRemoved & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
MsgBox "The transporter was assigned on " & rs.Fields("StartDate") & ".Please enter a valid date."
Exit Sub
End If

Set cn = New ADODB.Connection
sql = "SELECT  startdate FROM d_Transport WHERE  (Sno = '" & txtSNo & "') AND (Trans_Code = '" & txtTCode & "')"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If rs.Fields("StartDate") = DTPDRemoved Then
oSaccoMaster.ExecuteThis ("SET dateformat DMY delete FROM d_Transport where SNO= '" & txtSNo & "' and Trans_Code= '" & txtTCode & "' AND StartDate= '" & DTPDRemoved & "'")
MsgBox "Record removed "
End If
End If

Set cn = New ADODB.Connection
sql = "d_sp_InactivateTrans '" & txtTCode & "','" & txtSNo & "','" & DTPDRemoved & "'"
oSaccoMaster.ExecuteThis (sql)
loadTransportAssignments
If cmdActive.Caption = "Activate" Then
cmdActive.Caption = "In Activate"
Else
'cmdActive.Caption = "Activate"
End If
MsgBox "Records successively updated."
Exit Sub
ErrorHandler:
MsgBox err.description


End Sub

Private Sub cmdAssign_Click()
On Error GoTo ErrorHandler

If txtTCode = "" Then
MsgBox "Please enter the transporters code", vbInformation
txtTCode.SetFocus
Exit Sub
End If
'If CboSubTrans = "" Then
'MsgBox "Please enter the sub transporters code", vbInformation
'CboSubTrans.SetFocus
'Exit Sub
'End If
If txtSNo = "" Then
MsgBox "Please enter the supplier number", vbInformation
txtSNo.SetFocus
Exit Sub
End If
If txtAmount = "" Then
MsgBox "Please enter the rate per Kg.", vbInformation
txtAmount.SetFocus
Exit Sub
End If

'txtTCode_Validate True
If txtTNames = "" Then
MsgBox "Please enter an existing transporter's code.", vbInformation
txtTCode.SetFocus
Exit Sub
End If

'txtSNo_Validate True
If txtSNames = "" Then
MsgBox "Please enter an existing supplier's number.", vbInformation
txtSNo.SetFocus
Exit Sub
End If

If Not IsNumeric(txtAmount) Then
MsgBox "Please enter a numeric character. " & txtAmount & " is not a number.", vbExclamation
txtAmount.SetFocus
Exit Sub
End If

Set rs = New ADODB.Recordset
Set cn = New ADODB.Connection
sql = "set dateformat dmy select trans_code,sno,active,startdate,dateinactivate from d_transport where active=0 and sno='" & txtSNo & "' AND DateInactivate >= '" & DTPDRemoved & "'"
Set rs = oSaccoMaster.GetRecordset(sql)

Set rs = New ADODB.Recordset
Set cn = New ADODB.Connection
sql = "set dateformat dmy select trans_code,sno,active from d_transport where active=1 and sno='" & txtSNo & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
MsgBox "This supplier Number has been assigned to transporter code : " & rs.Fields("trans_code") & ""
Exit Sub
Else
'sql = "d_sp_TransAssign '" & txtTCode & "','" & CboSubTrans & "'," & txtSNo & "," & Format(txtAmount, "#0.00") & ",'" & DTPDRemoved & "','" & user & "'"
sql = "d_sp_TransAssign '" & txtTCode & "'," & txtSNo & "," & Format(txtAmount, "#0.00") & ",'" & DTPDRemoved & "','" & user & "'"
oSaccoMaster.ExecuteThis (sql)
End If

loadTransportAssignments
MsgBox "Records successively updated."
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub
Public Sub loadTransportAssignments()
    
    With lvwTransportassign
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs = CreateObject("adodb.recordset")
    
    sql = "Select * from d_Transport order by SNo, Trans_Code"
    
    Set rs = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
    cn.Open Provider, "bi"
    
    rs.Open sql, cn
    
    With lvwTransportassign
        
        .ColumnHeaders.Add , , "Trans Code"
        '.ColumnHeaders.Add , , "Sub_Trans Code"
        .ColumnHeaders.Add , , "SNo"
        .ColumnHeaders.Add , , "Rate"
        .ColumnHeaders.Add , , "Start Date"
        .ColumnHeaders.Add , , "Active"
        .ColumnHeaders.Add , , "Date InActived"
        While Not rs.EOF
        
            Set li = .ListItems.Add(, , Trim(rs.Fields("Trans_Code")))
           ' If Not IsNull(rs.Fields("Sub_Trans_Code")) Then li.ListSubItems.Add , , Trim(rs.Fields("Sub_Trans_Code"))
            If Not IsNull(rs.Fields("Sno")) Then li.ListSubItems.Add , , Trim(rs.Fields("Sno"))
            If Not IsNull(rs.Fields("Rate")) Then li.ListSubItems.Add , , Trim(rs.Fields("Rate"))
            If Not IsNull(rs.Fields("StartDate")) Then li.ListSubItems.Add , , Trim(rs.Fields("StartDate"))
            If Not IsNull(rs.Fields("Active")) Then li.ListSubItems.Add , , Trim(rs.Fields("Active"))
            If Not IsNull(rs.Fields("DateInactivate")) Then li.ListSubItems.Add , , Trim(rs.Fields("DateInactivate"))
            rs.MoveNext
        
        Wend
        
    End With
    
    rs.Close
    
    Set rs = Nothing
    
lvwTransportassign.View = lvwReport

End Sub
Sub LoadFarmers()
    
    With lvwTransportassign
    
        .ListItems.Clear
        
        .ColumnHeaders.Clear
    
    End With

    Set rs = CreateObject("adodb.recordset")
    
    sql = "Select * from d_Transport where Trans_Code='" & txtTCode & "' order by SNo, Trans_Code"
    
    Set rs = CreateObject("adodb.recordset")
    
    Set clsClass = New cdbase
    
    Provider = clsClass.OpenCon
    
    Set cn = CreateObject("adodb.connection")
    
    cn.Open Provider, "bi"
    
    rs.Open sql, cn
    
    With lvwTransportassign
        
        .ColumnHeaders.Add , , "Trans Code"
        '.ColumnHeaders.Add , , "Sub_Trans Code"
        .ColumnHeaders.Add , , "SNo"
        .ColumnHeaders.Add , , "Rate"
        .ColumnHeaders.Add , , "Start Date"
        .ColumnHeaders.Add , , "Active"
        .ColumnHeaders.Add , , "Date InActived"
        While Not rs.EOF
        
            Set li = .ListItems.Add(, , Trim(rs.Fields("Trans_Code")))
           ' If Not IsNull(rs.Fields("Sub_Trans_Code")) Then li.ListSubItems.Add , , Trim(rs.Fields("Sub_Trans_Code"))
            If Not IsNull(rs.Fields("Sno")) Then li.ListSubItems.Add , , Trim(rs.Fields("Sno"))
            If Not IsNull(rs.Fields("Rate")) Then li.ListSubItems.Add , , Trim(rs.Fields("Rate"))
            If Not IsNull(rs.Fields("StartDate")) Then li.ListSubItems.Add , , Trim(rs.Fields("StartDate"))
            If Not IsNull(rs.Fields("Active")) Then li.ListSubItems.Add , , Trim(rs.Fields("Active"))
            If Not IsNull(rs.Fields("DateInactivate")) Then li.ListSubItems.Add , , Trim(rs.Fields("DateInactivate"))
            rs.MoveNext
        
        Wend
        
    End With
    
    rs.Close
    
    Set rs = Nothing
    
lvwTransportassign.View = lvwReport
End Sub

Private Sub cmdreload_Click()
 Form_Load
End Sub

Private Sub cmdshow_Click()
 If txtTCode = "" Then
  MsgBox "Input Valid Transporter Number First", vbInformation
  Exit Sub
 End If
  txtTCode_Validate True
  If txtTNames <> "" Then
  LoadFarmers
  End If
End Sub

Private Sub Form_Load()
DTPDRemoved = Format(Get_Server_Date, "dd/mm/yyyy")
txtAmount = Format(0#, "#,###0.00")
loadTransportAssignments
End Sub

Public Sub edit(selected As String)
Dim myclass As cdbase
Set myclass = New cdbase
Set cn = CreateObject("adodb.connection")
Provider = myclass.OpenCon
cn.Open Provider, "BI"
Set rs = CreateObject("adodb.recordset")
sql = "select * from d_Transport where Trans_Code='" & selected & "' AND Sno ='" & lvwTransportassign.SelectedItem.ListSubItems(1).Text & "'"
rs.Open sql, cn
If Not rs.EOF Then
txtTCode = selected
txtSNo = rs!sno
txtAmount = rs!rate
End If
If rs!Active = True Then
cmdActive.Enabled = True
cmdActive.Caption = "In Activate"
Else
cmdActive.Enabled = True
cmdActive.Caption = "Activate"

End If
End Sub
Private Sub lvwTransportassign_DblClick()
edit lvwTransportassign.SelectedItem
txtSNo_Validate True
txtTCode_Validate True
End Sub

Private Sub Picture1_Click()
        Me.MousePointer = vbHourglass
        frmSearchSupplier.Show vbModal
        txtSNo = sel
        txtSNo_Validate True
        Me.MousePointer = 0
End Sub

Private Sub Picture5_Click()
        Me.MousePointer = vbHourglass
        frmSearchTransporter.Show vbModal
        txtTCode = sel
        txtTCode_Validate True
        Me.MousePointer = 0

End Sub



Private Sub txtAmount_Click()
If txtAmount = "0.00" Then
txtAmount = ""
End If
End Sub

Private Sub Txtamount_Validate(Cancel As Boolean)
txtAmount = Format(txtAmount, "#,###0.00")
End Sub

Private Sub txtSNo_Validate(Cancel As Boolean)
Dim a, t As Boolean
Set rs = New ADODB.Recordset
sql = "d_sp_SelectSuppliers '" & txtSNo & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(2)) Then txtSNames = rs.Fields(2)
Else
txtSNames = ""
End If

End Sub

Private Sub txtTCode_Validate(Cancel As Boolean)
Dim Transcode As String
Set rs = New ADODB.Recordset
sql = "d_sp_SelectTrans '" & txtTCode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtTNames = rs.Fields(0)
If Not IsNull(rs.Fields(1)) Then Transcode = rs.Fields(1)
Else
txtTNames = ""
End If
'//PUT THE NAME OF THE SUBTRANSPORTER
CboSubTrans.Clear
txtsubtransname.Text = ""
'Set rs = CreateObject("adodb.recordset")
'    sql = "SELECT subtranscode FROM d_subTransporters where transcode='" & txtTCode & "'"
'    Set rs = oSaccoMaster.GetRecordset(sql)
'    If rs.EOF Then Exit Sub
'    With rs
'        While Not .EOF
'         CboSubTrans.AddItem rs.Fields(0)
'         .MoveNext
'        Wend
'    End With
End Sub
