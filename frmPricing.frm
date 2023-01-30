VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPricing 
   BackColor       =   &H80000013&
   Caption         =   "PRICING UPDATE"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8850
   Icon            =   "frmPricing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   8850
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.CheckBox Check1 
         Caption         =   "Specific Sno"
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtsno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6360
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
      Begin VB.PictureBox Picture5 
         Height          =   255
         Left            =   8040
         Picture         =   "frmPricing.frx":164A
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   12
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox txtBranchname 
         Height          =   285
         Left            =   2520
         TabIndex        =   11
         Text            =   "ALL BRANCHES"
         Top             =   1800
         Width           =   3015
      End
      Begin VB.ComboBox CboBcode 
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Text            =   "All"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtCurrentPrice 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtNewPrice 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton cmdcLOSE 
         Caption         =   "Close"
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   2760
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPStartFrom 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   1320
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   97648641
         CurrentDate     =   40095
      End
      Begin VB.Label lblsuppliername 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   15
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Label5 
         Caption         =   "Supplier No."
         Height          =   255
         Left            =   5280
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000018&
         Caption         =   "BranchCode"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000018&
         Caption         =   "Current Price"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000018&
         Caption         =   "New Price:"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000018&
         Caption         =   "Start From"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   1440
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmPricing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CboBcode_Click()
Set rs = oSaccoMaster.GetRecordset("select  * from d_branch where bcode='" & CboBcode & "'")
While Not rs.EOF
txtBranchName = rs.Fields("bname")
rs.MoveNext
Wend
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub cmdupdate_Click()
On Error GoTo ErrorHandler
Dim UpdateP As New ADODB.Connection
If Trim(txtNewPrice) = "" Then
MsgBox "Enter the new price."
txtNewPrice.SetFocus
Exit Sub
End If
If Not IsNumeric(txtNewPrice) Then
MsgBox "Please enter a number." & txtNewPrice & " is not a number", vbExclamation
txtNewPrice.SetFocus
Exit Sub
End If
If Check1.Value = vbChecked Then
   If lblsuppliername = "" Then
    MsgBox "Select sno to update Price", vbInformation
    Exit Sub
    End If
   
End If

With UpdateP
   .Open SelectedDsn
     .BeginTrans
         If Check1.Value = vbUnchecked Then
            sql = "Save_Price '" & Format(DTPStartFrom, "dd/mm/yyyy") & "'," & txtNewPrice & ""
            oSaccoMaster.ExecuteThis (sql)
            txtCurrentPrice = txtNewPrice
            txtNewPrice = ""
            '//select the date and branch
            If Trim$(CboBcode.Text) = "All" Then
                sql = "set dateformat dmy Update d_milkintake set ppu= " & CDbl(txtCurrentPrice) & " ,pamount=qsupplied  * " & CDbl(txtCurrentPrice) & " where transdate>= '" & Format(DTPStartFrom, "dd/mm/yyyy") & "' "
               oSaccoMaster.ExecuteThis (sql)
               .CommitTrans
                MsgBox "Records successively updated."
                frmPricing.Caption = "Pricing Updates"
            Exit Sub
        Else

         sql = "set dateformat dmy  Update d_milkintake set ppu=" & txtCurrentPrice & ",pamount=qsupplied  * " & txtCurrentPrice & "  where transdate>= '" & Format(DTPStartFrom, "dd/mm/yyyy") & "' and branchcode=" & Trim(CboBcode) & ""
           oSaccoMaster.ExecuteThis (sql)
           .CommitTrans
            MsgBox "Records successively updated."
            frmPricing.Caption = "Pricing Updates"
            Exit Sub
        End If
    Else
                  sql = "set dateformat dmy  Update d_milkintake set ppu=" & CDbl(txtNewPrice) & " ,pamount=qsupplied  * " & CDbl(txtNewPrice) & " where transdate>= '" & Format(DTPStartFrom, "dd/mm/yyyy") & "' and sno='" & txtSNo & "' "
            oSaccoMaster.ExecuteThis (sql)
            .CommitTrans
             MsgBox "Records successively updated."
             frmPricing.Caption = "Pricing Updates"
    Exit Sub
    End If
ErrorHandler:
MsgBox err.description
End With
End Sub

Private Sub Form_Load()
DTPStartFrom = Format(Get_Server_Date, "dd/mm/yyyy")
DTPStartFrom.MaxDate = Format(Get_Server_Date, "dd/mm/yyyy")
Dim bcode As Integer
Dim Bname As String
Set rs = CreateObject("adodb.recordset") '
    rs.Open "SELECT bcode FROM d_Branch order by bcode ", cn
    If rs.EOF Then Exit Sub
    With rs
        While Not .EOF
         CboBcode.AddItem rs.Fields(0)
         .MoveNext
        Wend
    End With
Set rs = New ADODB.Recordset
sql = "Pick_Current_Price"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
txtCurrentPrice = rs.Fields(0)
Else
txtCurrentPrice = 0
End If
txtCurrentPrice = Format(txtCurrentPrice, "####0.00")
End Sub

Private Sub Picture5_Click()
Me.MousePointer = vbHourglass
        frmSearchSupplier.Show vbModal
        txtSNo = sel
        txtSNo_KeyPress 13
        Me.MousePointer = 0
End Sub

Private Sub txtCurrentPrice_Validate(Cancel As Boolean)
txtCurrentPrice = Format(txtCurrentPrice, "####0.00")
End Sub

Private Sub txtNewPrice_Validate(Cancel As Boolean)
txtNewPrice = Format(txtNewPrice, "####0.00")
End Sub

Private Sub txtSNo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
    Exit Sub
End If
Set rs = New ADODB.Recordset
sql = "set dateformat dmy exec d_sp_SupplierEnquiry '" & txtSNo & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If rs.EOF Then
  lblsuppliername = ""
    MsgBox "There is no supplier with number " & txtSNo
    Exit Sub
End If
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then lblsuppliername = rs.Fields(0)
End If
sql = ""
sql = "select price from d_Price "
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
txtCurrentPrice = rs.Fields(0)
End If
End Sub

Private Sub txtSNo_LostFocus()
 Set rs = New ADODB.Recordset
    sql = "set dateformat dmy exec d_sp_SupplierEnquiry '" & txtSNo & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If rs.EOF Then
    lblsuppliername = ""
        MsgBox "There is no supplier with number " & txtSNo
        Exit Sub
    End If
    If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then lblsuppliername = rs.Fields(0)
    End If
    '//get the current price before you update it again
    
    sql = ""
    sql = "select price from d_Price "
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
    txtCurrentPrice = rs.Fields(0)
    End If

End Sub
