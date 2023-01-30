VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTellerDaily 
   BackColor       =   &H00C0C000&
   Caption         =   "Teller Daily Transactions"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10995
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTellerDailly.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   10995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFindacc 
      Caption         =   "<>"
      Height          =   315
      Left            =   3240
      TabIndex        =   30
      Top             =   960
      Width           =   465
   End
   Begin VB.CheckBox chkmonthly 
      Caption         =   "Period"
      Height          =   210
      Left            =   9000
      TabIndex        =   29
      Top             =   960
      Width           =   855
   End
   Begin VB.CheckBox chkselectall 
      Caption         =   "Select All"
      Height          =   435
      Left            =   9960
      TabIndex        =   26
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdReverse 
      Caption         =   "&Reverse"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8745
      TabIndex        =   25
      Top             =   7140
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10320
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTellerDailly.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   1058
      ButtonWidth     =   847
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Vourcher Listing"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Daily Summary"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.TextBox Text4 
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
      Left            =   7065
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   7590
      Width           =   1590
   End
   Begin VB.TextBox txtUnresDebits 
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
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   7590
      Width           =   1665
   End
   Begin VB.TextBox txtUnresCredits 
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
      Left            =   3345
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   7590
      Width           =   1665
   End
   Begin VB.TextBox txtResCredits 
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
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   7155
      Width           =   1665
   End
   Begin VB.TextBox txtResDebits 
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
      Left            =   5355
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   7155
      Width           =   1665
   End
   Begin VB.TextBox Text1 
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
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   7155
      Width           =   1590
   End
   Begin VB.TextBox txtDiff 
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
      Left            =   7095
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   8055
      Width           =   1590
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtWithdrawals 
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
      Left            =   5370
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   8055
      Width           =   1665
   End
   Begin VB.TextBox txtDeposits 
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
      Left            =   3375
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   8055
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   7
      Top             =   6885
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   135
      TabIndex        =   5
      Top             =   945
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
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
      Format          =   121110531
      CurrentDate     =   39564
   End
   Begin VB.TextBox txtTellerName 
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
      Left            =   3765
      TabIndex        =   3
      Top             =   945
      Width           =   3705
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
      Left            =   1590
      TabIndex        =   1
      Top             =   945
      Width           =   1545
   End
   Begin MSComctlLib.ListView lvwTransactions 
      Height          =   5505
      Left            =   75
      TabIndex        =   0
      Top             =   1350
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   9710
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "VoucherNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Account No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Credits"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Debits"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Trans Description"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "VNo"
         Object.Width           =   18
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   7560
      TabIndex        =   27
      Top             =   960
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
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
      Format          =   121110531
      CurrentDate     =   39564
   End
   Begin VB.Label lblclosingbal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   9600
      TabIndex        =   34
      Top             =   7680
      Width           =   1395
   End
   Begin VB.Label Label13 
      Caption         =   "Closing Bal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8760
      TabIndex        =   33
      Top             =   7680
      Width           =   825
   End
   Begin VB.Label txtBalByRange 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   32
      Top             =   8040
      Width           =   2235
   End
   Begin VB.Label Label11 
      Caption         =   "Open Bal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   31
      Top             =   7680
      Width           =   1065
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "EndDate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7680
      TabIndex        =   28
      Top             =   720
      Width           =   660
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Debits"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6480
      TabIndex        =   23
      Top             =   6870
      Width           =   525
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4410
      TabIndex        =   22
      Top             =   6870
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "UnCleared Vouchers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1560
      TabIndex        =   21
      Top             =   7635
      Width           =   1710
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Cleared Vouchers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1755
      TabIndex        =   20
      Top             =   7200
      Width           =   1500
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7890
      TabIndex        =   19
      Top             =   6960
      Width           =   60
   End
   Begin VB.Label lblDiff 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8850
      TabIndex        =   12
      Top             =   8085
      Width           =   60
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Totals"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2760
      TabIndex        =   9
      Top             =   8085
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Trans Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   6
      Top             =   720
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Account Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3795
      TabIndex        =   4
      Top             =   705
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Teller Account No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1620
      TabIndex        =   2
      Top             =   705
      Width           =   1455
   End
End
Attribute VB_Name = "frmTellerDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExport_Click()
    On Error GoTo sysError
    If lvwTransactions.ListItems.Count > 0 Then
        Dim rsDaily As New Recordset
        Set rsDaily = oSaccoMaster.GetRecordset("Exec Delete_DailyTrans")
        For I = 1 To lvwTransactions.ListItems.Count
            Set li = lvwTransactions.ListItems(I)
            If Not Save_Daily_Trans(CStr(li), CStr(li.SubItems(1)), CDbl(li.SubItems(2)), _
            CDbl(li.SubItems(3)), CStr(li.SubItems(4)), txtAccNo, UCase(txtTellerName), _
            DTPicker1, UCase(Current_User.UserID), 1, ErrorMessage) Then
                If ErrorMessage <> "" Then
                    MsgBox ErrorMessage, vbInformation, Me.Caption
                    ErrorMessage = ""
                    Exit Sub
                End If
            End If
        Next I
    Else
        MsgBox "There are no Transactions to be Printed", vbInformation, Me.Caption
        Exit Sub
    End If
    reportname = "Daily Teller.rpt"
    Show_Sales_Crystal_Report "", reportname, ""
    'MsgBox "Records Exported successfully", vbInformation, Me.Caption
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub chkselectall_Click()
If chkselectall.value = vbChecked Then
For I = 1 To lvwTransactions.ListItems.Count
    Set li = lvwTransactions.ListItems(I)
    If li.Checked = False Then
     li.Checked = True
    End If
    Next I
Else
For I = 1 To lvwTransactions.ListItems.Count
    Set li = lvwTransactions.ListItems(I)
    If li.Checked = True Then
     li.Checked = False
    End If
    Next I
End If
lvwTransactions_Click
End Sub

Private Sub cmdFindacc_Click()
  frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            txtAccNo.Text = SearchValue
            SearchValue = ""
            Continue = False
        End If
    End If
End Sub

Private Sub cmdReverse_Click()
    On Error GoTo sysError
    Dim VoucherNo As String, transdate As Date, accno As String
    If lvwTransactions.ListItems.Count > 0 Then
        If MsgBox("Do you want to reverse the selected entry?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
        VoucherNo = lvwTransactions.SelectedItem
        transdate = DTPicker1
        accno = lvwTransactions.SelectedItem.SubItems(1)
        If Not Execute_Command("Set DateFormat DMY Exec Reverse_Entry '" & VoucherNo _
        & "','" & transdate & "'", ErrorMessage) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
                Exit Sub
            End If
        End If
        MyRecord = txtAccNo
        txtAccNo_Change
    End If
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdsave_Click()
    Dim rsUpdate As New Recordset
    If MsgBox("Do you want to update the verified transactions?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
        Exit Sub
    End If
    For I = 1 To lvwTransactions.ListItems.Count
        If lvwTransactions.ListItems(I).Checked = True Then
            'XXXXXXXXXXXXXXX UPDATE POSTED XXXXXXXXXXXXXXXX
            Set rsUpdate = oSaccoMaster.GetRecordset("Update CustomerBalance Set Posted=1 " _
            & "Where CustomerBalanceID=" & CLng(lvwTransactions.ListItems(I).SubItems(5)))
        Else
            Set rsUpdate = oSaccoMaster.GetRecordset("Update CustomerBalance Set Posted=0 " _
            & "Where CustomerBalanceID=" & CLng(lvwTransactions.ListItems(I).SubItems(5)))
        End If
    Next I
    MsgBox "Records Updated Successfully", vbInformation, Me.Caption
End Sub

Private Sub DTPicker1_Change()
    txtAccNo_Change
End Sub

Private Sub Form_Load()
    DTPicker1 = Format(Get_Server_Date, " dd-MM-yyyy")
    DTPicker2 = DTPicker1
    If MyRecord <> "" Then
        txtAccNo = MyRecord
    End If
    MyRecord = ""
End Sub





Private Sub lvwTransactions_Click()
    On Error GoTo sysError
    Dim Credits As Double, Debits As Double
    If lvwTransactions.ListItems.Count > 0 Then
        For I = 1 To lvwTransactions.ListItems.Count
            Set li = lvwTransactions.ListItems(I)
            If lvwTransactions.ListItems(I).Checked = True Then
                Credits = Credits + CDbl(li.SubItems(2))
                Debits = Debits + CDbl(li.SubItems(3))
            End If
        Next I
        txtResCredits = Format$(Credits, Cfmt)
        txtResDebits = Format(Debits, Cfmt)
        txtUnresCredits = Format(CDbl(txtDeposits) - Credits, Cfmt)
        txtUnresDebits = Format(CDbl(txtWithdrawals) - Debits, Cfmt)
    End If
    Debits = 0
    Credits = 0
    Exit Sub
sysError:
    MsgBox err.description
End Sub

Private Sub lvwTransactions_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo sysError
    Dim Credits As Double, Debits As Double
    For I = 1 To lvwTransactions.ListItems.Count
        Set li = lvwTransactions.ListItems(I)
        If lvwTransactions.ListItems(I).Checked = True Then
            Credits = Credits + CDbl(li.SubItems(2))
            Debits = Debits + CDbl(li.SubItems(3))
        End If
    Next I
    txtResCredits = Format$(Credits, Cfmt)
    txtResDebits = Format(Debits, Cfmt)
    txtUnresCredits = Format(CDbl(txtDeposits) - Credits, Cfmt)
    txtUnresDebits = Format(CDbl(txtWithdrawals) - Debits, Cfmt)
    Debits = 0
    Credits = 0
    Exit Sub
sysError:
    MsgBox err.description
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    On Error GoTo sysError
    Dim rsCheques As New Recordset, VoucherNo As String, CustAccNo As String, _
    Amount As Double
    If lvwTransactions.ListItems.Count > 0 Then
        Dim rsDaily As New Recordset
        Set rsDaily = oSaccoMaster.GetRecordset("Exec Delete_DailyTrans")
        For I = 1 To lvwTransactions.ListItems.Count
            Set li = lvwTransactions.ListItems(I)
            If Not Save_Daily_Trans(CStr(li), CStr(li.SubItems(1)), CDbl(li.SubItems(2)), _
            CDbl(li.SubItems(3)), CStr(li.SubItems(4)), txtAccNo, UCase(txtTellerName), _
            DTPicker1, UCase(Current_User.UserID), 1, ErrorMessage) Then
                If ErrorMessage <> "" Then
                    MsgBox ErrorMessage, vbInformation, Me.Caption
                    ErrorMessage = ""
                    Exit Sub
                End If
            End If
        Next I
    Else
        MsgBox "There are no Transactions to be Printed", vbInformation, Me.Caption
        Exit Sub
    End If
    
    Select Case UCase(ButtonMenu.Text)
        Case "VOURCHER LISTING"
        reportname = "Daily Teller.rpt"
        Case "DAILY SUMMARY"
'        Dim rsTrans As New Recordset
'        Set rsTrans = oSaccoMaster.GetRecordset("Set DateFormat DMY Exec Get_Teller_Daily_Summary '" _
'        & txtAccNo & "','" & DTPicker1 & "'")
'        If Not Save_Teller_Trans(txtAccNo, DTPicker1, rsTrans, ErrorMessage) Then
'            MsgBox ErrorMessage, vbInformation, Me.Caption
'            ErrorMessage = ""
'        End If
        reportname = "Daily Teller Summary.rpt"
    End Select
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
    reportname = ""
    STRFORMULA = ""
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtAccNo_Change()
    If Trim$(txtAccNo) <> "" Then
        Dim rsTrans As New Recordset, teller As String
        Dim UserID As String, RangeOpeningBal As Double, Bal As Double
        Dim rsAcc As New Recordset, Deposits As Double, _
        Withdrawals As Double, DIFF As Double, Contra As String
        lvwTransactions.ListItems.clear
        lblDiff = ""
        Set rsAcc = oSaccoMaster.GetRecordset("Select * From GlSetup where AccNo='" _
        & txtAccNo & "'")
        With rsAcc
            If Not .EOF Then
                UserID = IIf(IsNull(!GlAccName), "", !GlAccName)
                teller = UserID
                If UserID <> "" Then
                    RangeOpeningBal = getGlBalance(txtAccNo, DTPicker1, DTPicker1)
                    txtBalByRange = RangeOpeningBal
                    If chkmonthly.value = vbChecked Then
                    If DTPicker1 > DTPicker2 Then
                     MsgBox "EndDate Should be Greater than FirstDate", vbInformation, Me.Caption
                     Exit Sub
                    End If
                    Set rsTrans = Get_MonthlyTeller_Transactions(txtAccNo, txtAccNo, DTPicker1, DTPicker2)
                    Else
                    Set rsTrans = Get_Teller_Transactions(txtAccNo, txtAccNo, DTPicker1)
                    End If
                    txtTellerName = UserID
                End If
                With rsTrans
                    If .State = adStateOpen Then
                        While Not .EOF
'                            If !TransDescription <> "Stamp Duty" Then
'                                If !TransDescription <> "Stamp D/C" Then
'                                    If !TransDescription <> "Comm/charges" Then
'                                        If !TransDescription <> "CW/Charges" Then
'                                            If !TransDescription <> "N Charges" Then
                                                '                                            If !TransDescription <> "N Charges" Then
                                                Set li = lvwTransactions.ListItems.Add(, , IIf(IsNull(!DocumentNo), "", !DocumentNo))
                                                li.SubItems(1) = IIf(IsNull(!Source), "", !Source)
                                                If LCase(!Craccno) = LCase(txtAccNo) Then
                                                    li.SubItems(2) = Format(IIf(IsNull(!Amount), 0, !Amount), Cfmt)
                                                    li.SubItems(3) = "0.00"
                                                Else
                                                    li.SubItems(3) = Format((IIf(IsNull(!Amount), 0, !Amount)), Cfmt)
                                                    li.SubItems(2) = "0.00"
                                                End If
                                                Deposits = Deposits + CDbl(li.SubItems(2))
                                                Withdrawals = Withdrawals + CDbl(li.SubItems(3))
                                                li.SubItems(4) = IIf(IsNull(!TransDescript), "", !TransDescript)
                                                li.SubItems(5) = !id
                                                If !ReconId = True Then
                                                    li.Checked = True
                                                End If
'                                            End If
'                                        End If
'                                    End If
'                                End If
'                            End If
                            .MoveNext
                        Wend
                        txtDeposits = Format(Deposits, Cfmt)
                        txtWithdrawals = Format(Withdrawals, Cfmt)
                        DIFF = Deposits - Withdrawals
                        Bal = RangeOpeningBal - DIFF
                        If DIFF > 0 Then
                            lblDiff = "Overs"
                        ElseIf DIFF < 0 Then
                            lblDiff = "BALANCE"
                            DIFF = DIFF * (-1)
                        End If
                        
                        txtDiff = Format(DIFF, Cfmt)
                        lblclosingbal = Format(Bal, Cfmt)
                        Deposits = 0
                        Withdrawals = 0
                        DIFF = 0
                    End If
                End With
            End If
        End With
    End If
End Sub

