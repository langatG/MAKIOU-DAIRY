VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmtransactionlisting 
   Caption         =   "CREDITORS PAYMENTS"
   ClientHeight    =   8490
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraOtherpayment 
      Caption         =   "Transaction Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   0
      TabIndex        =   24
      Top             =   3000
      Width           =   15735
      Begin VB.CheckBox chkperiodicreceipts 
         Caption         =   "Print Period Vouchers"
         Height          =   255
         Left            =   4920
         TabIndex        =   35
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CommandButton cmdprint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3840
         TabIndex        =   34
         Top             =   1680
         Width           =   1020
      End
      Begin VB.TextBox txtNarration 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "##-##-####"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3000
         TabIndex        =   33
         ToolTipText     =   "The person who is taking the money"
         Top             =   600
         Width           =   3510
      End
      Begin VB.TextBox TxtOtherPAcc 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "##-##-####"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
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
         Left            =   1320
         TabIndex        =   32
         Top             =   210
         Width           =   1305
      End
      Begin VB.PictureBox Picture21 
         Height          =   285
         Left            =   2685
         Picture         =   "frmtransactionlisting.frx":0000
         ScaleHeight     =   225
         ScaleWidth      =   240
         TabIndex        =   31
         Top             =   225
         Width           =   300
      End
      Begin VB.TextBox TxtOtherPayment 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   2985
         MaxLength       =   12
         TabIndex        =   30
         Text            =   "0"
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   29
         Top             =   1680
         Width           =   990
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2610
         TabIndex        =   28
         Top             =   1665
         Width           =   1095
      End
      Begin VB.CommandButton cmdpost 
         Caption         =   "Post"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1380
         TabIndex        =   27
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Frame Frame2 
         Caption         =   "Unpaid Invoices"
         Height          =   4935
         Left            =   9120
         TabIndex        =   25
         Top             =   360
         Width           =   6135
         Begin MSComctlLib.ListView lvwItems 
            Height          =   4455
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   7858
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "RNO"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "INVOICE NO"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Amount"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "GL Account No"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin MSComctlLib.ListView lvwTrans 
         Height          =   3255
         Left            =   120
         TabIndex        =   36
         Top             =   2040
         Width           =   8925
         _ExtentX        =   15743
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TransDate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Amount"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Dr AccNo"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cr AccNo"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DocumentNo"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Source"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Narration"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "DocPosted"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Cheque No."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Payee:"
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
         Left            =   2280
         TabIndex        =   40
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lblOtherPaymentAcc 
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
         Left            =   3000
         TabIndex        =   39
         Top             =   210
         Width           =   3495
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
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
         Index           =   4
         Left            =   2280
         TabIndex        =   38
         Top             =   990
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "A/C"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2985
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15735
      Begin VB.Frame Frame 
         ClipControls    =   0   'False
         Height          =   1695
         Left            =   225
         TabIndex        =   2
         Top             =   690
         Width           =   14895
         Begin VB.CommandButton Command2 
            Caption         =   "Print Statement"
            Height          =   315
            Left            =   11400
            TabIndex        =   43
            Top             =   1080
            Width           =   2415
         End
         Begin VB.ComboBox cbocreditors 
            Height          =   315
            Left            =   11280
            TabIndex        =   42
            Top             =   360
            Width           =   3015
         End
         Begin VB.OptionButton Optcheque 
            Caption         =   "Cheque"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1080
            TabIndex        =   14
            Top             =   300
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Optcash 
            Caption         =   "Cash"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   13
            Top             =   300
            Width           =   855
         End
         Begin VB.Frame Frame7 
            Height          =   570
            Left            =   2130
            TabIndex        =   8
            Top             =   120
            Width           =   5895
            Begin VB.TextBox txtReceiptsno 
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
               Left            =   1200
               TabIndex        =   10
               Top             =   195
               Width           =   1935
            End
            Begin VB.TextBox txtChequeno 
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
               Left            =   4065
               TabIndex        =   9
               Top             =   195
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.Label lblCheque 
               AutoSize        =   -1  'True
               Caption         =   "Cheque No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   3240
               TabIndex        =   12
               Top             =   240
               Visible         =   0   'False
               Width           =   795
            End
            Begin VB.Label lblVoucher 
               AutoSize        =   -1  'True
               Caption         =   "Voucher No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   210
               TabIndex        =   11
               Top             =   225
               Width           =   870
            End
         End
         Begin VB.TextBox txtAmountDue 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   8100
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   7
            Top             =   780
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.TextBox TxtDRAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   2130
            Locked          =   -1  'True
            MaxLength       =   9
            TabIndex        =   6
            Text            =   "0"
            Top             =   1215
            Width           =   1410
         End
         Begin VB.PictureBox Picture4 
            Height          =   285
            Left            =   3525
            Picture         =   "frmtransactionlisting.frx":02C2
            ScaleHeight     =   225
            ScaleWidth      =   240
            TabIndex        =   5
            Top             =   780
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
            Height          =   300
            Left            =   2130
            TabIndex        =   4
            Top             =   780
            Width           =   1410
         End
         Begin VB.CommandButton cmdnew 
            Caption         =   "New"
            Height          =   330
            Left            =   8085
            TabIndex        =   3
            Top             =   210
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Creditors"
            Height          =   255
            Left            =   10080
            TabIndex        =   41
            Top             =   360
            Width           =   975
         End
         Begin VB.Label LblStatus 
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
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   4905
            TabIndex        =   19
            Top             =   1200
            Width           =   1320
         End
         Begin VB.Label Labal 
            Caption         =   "Avaliable Amount"
            Height          =   255
            Left            =   750
            TabIndex        =   18
            Top             =   1260
            Width           =   1335
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
            Left            =   3825
            TabIndex        =   17
            Top             =   780
            Width           =   4170
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Bank Account"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   570
            TabIndex        =   16
            Top             =   825
            Width           =   1020
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Account Status"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3735
            TabIndex        =   15
            Top             =   1245
            Width           =   1125
         End
      End
      Begin VB.TextBox txtpaymentinrespectof 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   2520
         Width           =   13095
      End
      Begin MSComCtl2.DTPicker dtptransdate 
         Height          =   315
         Left            =   1635
         TabIndex        =   20
         Top             =   270
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   105185283
         CurrentDate     =   39954
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Transaction Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   345
         TabIndex        =   23
         Top             =   330
         Width           =   1230
      End
      Begin VB.Label Label1 
         Caption         =   "ACCOUNTS PAYABLES- CREDITORS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   3120
         TabIndex        =   22
         Top             =   240
         Width           =   7335
      End
      Begin VB.Label Label4 
         Caption         =   "Payment in Respect Of:"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   2520
         Width           =   1935
      End
   End
   Begin VB.Menu mnufiles 
      Caption         =   "File"
      Begin VB.Menu mnuagingreport 
         Caption         =   "Aging Report"
      End
   End
End
Attribute VB_Name = "frmtransactionlisting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DRaccno As String
Dim Craccno As String
Dim IntAccNo As String
Dim glmemno As String
Dim glnamE1 As String
Dim Amount As Currency
Dim ReceiptNo As String
Dim mysql As String
Dim loanno1 As String
Dim loanno2 As String
Dim loanno3 As String
Dim loanno4 As String
Dim loanno5 As String
Dim loanno6 As String
Dim loanno7 As String
Dim selected1 As String
Private Sub cboVendor_Change()
Lvwitems.ListItems.Clear
If Trim(cboVendor) = "" Then
Exit Sub
End If
Set rs = oSaccoMaster.GetRecordset("d_sp_InvVendor '" & cboVendor & "'")
While Not rs.EOF
If Not IsNull(rs.Fields(0)) Then
Set li = Lvwitems.ListItems.Add(, , rs.Fields(0))
End If
If Not IsNull(rs.Fields(1)) Then li.SubItems(1) = rs.Fields(1) & ""
If Not IsNull(rs.Fields(2)) Then li.SubItems(2) = rs.Fields(2) & ""
If Not IsNull(rs.Fields(3)) Then li.SubItems(3) = rs.Fields(3) & ""
rs.MoveNext
Wend
End Sub

Private Sub cboVendor_Click()
cboVendor_Change
End Sub

Private Sub Chkpaymentapproved_Click()
If Chkpaymentapproved = vbChecked Then
cboVendor.Visible = True
Frame2.Visible = True
Vendor.Visible = True
cboVendor_Change
Else
cboVendor.Visible = False
Frame2.Visible = False
Vendor.Visible = False
End If
End Sub

Private Sub cmdadd33_Click()
If Lvwitems.ListItems.Count = 0 Then
    MsgBox "There are no items", vbInformation, "NO ITEMS"
        Lvwitems.SetFocus
    Exit Sub
End If

'Set li = lvwselecteditems.ListItems.Add(, , Lvwitems.SelectedItem)
'                        li.SubItems(1) = Lvwitems.SelectedItem.ListSubItems(1) & ""
'                        li.SubItems(2) = Lvwitems.SelectedItem.ListSubItems(2) & ""
'                        li.SubItems(3) = Lvwitems.SelectedItem.ListSubItems(3) & ""
'                        li.SubItems(4) = 0# & ""
'
'Lvwitems.ListItems.Remove (Lvwitems.SelectedItem.Index)
'//get the checking account
sql = ""
sql = "SELECT     *   FROM         d_Approve2  WHERE     (RNo = '" & li & "') and approved=1   ORDER BY id DESC"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
TxtOtherPAcc = rs.Fields("glacc")
End If
Set li = lvwTrans.ListItems.Add(, , dtptransdate)
    li.SubItems(1) = Format(CDbl(TxtOtherPayment), "#,##0.00")
    li.SubItems(2) = TxtOtherPAcc
    li.SubItems(3) = txtcontra
    li.SubItems(4) = txtReceiptsno
    li.SubItems(5) = TxtOtherPAcc & "-" & (lblOtherPaymentAcc)
    li.SubItems(6) = txtNarration
    li.SubItems(7) = 1
    li.SubItems(8) = txtChequeno

'lvwSelectedItems_DblClick
End Sub

Private Sub cmdnew_Click()
Dim rsr As New ADODB.Recordset
Dim rsg As New ADODB.Recordset
Dim I As Object
Dim Mylength As Integer
'//if this record is new then look for receipts no

''//clear all textboxes





mysql = ""
mysql = "select GenerateReceiptno from param"

Set rsg = oSaccoMaster.GetRecordset(mysql)
If Not rsg.EOF Then
    ''''check check
    If rsg!GenerateReceiptno = True Then
    
        mysql = ""
        mysql = "select top 1 * from Receiptno where receiptno like 'PV-%' order by Receipthnoid desc"
        
        Set rsr = oSaccoMaster.GetRecordset(mysql)
        
        If Not rsr.EOF Then
            Mylength = CInt(Mid(rsr!ReceiptNo, 5, 10))
            Mylength = Mylength + 1
            txtReceiptsno = Padding(Mylength)
            txtReceiptsno = "PV-" & txtReceiptsno
        Else
            Mylength = 1
            txtReceiptsno = "PV-" & Padding(Mylength)
            
        End If
Else
    ''//receiptno  will be keyed in
End If
End If

End Sub

Private Sub cmdPost_Click()
    On Error GoTo SysError
    If Check_Period_If_Closed(dtptransdate) = True Then
         Exit Sub
     End If
    Dim Cubaccount As Cub_Acc_Details
    Dim Account As Acc_Details
    Dim chequeno As String
    Dim DRaccno As String, Craccno As String, Amount As Double, transdate As Date, _
    TransDescription As String, TransSource As String, DocumentNo As String, CashBook As String, doc_posted As String, IDENTI As Long
    If lvwTrans.ListItems.Count > 0 Then
        If MsgBox("Do you want post the entry?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
    Else
        MsgBox "There are no transactions to be posted", vbInformation, Me.Caption
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    For I = 1 To lvwTrans.ListItems.Count
    ' If lvwTrans.ListItems.Item(I).Checked = True Then
        Set li = lvwTrans.ListItems(I)
        transdate = li
        Amount = CDbl(lvwTrans.ListItems(I).SubItems(1))
        DRaccno = lvwTrans.ListItems(I).SubItems(2)
        Craccno = lvwTrans.ListItems(I).SubItems(3)
        DocumentNo = lvwTrans.ListItems(I).SubItems(4)
        TransSource = lvwTrans.ListItems(I).SubItems(5)
        TransDescription = lvwTrans.ListItems(I).SubItems(6)
        chequeno = lvwTrans.ListItems(I).SubItems(8) ' chequeno
        doc_posted = lvwTrans.ListItems(I).SubItems(7)
        GetTransactionNo
'IDENTI = lvwTrans.ListItems(i).SubItems(9)
        CashBook = 1
        If DocumentNo = "" Then DocumentNo = "CB"

                If chknonmemberpostings = vbChecked Then
                doc_posted = 1
                Else
                doc_posted = 0
                End If
        
        Set rs = oSaccoMaster.GetRecordset("sp_chequeno_used '" & Trim(chequeno) & "','" & Trim(TransSource) & "'")
        If Not rs.EOF Then

         End If

       
        
        If Not Save_GLTRANSACTION(transdate, Amount, DRaccno, Craccno, DocumentNo, _
        TransSource, User, ErrorMessage, TransDescription, CashBook, doc_posted, chequeno, transactionNo, "") Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        
        
         sql = " INSERT INTO BankAccount"
         sql = sql & "             (transdate, AccName, Pvcno, Amount, Naration, auditid,PIRO,chequeno)"
         sql = sql & "  VALUES     ('" & transdate & "','" & TransSource & "','" & DocumentNo & "'," & Amount & ",'" & TransDescription & "','" & User & "','" & txtpaymentinrespectof & "','" & chequeno & "')"
         oSaccoMaster.ExecuteThis (sql)
        
       ' End If
    Next I
    
    '//clear listview
    
          sql = "Update d_Requisition set [status]='Paid'   where RNo='" & selected1 & "'"
  oSaccoMaster.ExecuteThis (sql)
          sql = "Update InvoiceReceived set paid =1   where RNo='" & selected1 & "'"
  oSaccoMaster.ExecuteThis (sql)
  'Update InvoiceReceived set paid =1 where rno=''
    mysql = ""
    mysql = "set dateformat dmy Insert into Receiptno(Receiptno,Auditdate,auditid)values('" & txtReceiptsno & "','" & Format(Get_Server_Date, "dd/MM/yyyy") & "','" & User & "')"
    oSaccoMaster.ExecuteThis (mysql)
    lvwTrans.ListItems.Clear
    Form_Load
    Me.MousePointer = vbDefault
    MsgBox "Posting Successfull", vbInformation, Me.Caption
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption
    Me.MousePointer = vbDefault

End Sub

Private Sub cmdprint_Click()
'//pettycashvoucher
If chkperiodicreceipts = vbChecked Then
 'STRFORMULA = "{PettyCash.Pvcno}='" & txtReceiptsno & "'"
    reportname = "bankvoucherlistings.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
'//periodics
Else
    STRFORMULA = "{PettyCash.Pvcno}='" & txtReceiptsno & "'"
    reportname = "bankvoucher.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
'//periodics
'pettycashvoucherperiodic
'PV-000002
End If
End Sub

Private Sub cmdsave_Click()
    On Error GoTo SysError
    If Trim$(CCur(TxtOtherPayment)) > CCur(TxtDRAmount) Then
        MsgBox "You do not have sufficient Amount in Petty Cash Account", vbInformation, Me.Caption
        Exit Sub
    End If
    
    If Trim(txtChequeno) = "" Then
       ' MsgBox "Please Enter The chequne No", vbInformation, Me.Caption
       ' Exit Sub
    End If
    
    If Val(TxtOtherPayment) = 0 Then
        MsgBox "Please enter a valid Amount", vbInformation, Me.Caption
        txtAmount.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim$(txtcontra) = "" Then
        MsgBox "Please enter the Account to Debit.", vbInformation, Me.Caption
        txtDrAccNo.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If txtnarations = "" Then
    'MsgBox "Please enter the naration", vbCritical
    'Exit Sub
    End If
    If Trim$(TxtOtherPAcc) = "" Then
        MsgBox "Please enter the Account to Credit.", vbInformation, Me.Caption
        txtCrAccNo.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    
    If Trim$(txtReceiptsno) = "" Then
        MsgBox "Please enter the Transaction Source", vbInformation, Me.Caption
        txtSource.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If Trim$(txtNarration) = "" Then
        MsgBox "Please enter the Transaction Description", vbInformation, Me.Caption
        txtNarration.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
   TxtDRAmount = CCur(TxtDRAmount) - CCur(TxtOtherPayment)
    Set li = lvwTrans.ListItems.Add(, , dtptransdate)
    li.SubItems(1) = Format(CDbl(TxtOtherPayment), "#,##0.00")
    li.SubItems(2) = TxtOtherPAcc
    li.SubItems(3) = txtcontra
    li.SubItems(4) = txtReceiptsno
    li.SubItems(5) = TxtOtherPAcc & "-" & (lblOtherPaymentAcc)
    li.SubItems(6) = txtNarration
    li.SubItems(7) = 1
    li.SubItems(8) = txtChequeno
    TxtOtherPayment = "0"
    TxtOtherPAcc = ""
    
    'txtReceiptsno = ""
    'txtNarration = ""
    txtnarations = ""
    Exit Sub
    'lblDrAc = ""
    lblOtherPaymentAcc = ""
    'txtchequeno = ""
SysError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
'//start the staff here
Dim invoiceno As String
Dim Amount As Double
Dim dcode As String
Dim rsp As New ADODB.Recordset
Dim gl As String
Dim glamount As Double
Dim days As Integer
Dim sdate As Date
Dim invdate As Date
Dim aamount As Double
Dim costcentre As String
Dim dramount As Double, cramount As Double, transdate As Date, descrip As String, balance As Double
Dim d1 As Integer, d2 As Integer, d3 As Integer, d4 As Integer, d5 As Integer
sql = ""
sql = "truncate table d_creditors_Statemetns"
oSaccoMaster.ExecuteThis (sql)

'first loop then second loop
sql = "SELECT a.RNo,b.InvoiceNo,b.Transdate,b.Amount ,a.Description,b.CreditorAccNo FROM d_Requisition a INNER JOIN InvoiceReceived b ON a.RNo=b.RNO WHERE a.CostCentre ='" & cbocreditors & "' order by a.RNo"
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
dcode = Trim(rs.Fields(0))
invoiceno = rs.Fields(1)
cracmount = rs.Fields(3)
dramount = 0
gl = rs.Fields(5)
transdate = rs.Fields(2)
desrip = rs.Fields(4)
balance = cracmount + balance
        '//save the invoce
        
        sql = ""
        sql = "set dateformat dmy INSERT INTO d_creditors_Statemetns(dcode,invoiceno,dramount,cramount,transdate,descrip , balance,CNAME) VALUES('" & Trim(dcode) & "','" & invoiceno & "'," & dramount & "," & cracmount & ",'" & transdate & "','" & desrip & "' ," & balance & ",'" & cbocreditors & "')"
        oSaccoMaster.ExecuteThis (sql)
        '/check if the amount is paid
        Dim dpaid As Date
        sql = ""
        sql = "SELECT amount,transdate FROM  GLTRANSACTIONS WHERE DrAccNo ='" & gl & "' AND DocumentNo='" & invoiceno & "'"
        Set rsp = oSaccoMaster.GetRecordset(sql)
        If Not rsp.EOF Then
        glamount = rsp.Fields(0)
        dpaid = rsp.Fields(1)
        cracmount = 0
       dramount = glamount
       balance = balance - dramount
       desrip = "Paid on" & dpaid
        sql = ""
        sql = "set dateformat dmy INSERT INTO d_creditors_Statemetns(dcode,invoiceno,dramount,cramount,transdate,descrip , balance,CNAME) VALUES('" & dcode & "','" & invoiceno & "'," & dramount & "," & cramount & ",'" & dpaid & "','" & desrip & "' ," & balance & ",'" & cbocreditors & "')"
        oSaccoMaster.ExecuteThis (sql)
        Else
        glamount = 0
        balance = balance
        End If
        
        cracmount = 0
        dramount = 0
        glamount = 0
      

rs.MoveNext
Wend

'CSTATEMENTS.rpt

 reportname = "CSTATEMENTS.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title

End Sub

Private Sub Form_Load()
Dim RsLoancode As New ADODB.Recordset
    Dim RsScheme As New ADODB.Recordset
    Dim rsCompany As New ADODB.Recordset
    'optCash_Click
    sql = ""
    Dim rs2 As New ADODB.Recordset
sql = "SELECT DISTINCT companyname  FROM ag_Supplier1"
Set rs2 = oSaccoMaster.GetRecordset(sql)
While Not rs2.EOF

cbocreditors.AddItem rs2.Fields(0)
rs2.MoveNext
Wend
    dtptransdate.Value = Format(Get_Server_Date, "dd/MM/yyyy")

    FraOtherpayment.Visible = True
    lblCheque.Visible = True
    Optcheque.Value = True
    optCash_Click
    'Load_Data
    
    Lvwitems.ListItems.Clear
    txtChequeno.Visible = True
    lblVoucher.Visible = True
Set rs = oSaccoMaster.GetRecordset("select rno,invoiceno,amount,creditoraccno from InvoiceReceived WHERE PAID=0 and rno is not null")
While Not rs.EOF
Set li = Lvwitems.ListItems.Add(, , rs.Fields(0))
                        li.SubItems(1) = rs.Fields(1) & ""
                        li.SubItems(2) = rs.Fields(2) & ""
                        li.SubItems(3) = rs.Fields(3) & ""
                       


rs.MoveNext
Wend

cboVendor = "<Select Vendor>"

End Sub

Private Sub lvwItems_DblClick()
edit Lvwitems.SelectedItem
End Sub
Public Sub edit(selected As String)
'//
Dim costcentre As String
sql = ""
sql = "SELECT rno,invoiceno,amount,creditoraccno FROM InvoiceReceived WHERE RNo='" & Trim(selected) & "' "
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
TxtOtherPayment = rs.Fields("amount")
TxtOtherPAcc = rs.Fields("creditoraccno")
selected1 = Trim(selected)
'//GET THE GL NAME BEFORE POSTINGS
'sql = ""
'sql = "SELECT GLNO FROM ag_Supplier1 WHERE CompanyName ='" & txtsupplier & "'"
'Set Rst = oSaccoMaster.GetRecordset(sql)
'If Not Rst.EOF Then
'txtcreditorAcc = Rst.Fields(0)
'End If

sql = ""
sql = "SELECT * FROM d_Requisition WHERE RNo='" & Trim(selected) & "' "
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then

costcentre = rs.Fields("costcentre")
txtNarration = costcentre
End If

Else
txtoriginalamount = ""
txtsupplier = ""
End If
End Sub

Private Sub mnuagingreport_Click()
'//start the staff here
Dim invoiceno As String
Dim Amount As Double
Dim dcode As String
Dim rsp As New ADODB.Recordset
Dim gl As String
Dim glamount As Double
Dim days As Integer
Dim sdate As Date
Dim invdate As Date
Dim aamount As Double
Dim costcentre As String
Dim d1 As Integer, d2 As Integer, d3 As Integer, d4 As Integer, d5 As Integer
sql = ""
sql = "truncate table d_creditors_aging"
oSaccoMaster.ExecuteThis (sql)

'first loop then second loop
sql = "select invoiceno from InvoiceReceived where PAID=0 order by invoiceno"
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
invoiceno = rs.Fields(0)
        sql = ""
        sql = "set dateformat dmy SELECT invoiceno,creditoraccno,amount,transdate,rno FROM InvoiceReceived  where invoiceno=" & invoiceno & " AND PAID=0 ORDER BY Invoiceno,rno"
        Set Rst = oSaccoMaster.GetRecordset(sql)
        While Not Rst.EOF
        invoiceno = Rst.Fields(0)
        Amount = Rst.Fields(2)
        dcode = Trim(Rst.Fields(4))
        gl = Rst.Fields(1)
        sdate = Format(Get_Server_Date, "dd / mm / yyyy")
        invdate = Rst.Fields(3)
        days = DateDiff("d", invdate, sdate)
        '//get the cost centre
        sql = ""
        Dim rs1 As New ADODB.Recordset
        Dim rs2 As New ADODB.Recordset
sql = "SELECT * FROM d_Requisition WHERE RNo='" & Trim(dcode) & "' "
Set rs1 = oSaccoMaster.GetRecordset(sql)
If Not rs1.EOF Then
Dim costcentre1 As String
costcentre = rs1.Fields("costcentre")

sql = ""
sql = "SELECT SUPPLIERID  FROM ag_Supplier1 WHERE CompanyName='" & Trim(costcentre) & "' "
Set rs2 = oSaccoMaster.GetRecordset(sql)
If Not rs2.EOF Then

costcentre1 = rs2.Fields("SUPPLIERID")

End If

End If

        '/check if the amount is paid
        sql = ""
        sql = "SELECT amount FROM  GLTRANSACTIONS WHERE DrAccNo ='" & gl & "' AND DocumentNo='" & invoiceno & "'"
        Set rsp = oSaccoMaster.GetRecordset(sql)
        If Not rsp.EOF Then
        glamount = rsp.Fields(0)
        Else
        glamount = 0
        End If
        If glamount > 0 Then
        aamount = Amount - glamount
        Else
        aamount = Amount
        End If
        
        If aamount = 0 Then
       GoTo horola1
        End If
        If days <= 30 Then
        d1 = days
        GoTo horola
        End If
        If days > 30 And days <= 60 Then
        d2 = days
        GoTo horola
        End If
        If days > 60 And days <= 90 Then
        d3 = days
        GoTo horola
        End If
        If days > 90 And days <= 180 Then
        d4 = days
        GoTo horola
        End If
        If days > 180 Then
        d5 = days
        GoTo horola
        End If
horola:
        sql = ""
        sql = "INSERT INTO d_creditors_aging(dcode,Invoiceno,amount,[upto 30],[upto 60],[upto 90],[upto 180],[over 180]) VALUES('" & costcentre1 & "','" & invoiceno & "'," & Amount & "," & d1 & "," & d2 & "," & d3 & "," & d4 & "," & d5 & ")"
        oSaccoMaster.ExecuteThis (sql)
        d1 = 0
        d2 = 0
        d3 = 0
        d4 = 0
        d5 = 0
        
        Rst.MoveNext
        Wend
horola1:
rs.MoveNext
Wend

'd_aging.rpt
 reportname = "C_aging.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
MsgBox "Record successfully generated", vbInformation

End Sub

Private Sub optCash_Click()
optCheque_Click
If Optcash = True Then
        txtcontra = ""
        lblcontra = GetLedgerDesc(txtcontra)
        txtChequeno.Visible = True
        lblCheque.Visible = True
        lblVoucher.Visible = True
    End If
End Sub

Private Sub optCheque_Click()
If Optcheque = True Then
   
    txtChequeno.Visible = True
    lblVoucher.Visible = True
    lblCheque.Visible = True
End If
End Sub

Private Sub Picture21_Click()
frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            TxtOtherPAcc = SearchValue
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
    On Error GoTo SysError
    Dim Account As Acc_Details
    Editing = True
    Account = Get_Acc_Details(txtcontra, ErrorMessage)
    If Account.AccNo <> "" Then
        lblcontra = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblcontra = ""
    End If
    '//GET THE BALANCE AMOUNT
       TxtDRAmount = getGlCurrentBalance(Account.AccNo)

    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub

Private Sub TxtOtherPAcc_Change()
 On Error GoTo SysError
    Dim Account As Acc_Details
    Editing = True
    Account = Get_Acc_Details(TxtOtherPAcc, ErrorMessage)
    If Account.AccNo <> "" Then
        lblOtherPaymentAcc = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblOtherPaymentAcc = ""
    End If
    Exit Sub
SysError:
    MsgBox err.description, vbInformation, Me.Caption

End Sub







