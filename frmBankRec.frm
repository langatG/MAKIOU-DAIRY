VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBankRec 
   BackColor       =   &H00C0C000&
   Caption         =   "Bank Reconciliation"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBankRec.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Print Recon Report"
      Height          =   345
      Left            =   8640
      TabIndex        =   103
      Top             =   7440
      Width           =   1755
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   5940
      TabIndex        =   89
      Top             =   7455
      Width           =   1140
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print Details"
      Height          =   375
      Left            =   7200
      TabIndex        =   83
      Top             =   7455
      Width           =   1200
   End
   Begin VB.CommandButton cmdOffset 
      Caption         =   "Change To Offset"
      Height          =   375
      Left            =   4320
      TabIndex        =   82
      Top             =   7455
      Width           =   1620
   End
   Begin VB.CommandButton cmdTransferFunds 
      Caption         =   "Transfer Funds"
      Height          =   375
      Left            =   2930
      TabIndex        =   73
      Top             =   7455
      Width           =   1395
   End
   Begin MSComctlLib.ListView lvwAccounts 
      Height          =   270
      Left            =   1470
      TabIndex        =   72
      Top             =   585
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   476
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "AccNo"
         Object.Width           =   18
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "AccName"
         Object.Width           =   10583
      EndProperty
   End
   Begin VB.TextBox txtBankName 
      Height          =   300
      Left            =   1467
      TabIndex        =   71
      Top             =   300
      Width           =   2700
   End
   Begin VB.ComboBox cboBank 
      Height          =   315
      Left            =   105
      TabIndex        =   68
      Top             =   300
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3150
      Top             =   7410
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   375
      Left            =   120
      TabIndex        =   67
      Top             =   7455
      Width           =   1395
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5940
      Left            =   120
      TabIndex        =   12
      Top             =   1380
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   10478
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   12632064
      TabCaption(0)   =   "Cash Book Transactions"
      TabPicture(0)   =   "frmBankRec.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(1)=   "fraTransfer"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Bank Debits"
      TabPicture(1)   =   "frmBankRec.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label15"
      Tab(1).Control(1)=   "Label16"
      Tab(1).Control(2)=   "Label17"
      Tab(1).Control(3)=   "Label18"
      Tab(1).Control(4)=   "Label24"
      Tab(1).Control(5)=   "Label25"
      Tab(1).Control(6)=   "lvwDebits"
      Tab(1).Control(7)=   "txtCrAccNo"
      Tab(1).Control(8)=   "txtCrAccName"
      Tab(1).Control(9)=   "txtDrNarration"
      Tab(1).Control(10)=   "txtDrAmount"
      Tab(1).Control(11)=   "txtDrDocumentNo"
      Tab(1).Control(12)=   "dtpDrTransDate"
      Tab(1).Control(13)=   "cmdAddDebits"
      Tab(1).Control(14)=   "lvwCrAccounts"
      Tab(1).Control(15)=   "cmdDrPost"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "Bank Credits"
      TabPicture(2)   =   "frmBankRec.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label19"
      Tab(2).Control(1)=   "Label20"
      Tab(2).Control(2)=   "Label21"
      Tab(2).Control(3)=   "Label22"
      Tab(2).Control(4)=   "Label23"
      Tab(2).Control(5)=   "Label26"
      Tab(2).Control(6)=   "dtpCrTransDate"
      Tab(2).Control(7)=   "lvwCredits"
      Tab(2).Control(8)=   "txtDrAccNo"
      Tab(2).Control(9)=   "txtDrAccName"
      Tab(2).Control(10)=   "txtCrNarration"
      Tab(2).Control(11)=   "txtCrAmount"
      Tab(2).Control(12)=   "lvwDrAccounts"
      Tab(2).Control(13)=   "txtCrDocumentNo"
      Tab(2).Control(14)=   "cmdAddCredit"
      Tab(2).Control(15)=   "cmdCrPost"
      Tab(2).ControlCount=   16
      TabCaption(3)   =   "Reconciliation Report"
      TabPicture(3)   =   "frmBankRec.frx":0496
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label7"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label8"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label9"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label10"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label11"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label12"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label13"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label14"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Label6"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Label27"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Label37"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Label38"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "txtReceipts"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "txtPayments"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "txtCBBalance"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "txtUnpresentedChq"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "txtDeposits"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "txtBankCredits"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "txtBankDebits"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "txtBankBal"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "txtOpeningBal"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "txtDifference"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "cmdPostRecon"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "cmdPost"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "txtprevuncredited"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "txtprevunpresented"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).ControlCount=   26
      Begin VB.TextBox txtprevunpresented 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3945
         Locked          =   -1  'True
         TabIndex        =   100
         Top             =   3360
         Width           =   1950
      End
      Begin VB.TextBox txtprevuncredited 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3930
         Locked          =   -1  'True
         TabIndex        =   99
         Top             =   3750
         Width           =   1950
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "&Post"
         Height          =   360
         Left            =   9240
         TabIndex        =   86
         Top             =   4695
         Width           =   1305
      End
      Begin VB.Frame fraTransfer 
         BackColor       =   &H00C0C000&
         Caption         =   "Transfer Funds"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         Left            =   -74760
         TabIndex        =   74
         Top             =   2400
         Visible         =   0   'False
         Width           =   11130
         Begin VB.Frame Frame1 
            Height          =   1455
            Left            =   6360
            TabIndex        =   91
            Top             =   720
            Width           =   4095
            Begin VB.ComboBox cboType 
               Height          =   315
               ItemData        =   "frmBankRec.frx":04B2
               Left            =   120
               List            =   "frmBankRec.frx":04BC
               TabIndex        =   97
               Top             =   945
               Width           =   1335
            End
            Begin VB.TextBox txtAmt 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               TabIndex        =   93
               Top             =   360
               Width           =   1815
            End
            Begin VB.CommandButton cmdfTransfer 
               Caption         =   "&Funds Transfer"
               Height          =   390
               Left            =   2160
               TabIndex        =   92
               Top             =   840
               Width           =   1305
            End
            Begin MSComCtl2.DTPicker dtpTDate 
               Height          =   300
               Left            =   2040
               TabIndex        =   95
               Top             =   360
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   529
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
               Format          =   105906179
               CurrentDate     =   39406
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               Caption         =   "Type"
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
               Left            =   135
               TabIndex        =   98
               Top             =   720
               Width           =   405
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               Caption         =   "Transaction Date"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2040
               TabIndex        =   96
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               Caption         =   "Amount"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   135
               TabIndex        =   94
               Top             =   120
               Width           =   675
            End
         End
         Begin VB.CheckBox optTransfer 
            Caption         =   "Transfer Funds"
            Height          =   255
            Left            =   6360
            TabIndex        =   90
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtAccNo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   45
            TabIndex        =   79
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtAccName 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1530
            TabIndex        =   78
            Top             =   480
            Width           =   4305
         End
         Begin VB.CommandButton cmdTransfer 
            Caption         =   "Change Accounts"
            Height          =   390
            Left            =   480
            TabIndex        =   77
            Top             =   1680
            Width           =   1425
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   390
            Left            =   3120
            TabIndex        =   76
            Top             =   1680
            Width           =   1425
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   765
            Left            =   1560
            TabIndex        =   75
            Top             =   780
            Visible         =   0   'False
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   1349
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "AccNo"
               Object.Width           =   18
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "AccName"
               Object.Width           =   10583
            EndProperty
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Account No"
            Height          =   210
            Left            =   60
            TabIndex        =   81
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Account Name"
            Height          =   195
            Left            =   1545
            TabIndex        =   80
            Top             =   240
            Width           =   1035
         End
      End
      Begin VB.CommandButton cmdPostRecon 
         Caption         =   "Post Reconciliation"
         Height          =   435
         Left            =   8520
         TabIndex        =   66
         Top             =   6960
         Width           =   2010
      End
      Begin VB.CommandButton cmdCrPost 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   -67875
         TabIndex        =   65
         Top             =   4785
         Width           =   1380
      End
      Begin VB.CommandButton cmdDrPost 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   -67875
         TabIndex        =   64
         Top             =   4785
         Width           =   1380
      End
      Begin VB.TextBox txtDifference 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3945
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   5520
         Width           =   1950
      End
      Begin MSComctlLib.ListView lvwCrAccounts 
         Height          =   810
         Left            =   -73350
         TabIndex        =   61
         Top             =   4380
         Visible         =   0   'False
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   1429
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "AccountName"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "AccNo"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdAddDebits 
         Caption         =   "&Add"
         Height          =   375
         Left            =   -69420
         TabIndex        =   60
         Top             =   4785
         Width           =   1380
      End
      Begin VB.CommandButton cmdAddCredit 
         Caption         =   "&Add"
         Height          =   375
         Left            =   -69420
         TabIndex        =   59
         Top             =   4785
         Width           =   1380
      End
      Begin MSComCtl2.DTPicker dtpDrTransDate 
         Height          =   330
         Left            =   -74700
         TabIndex        =   55
         Top             =   4845
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   " dd-MM-yyyy"
         Format          =   147587075
         CurrentDate     =   39407
      End
      Begin VB.TextBox txtDrDocumentNo 
         Height          =   300
         Left            =   -65955
         TabIndex        =   53
         Top             =   4080
         Width           =   1590
      End
      Begin VB.TextBox txtCrDocumentNo 
         Height          =   300
         Left            =   -65955
         TabIndex        =   51
         Top             =   4095
         Width           =   1590
      End
      Begin MSComctlLib.ListView lvwDrAccounts 
         Height          =   810
         Left            =   -73335
         TabIndex        =   50
         Top             =   4395
         Visible         =   0   'False
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   1429
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "AccountName"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "AccNo"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtCrAmount 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   -67575
         TabIndex        =   48
         Top             =   4095
         Width           =   1605
      End
      Begin VB.TextBox txtCrNarration 
         Height          =   300
         Left            =   -70605
         TabIndex        =   46
         Top             =   4095
         Width           =   3045
      End
      Begin VB.TextBox txtDrAccName 
         Height          =   300
         Left            =   -73350
         TabIndex        =   44
         Top             =   4095
         Width           =   2760
      End
      Begin VB.TextBox txtDrAccNo 
         Height          =   300
         Left            =   -74700
         TabIndex        =   42
         Top             =   4095
         Width           =   1350
      End
      Begin VB.TextBox txtDrAmount 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   -67575
         TabIndex        =   40
         Top             =   4080
         Width           =   1605
      End
      Begin VB.TextBox txtDrNarration 
         Height          =   300
         Left            =   -70605
         TabIndex        =   38
         Top             =   4080
         Width           =   3045
      End
      Begin VB.TextBox txtCrAccName 
         Height          =   300
         Left            =   -73350
         TabIndex        =   36
         Top             =   4080
         Width           =   2760
      End
      Begin VB.TextBox txtCrAccNo 
         Height          =   300
         Left            =   -74700
         TabIndex        =   34
         Top             =   4080
         Width           =   1350
      End
      Begin VB.TextBox txtOpeningBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3945
         TabIndex        =   33
         Top             =   682
         Width           =   1950
      End
      Begin VB.TextBox txtBankBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3945
         TabIndex        =   31
         Top             =   5115
         Width           =   1950
      End
      Begin VB.TextBox txtBankDebits 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3945
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   4710
         Width           =   1950
      End
      Begin VB.TextBox txtBankCredits 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3945
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   4320
         Width           =   1950
      End
      Begin VB.TextBox txtDeposits 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3945
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2955
         Width           =   1950
      End
      Begin VB.TextBox txtUnpresentedChq 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2572
         Width           =   1950
      End
      Begin VB.TextBox txtCBBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3945
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1792
         Width           =   1950
      End
      Begin VB.TextBox txtPayments 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3945
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1417
         Width           =   1950
      End
      Begin VB.TextBox txtReceipts 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3945
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1057
         Width           =   1950
      End
      Begin MSComctlLib.ListView lvwCredits 
         Height          =   3060
         Left            =   -74715
         TabIndex        =   14
         Top             =   765
         Width           =   10110
         _ExtentX        =   17833
         _ExtentY        =   5398
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TransDate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Narration"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "AccNo to Debit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Document No"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4800
         Left            =   -75000
         TabIndex        =   13
         Top             =   390
         Width           =   12000
         _ExtentX        =   21167
         _ExtentY        =   8467
         View            =   3
         LabelEdit       =   1
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TransDate"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Document No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "TransDescription"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Debit"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Credit"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Balance"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Statement Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "ID"
            Object.Width           =   18
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Chequeno"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvwDebits 
         Height          =   3060
         Left            =   -74715
         TabIndex        =   15
         Top             =   765
         Width           =   10110
         _ExtentX        =   17833
         _ExtentY        =   5398
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TransDate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Narration"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "AccNo to Credit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Document No"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpCrTransDate 
         Height          =   330
         Left            =   -74700
         TabIndex        =   57
         Top             =   4845
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   " dd-MM-yyyy"
         Format          =   147587075
         CurrentDate     =   39407
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "LESS:Previous Unpresented Cheques ...."
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
         Left            =   600
         TabIndex        =   102
         Top             =   3420
         Width           =   3330
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "ADD:Previous Uncredited Cheques ......"
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
         Left            =   600
         TabIndex        =   101
         Top             =   3810
         Width           =   3165
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Difference ......................................................"
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
         Left            =   615
         TabIndex        =   63
         Top             =   5580
         Width           =   3330
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Trans Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74580
         TabIndex        =   58
         Top             =   4620
         Width           =   930
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Trans Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74580
         TabIndex        =   56
         Top             =   4620
         Width           =   930
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Document No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -65940
         TabIndex        =   54
         Top             =   3855
         Width           =   1125
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Document No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -65940
         TabIndex        =   52
         Top             =   3870
         Width           =   1125
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -66675
         TabIndex        =   49
         Top             =   3870
         Width           =   675
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Narration"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -70545
         TabIndex        =   47
         Top             =   3870
         Width           =   795
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Acc Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -72750
         TabIndex        =   45
         Top             =   3870
         Width           =   825
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Acc No to Credit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74685
         TabIndex        =   43
         Top             =   3870
         Width           =   1335
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -66675
         TabIndex        =   41
         Top             =   3855
         Width           =   675
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Narration"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -70545
         TabIndex        =   39
         Top             =   3855
         Width           =   795
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Acc Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -72750
         TabIndex        =   37
         Top             =   3855
         Width           =   825
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Acc No to Credit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74685
         TabIndex        =   35
         Top             =   3855
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Opening Balance .........................................."
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
         Left            =   615
         TabIndex        =   32
         Top             =   735
         Width           =   3300
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Balance as per Bank Statement ................"
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
         Left            =   615
         TabIndex        =   23
         Top             =   5175
         Width           =   3315
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "LESS: Direct Bank Debits ............................"
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
         Left            =   615
         TabIndex        =   22
         Top             =   4770
         Width           =   3300
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "ADD: Direct Bank Credits ............................"
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
         Left            =   615
         TabIndex        =   21
         Top             =   4380
         Width           =   3315
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "LESS: Uncredited Cheques ........................"
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
         Left            =   615
         TabIndex        =   20
         Top             =   3015
         Width           =   3315
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "ADD: Unpresented Cheques ......................"
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
         Left            =   615
         TabIndex        =   19
         Top             =   2625
         Width           =   3330
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Cash Book Balance ......................................"
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
         Left            =   615
         TabIndex        =   18
         Top             =   1845
         Width           =   3315
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Payments ......................................................"
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
         Left            =   615
         TabIndex        =   17
         Top             =   1470
         Width           =   3300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Receipts ........................................................"
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
         Left            =   615
         TabIndex        =   16
         Top             =   1110
         Width           =   3285
      End
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   1525
      TabIndex        =   11
      Top             =   7455
      Width           =   1395
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   405
      Left            =   10815
      TabIndex        =   6
      Top             =   7410
      Width           =   1335
   End
   Begin VB.TextBox txtbankbalance 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6276
      TabIndex        =   5
      Top             =   540
      Width           =   1575
   End
   Begin VB.TextBox txtOpBalance 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4434
      TabIndex        =   3
      Top             =   300
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker dtpReconciliation 
      Height          =   300
      Left            =   4980
      TabIndex        =   1
      Top             =   900
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   529
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
      Format          =   182321155
      CurrentDate     =   39406
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   300
      Left            =   8130
      TabIndex        =   7
      Top             =   300
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   529
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
      Format          =   182321155
      CurrentDate     =   39406
   End
   Begin MSComCtl2.DTPicker dtpFinishDate 
      Height          =   300
      Left            =   9660
      TabIndex        =   9
      Top             =   300
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
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
      Format          =   182321155
      CurrentDate     =   39406
   End
   Begin MSComCtl2.DTPicker dtpStatement 
      Height          =   300
      Left            =   7860
      TabIndex        =   87
      Top             =   900
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   529
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
      Format          =   182321155
      CurrentDate     =   39406
   End
   Begin VB.TextBox txtStatementNo 
      Height          =   300
      Left            =   1320
      TabIndex        =   84
      Top             =   900
      Width           =   1830
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "Statement Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6420
      TabIndex        =   88
      Top             =   945
      Width           =   1410
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "Statement No:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   85
      Top             =   945
      Width           =   1215
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "Bank Code"
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
      Left            =   120
      TabIndex        =   70
      Top             =   75
      Width           =   885
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "Bank Name"
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
      Left            =   1290
      TabIndex        =   69
      Top             =   90
      Width           =   915
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Finish Date"
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
      Left            =   9735
      TabIndex        =   10
      Top             =   90
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Start  Date"
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
      Left            =   8235
      TabIndex        =   8
      Top             =   90
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Closing Balance(Bank )"
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
      Left            =   6150
      TabIndex        =   4
      Top             =   210
      Width           =   1875
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Opening Balance (CBook)"
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
      Left            =   3780
      TabIndex        =   2
      Top             =   90
      Width           =   2070
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Reconcilaition Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   0
      Top             =   945
      Width           =   1680
   End
End
Attribute VB_Name = "frmBankRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim li As ListItem

Private Sub cboBank_Change()
    On Error GoTo sysError
     Get_GL_AccDetails (cboBank)
    If GlAccName <> "" Then
        txtBankName = GlAccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        txtBankName = ""
    End If
    Exit Sub
sysError:
    MsgBox err.description
End Sub

Private Sub cboBank_Click()
cboBank_Change
End Sub

Private Sub cmdAddCredit_Click()
    On Error GoTo sysError
    If Trim$(txtDrAccNo) = "" Then
        MsgBox "Please select an account to Debit", vbInformation, Me.Caption
        txtDrAccName.SetFocus
        Exit Sub
    End If
    If Trim$(txtDrAccName) = "" Then
        MsgBox "Please select an account to Debit", vbInformation, Me.Caption
        txtDrAccName.SetFocus
        Exit Sub
    End If
    If Trim$(txtCrNarration) = "" Then
        MsgBox "Please enter a description for the Transaction", vbInformation, Me.Caption
        txtCrNarration.SetFocus
        Exit Sub
    End If
    If Trim$(txtCrAmount) = "" Then
        MsgBox "Please enter an Amount.", vbInformation, Me.Caption
        txtCrAmount.SetFocus
        Exit Sub
    End If
    If Trim$(txtCrDocumentNo) = "" Then
        MsgBox "Please enter a Document No", vbInformation, Me.Caption
        txtCrDocumentNo.SetFocus
        Exit Sub
    End If
    If dtpCrTransDate < dtpStartDate Or dtpCrTransDate > dtpFinishDate Then
       MsgBox "Bank Credits Transdate should be Between :" & dtpStartDate & "And :" & dtpFinishDate
     Exit Sub
    End If
    Set li = lvwCredits.ListItems.Add(, , dtpCrTransDate)
    li.SubItems(1) = txtCrNarration
    li.SubItems(2) = txtDrAccNo
    li.SubItems(3) = Format(txtCrAmount, Cfmt)
    li.SubItems(4) = txtCrDocumentNo
    txtCrNarration = ""
    txtCrAmount = ""
    txtCrDocumentNo = ""
    txtDrAccName = ""
    txtDrAccNo = ""
    txtDrAccNo.SetFocus
    'SendKeys "{Home}+{End}"
    'Load_Statement
    Calculate_Summaries
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdAddDebits_Click()
    On Error GoTo sysError
    On Error GoTo sysError
    If Trim$(txtCrAccNo) = "" Then
        MsgBox "Please select an account to Credit", vbInformation, Me.Caption
        txtCrAccName.SetFocus
        Exit Sub
    End If
    If Trim$(txtCrAccName) = "" Then
        MsgBox "Please select an account to Credit", vbInformation, Me.Caption
        txtDrAccName.SetFocus
        Exit Sub
    End If
    If Trim$(txtDrNarration) = "" Then
        MsgBox "Please enter a description for the Transaction", vbInformation, Me.Caption
        txtDrNarration.SetFocus
        Exit Sub
    End If
    If Trim$(txtDrAmount) = "" Then
        MsgBox "Please enter an Amount.", vbInformation, Me.Caption
        txtDrAmount.SetFocus
        Exit Sub
    End If
    If Trim$(txtDrDocumentNo) = "" Then
        MsgBox "Please enter a Document No", vbInformation, Me.Caption
        txtDrDocumentNo.SetFocus
        Exit Sub
    End If
    If dtpDrTransDate < dtpStartDate Or dtpDrTransDate > dtpFinishDate Then
       MsgBox "Bank Debits Transdate should be Between :" & dtpStartDate & "And :" & dtpFinishDate
     Exit Sub
    End If
    Set li = lvwDebits.ListItems.Add(, , dtpDrTransDate)
    li.SubItems(1) = txtDrNarration
    li.SubItems(2) = txtCrAccNo
    li.SubItems(3) = Format(txtDrAmount, Cfmt)
    li.SubItems(4) = txtDrDocumentNo
    txtDrNarration = ""
    txtDrAmount = ""
    txtDrDocumentNo = ""
    txtCrAccName = ""
    txtCrAccNo = ""
    txtCrAccNo.SetFocus
    'SendKeys "{Home}+{End}"
    'Load_Statement
    Calculate_Summaries
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdCancel_Click()
    fraTransfer.Visible = False
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdCrPost_Click()
    On Error GoTo sysError
    If MsgBox("Do you want to remove the selected Item?", vbQuestion + _
    vbYesNo, "Remove Item") = vbNo Then
        Exit Sub
    End If
    lvwCredits.ListItems.Remove lvwCredits.SelectedItem.Index
    Calculate_Summaries
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdDrPost_Click()
    On Error GoTo sysError
    If MsgBox("Do you want to remove the selected Item?", vbQuestion + _
    vbYesNo, "Remove Item") = vbNo Then
        Exit Sub
    End If
    lvwDebits.ListItems.Remove lvwDebits.SelectedItem.Index
    Calculate_Summaries
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdexport_Click()
    On Error GoTo SsyError
    Dim MyFso As New FileSystemObject, strData As String, MFile As TextStream, _
    FileName As String, I As Long, li As ListItem
    If ListView1.ListItems.Count > 0 Then
        With CommonDialog1
            .Filter = "Comma Seperated Values|*.csv"
            .ShowSave
            If .FileName <> "" Then
                FileName = .FileName
            End If
            .FileName = ""
        End With
        Set MFile = MyFso.OpenTextFile(FileName, ForWriting, True)
        strData = "TransDate,MemberNo,Names,Receipts,Payments,Document No"
        MFile.WriteLine strData
        strData = ""
        For I = 1 To ListView1.ListItems.Count
            Set li = ListView1.ListItems(I)
            strData = li & "," & li.SubItems(6) & "," & li.SubItems(1) & "," & CDbl(li.SubItems(2)) _
            & "," & CDbl(li.SubItems(3)) & "," & li.SubItems(5)
            MFile.WriteLine strData
            strData = ""
        Next I
    Else
        MsgBox "There are no records to be exported", vbInformation, Me.Caption
    End If
    Exit Sub
SsyError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdfTransfer_Click()
    Dim I As Long, AccNo As String, ContraAccNo As String, _
    tVNo As String, desc As String, ttype As String
    If txtBankName = "" Or txtAccName = "" Then
        MsgBox "Ensure both accounts are captured before transfer.", vbInformation
        Exit Sub
    End If
    
    If txtAmt = "0" Or txtAmt = "" Then
        MsgBox "Please enter a valid amount.", vbInformation
        Exit Sub
    End If
    
    If cboType = "" Then
        MsgBox "Please select the transaction type.", vbInformation
        Exit Sub
    End If
    
    desc = "Account Transfers"
    For I = 1 To 2
        tVNo = IIf(I = 1, "From " & txtAccNo, "From " & cboBank)
        AccNo = IIf(I = 1, cboBank, txtAccNo)
        ttype = IIf(I = 1, cboType, IIf(cboType = "DR", "CR", "DR"))
        ContraAccNo = IIf(I = 1, txtAccNo, cboBank)
'        If Not Save_To_Customer_Balance(txtAccNo, txtAccNo, txtAccNo, txtAccName, txtAmt, 0, _
'        accno, desc, dtpTDate, 0, "", month(dtpTDate), 0, 0, ttype, 0, tVNo, current_user.UserId, _
'        "", ContraAccNo, dtpTDate, 0, 0, "", 0, transactionNo, 0, ErrorMessage) Then
'            If ErrorMessage <> "" Then
'                MsgBox ErrorMessage, vbInformation
'                ErrorMessage = ""
'                Exit Sub
'            End If
'        End If
    Next I
    MsgBox "Transfer of funds was successfull.", vbInformation
End Sub

Private Sub cmdload_Click()
    'On Error GoTo SysError
    Dim rsRecon As New Recordset, BankBal As Double, bCredits As Double, bDebits As Double, _
    RsDesc As New Recordset, PMonth As Date
    If Trim(cboBank) = "" Then
        MsgBox "Please indicate the Bank Account to Reconcile", vbInformation, Me.Caption
        txtBankName.SetFocus
       ' SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If dtpStartDate > dtpFinishDate Then
        MsgBox "The StartDate should be Earlier than the FinishDate", vbInformation, Me.Caption
        Exit Sub
    End If
    'Get_GL_AccDetails cboBank
    If dtpStartDate < EarliestTransDate Then
        MsgBox "The StartDate should  not be earlier than " & Format(EarliestTransDate, _
        "dd-MM-yyyy"), vbInformation, Me.Caption
        dtpStartDate.SetFocus
        Exit Sub
    End If
    Set rsRecon = Get_Records("Select OpeningBal as Amount From GLSETUP Where AccNo='" & _
    cboBank & "'", ErrorMessage)
    With rsRecon
        If .State = adStateOpen Then
            If Not .EOF Then
                BankBal = BankBal + IIf(IsNull(!Amount), 0, !Amount)
            End If
        End If
    End With
    txtOpeningBal = Format(BankBal, Cfmt)
    txtOpBalance = txtOpeningBal
    Startdate = DateSerial(year(dtpStartDate), month(dtpStartDate) - 1, 1)
    FinishDate = DateSerial(year(dtpStartDate), month(dtpStartDate), 1 - 1)
    
'    Set rsRecon = Get_Records("Set DateFormat DMY Select * From CUSTOMERBALANCE where AccNo='" & _
'    cboBank & "' and Transdate>='01/01/2016' AND Transdate <'" & dtpStartDate & "' and reconciled=0  Order By TransDate,CustomerBalanceID", ErrorMessage)
'
    Set rsRecon = oSaccoMaster.GetRecordset("SET DATEFORMAT DMY EXEC getBankTransactions '" & cboBank & "','01/01/2016','" & dtpFinishDate & "'")
    
    ListView1.ListItems.Clear
    With rsRecon
        While Not .EOF
            DoEvents
            Set li = ListView1.ListItems.Add(, , !transdate)
            'li.Checked = True
            li.SubItems(1) = IIf(IsNull(!DocumentNo), "", !DocumentNo)
            li.SubItems(2) = IIf(IsNull(!TransDescription), "", !TransDescription)
            li.SubItems(3) = Format(IIf(!transtype <> "DR", 0, !Amount), Cfmt)
            bDebits = bDebits + CDbl(li.SubItems(3))
            li.SubItems(4) = Format(IIf(!transtype <> "CR", 0, !Amount), Cfmt)
            bCredits = bCredits + CDbl(li.SubItems(4))
            BankBal = BankBal + CDbl(li.SubItems(3)) - CDbl(li.SubItems(4))
            li.SubItems(5) = Format(BankBal, Cfmt)
            li.SubItems(6) = li
            li.SubItems(7) = !id
            li.SubItems(8) = IIf(IsNull(!chequeno), "", !chequeno)
            .MoveNext
        Wend
    End With

    txtReceipts = Format(bDebits, Cfmt)
    txtPayments = Format(bCredits, Cfmt)
    txtCBBalance = Format(BankBal, Cfmt)
    'Load_Statement
    Calculate_Summaries
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Load_Statement()
    On Error GoTo SysEror
    Dim UnprecCheques As Double, BankDebits As Double, BankCredits As Double, _
    UnDebs As Double
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = False Then
            If CDbl(ListView1.ListItems(I).SubItems(4)) <> 0 Then
                UnprecCheques = UnprecCheques + CDbl(ListView1.ListItems(I).SubItems(4))
            End If
            If CDbl(ListView1.ListItems(I).SubItems(3)) <> 0 Then
                UnDebs = UnDebs + CDbl(ListView1.ListItems(I).SubItems(3))
            End If
        Else
        End If
    Next I
    For I = 1 To lvwDebits.ListItems.Count
        If CDbl(lvwDebits.ListItems(I).SubItems(3)) <> 0 Then
            BankDebits = BankDebits + CDbl(lvwDebits.ListItems(I).SubItems(3))
        End If
    Next
    For I = 1 To lvwCredits.ListItems.Count
        If CDbl(lvwCredits.ListItems(I).SubItems(3)) <> 0 Then
            BankCredits = BankCredits + CDbl(lvwCredits.ListItems(I).SubItems(3))
        End If
    Next
    txtUnpresentedChq = Format(UnprecCheques, Cfmt)
    txtBankCredits = Format(BankCredits, Cfmt)
    txtBankDebits = Format(BankDebits, Cfmt)
    txtDeposits = Format(UnDebs, Cfmt)
    Exit Sub
SysEror:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdrefresh_Click()
    On Error GoTo sysError
    Dim UnprecCheques As Double
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked = False Then
            If CDbl(ListView1.ListItems(I).SubItems(2)) <> 0 Then
                UnprecCheques = UnprecCheques + CDbl(ListView1.ListItems(I).SubItems(2))
            End If
        Else
        End If
    Next I
    txtUnpresentedChq = Format(UnprecCheques, Cfmt)
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdOffset_Click()
    On Error GoTo sysError
    Dim FromAcc As String, ToAcc As String, DocumentNo As String, TransT As String, _
    transdate As Date, mCredit As String
    Dim rsTransfer As New Recordset
    If MsgBox("Do you want to Change the selected Transaction to an Offset?", _
    vbQuestion + vbYesNo, Me.Caption) = vbNo Then
        Exit Sub
    End If
    If MsgBox("This process is not reversible. Do you want to continue?", _
    vbExclamation + vbYesNo, Me.Caption) = vbNo Then
        Exit Sub
    End If
    If ListView1.ListItems.Count > 0 Then
        Set li = ListView1.SelectedItem
        transdate = li
        DocumentNo = CStr(li.SubItems(5))
        TransT = CStr(li.SubItems(6))
        Select Case CDbl(li.SubItems(3)) 'PAYMENTS
            Case 0 'XXXXXXXXXXX RECEIPTS XXXXXXXXX
            mCredit = "DR"
            Case Else 'XXXXXXXX PAYMENTS XXXXXXXXX
            mCredit = "CR"
        End Select
        Set rsTransfer = oSaccoMaster.GetRecordset("Set DateFormat DMY Update CUSTOMERBALANCE" _
        & " Set IDNo='2' where AccNo='" & cboBank & "' and VNo='" & TransT & _
        "' and ChequeNo='" & DocumentNo & "' and TransDate='" & transdate & "' and TransType='" _
        & mCredit & "'")
        cmdload_Click
    End If
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Calculate_Summaries()
    Dim CashBookOpBal As Double, Credits As Double, Debits As Double, _
    StateOpenBal As Double, UnpresentedCheques As Double, UncreditedCheques As Double, _
    BankDebits As Double, BankCredits As Double, I As Long, Unpre As Double, Uncre As Double
    Dim PUnpre As Double, PUncre As Double
    Dim LastMonth As Date
    On Error GoTo sysError
    'If Editing Then
        'Start with Cashbook Entries
        Unpre = 0
        Uncre = 0
        PUncre = 0
        PUnpre = 0
        LastMonth = Format(dtpStartDate, "dd/mm/yyyy")
        If Trim$(txtOpeningBal) = "" Then txtOpeningBal = "0.00"
        If Trim$(txtOpBalance) = "" Then txtOpBalance = "0.00"
        CashBookOpBal = CDbl(txtOpeningBal)
        StateOpenBal = CDbl(txtOpBalance)
        For I = 1 To ListView1.ListItems.Count
            Set li = ListView1.ListItems(I)
            Debits = Debits + CDbl(li.SubItems(3))
            Credits = Credits + CDbl(li.SubItems(4))
            If LastMonth > li Then
'                UnpresentedCheques = UnpresentedCheques + CDbl(li.SubItems(4))
'                UncreditedCheques = UncreditedCheques + CDbl(li.SubItems(3))
                Unpre = Unpre + CDbl(li.SubItems(4))
                Uncre = Uncre + CDbl(li.SubItems(3))
                If li.Checked = True Then
                PUnpre = PUnpre + CDbl(li.SubItems(4))
                PUncre = PUncre + CDbl(li.SubItems(3))
                End If
            End If
            If ListView1.ListItems(I).Checked = False Then
                UnpresentedCheques = UnpresentedCheques + CDbl(li.SubItems(4))
                UncreditedCheques = UncreditedCheques + CDbl(li.SubItems(3))
            End If
        Next I
        'Bank Debits
        For I = 1 To lvwCredits.ListItems.Count
            Set li = lvwCredits.ListItems(I)
            BankCredits = BankCredits + CDbl(li.SubItems(3))
        Next I
        'Bank Credits
        For I = 1 To lvwDebits.ListItems.Count
            Set li = lvwDebits.ListItems(I)
            BankDebits = BankDebits + CDbl(li.SubItems(3))
        Next I
    'End If
    
    txtReceipts = Format(Debits - PUncre, Cfmt)
    txtPayments = Format(Credits - PUnpre, Cfmt)
    txtprevuncredited = PUncre
    txtprevunpresented = PUnpre
    txtCBBalance = Format(CashBookOpBal - Credits + Debits, Cfmt)
    txtUnpresentedChq = Format(UnpresentedCheques, Cfmt)
    txtDeposits = Format(UncreditedCheques, Cfmt)
    txtBankBal = Format(StateOpenBal - Credits + Debits + UnpresentedCheques - UncreditedCheques + BankCredits - BankDebits - PUnpre + PUncre, Cfmt)
    txtBankCredits = Format(BankCredits, Cfmt)
    txtBankDebits = Format(BankDebits, Cfmt)
    UncreditedCheques = 0
    UnpresentedCheques = 0
    BankCredits = 0
    BankDebits = 0
    Credits = 0
    Debits = 0
    Editing = False
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdPost_Click()
    On Error GoTo sysError
    Dim StatementDate As Date, ReconcileDate As Date, transdate As Date, _
    Amount As Double, transtype As String, DocumentNo As String, TransDesc As _
    String, InStatement As String, balance As Double, AccNo As String
    
    balance = IIf(Trim(txtOpBalance) <> "", 0, CDbl(txtOpBalance))
    If Trim$(txtStatementNo) = "" Then
        MsgBox "Please enter the statement No.", vbInformation, Me.Caption
        txtStatementNo.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    If ListView1.ListItems.Count = 0 Then
        If lvwCrAccounts.ListItems.Count = 0 Then
            If lvwDrAccounts.ListItems.Count = 0 Then
                MsgBox "There are no entries to be posted", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
    End If
    StatementDate = dtpStartDate
    transdate = StatementDate
    TransDesc = "Opening Balance"
    transtype = "DR"
    Amount = CDbl(txtOpBalance)
    balance = CDbl(txtOpBalance)
    ReconcileDate = dtpReconciliation
    InStatement = "Yes"
    If Not SAVE_BANKREC(cboBank, StatementDate, transdate, TransDesc, transtype, _
    txtStatementNo, Amount, balance, ReconcileDate, "Opening Balance", InStatement, _
    current_user.UserID, ErrorMessage) Then
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
    End If
    For I = 1 To ListView1.ListItems.Count
        DoEvents
        Set li = ListView1.ListItems(I)
        If li.Checked Then
            InStatement = "Yes"
        Else
            InStatement = "No"
        End If
        StatementDate = li.SubItems(6)
        ReconcileDate = dtpReconciliation
        transdate = li
        If Val(li.SubItems(3)) > 0 Then
            transtype = "DR"
            Amount = CDbl(li.SubItems(3))
        End If
        If Val(li.SubItems(4)) > 0 Then
            transtype = "CR"
            Amount = CDbl(li.SubItems(4))
        End If
        DocumentNo = li.SubItems(1)
        TransDesc = li.SubItems(2)
        If Not SAVE_BANKREC(cboBank, StatementDate, transdate, TransDesc, transtype, _
        txtStatementNo, Amount, balance, dtpReconciliation, "", InStatement, _
        current_user.UserID, ErrorMessage) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        If InStatement = "Yes" Then
            If Not Execute_Command("Update GlTransactions Set Recon=1 where " _
            & "ID='" & li.SubItems(7) & "'", ErrorMessage) Then
                If ErrorMessage <> "" Then
                    MsgBox ErrorMessage, vbInformation, Me.Caption
                    ErrorMessage = ""
                End If
            End If
        End If
    Next I
    For I = 1 To lvwDebits.ListItems.Count
        DoEvents
        Set li = lvwDebits.ListItems(I)
        transdate = li
        TransDesc = li.SubItems(1)
        InStatement = "Yes"
        transtype = "DR"
        Amount = CDbl(li.SubItems(3))
        DocumentNo = li.SubItems(4)
        StatementDate = transdate
        ReconcileDate = dtpReconciliation
        
        'GetTransactionNo
        NewTransaction Amount, transdate, "Bank Debit"

        
        If Not SAVE_BANKREC(cboBank, StatementDate, transdate, TransDesc, transtype, _
        txtStatementNo, Amount, balance, ReconcileDate, "Bank Debit", InStatement, _
        current_user.UserID, ErrorMessage) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        
        AccNo = li.SubItems(2)
        
        Get_GL_AccDetails (AccNo)
        If GlAccName <> "" Then
        
            If Not Save_GLTRANSACTION(transdate, Amount, AccNo, cboBank, DocumentNo, AccNo, current_user.UserID, "", TransDesc, 1, 1, DocumentNo, transactionNo, "fosa") Then
             GoTo sysError
            End If
           
            If Not Execute_Command("set dateformat dmy Update GlTransactions Set Recon=1 where " _
            & "Transdate='" & transdate & "' and Draccno='" & AccNo & "' and Craccno='" & cboBank & "' And Amount=" & Amount & " and DocumentNo='" & DocumentNo & "' and Transdescript='" & TransDesc & "'  and Auditid='" & current_user.UserID & "'", ErrorMessage) Then
                If ErrorMessage <> "" Then
                    MsgBox ErrorMessage, vbInformation, Me.Caption
                    ErrorMessage = ""
                End If
            End If
            
        End If
    Next I
    For I = 1 To lvwCredits.ListItems.Count
        DoEvents
        Set li = lvwCredits.ListItems(I)
        transdate = li
        InStatement = "Yes"
        TransDesc = li.SubItems(1)
        transtype = "CR"
        Amount = CDbl(li.SubItems(3))
        DocumentNo = li.SubItems(4)
        StatementDate = transdate
        ReconcileDate = dtpReconciliation
        
        'GetTransactionNo
        NewTransaction Amount, transdate, "Bank Credit"

        If Not SAVE_BANKREC(cboBank, StatementDate, transdate, TransDesc, transtype, _
        txtStatementNo, Amount, balance, ReconcileDate, "Bank Credit", InStatement, _
        current_user.UserID, ErrorMessage) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        AccNo = ""
        AccNo = li.SubItems(2)
        Get_GL_AccDetails (AccNo)
        If GlAccName <> "" Then
        
            If Not Save_GLTRANSACTION(transdate, Amount, cboBank, AccNo, DocumentNo, AccNo, current_user.UserID, "", TransDesc, 1, 1, DocumentNo, transactionNo, "fosa") Then
              GoTo sysError
            End If
            
            If Not Execute_Command("set dateformat dmy Update GlTransactions Set Recon=1 where " _
            & "Transdate='" & transdate & "' and Draccno='" & cboBank & "' and Craccno='" & AccNo & "' And Amount=" & Amount & " and DocumentNo='" & DocumentNo & "' and Transdescript='" & TransDesc & "'  and Auditid='" & current_user.UserID & "'", ErrorMessage) Then
                If ErrorMessage <> "" Then
                    MsgBox ErrorMessage, vbInformation, Me.Caption
                    ErrorMessage = ""
                End If
            End If
        End If
    Next I
    
    sql = "set dateformat dmy if exists (select * from bankrecon WHERE [AccNo] = '" & cboBank & "'  and mmonth='" & month(dtpStartDate) & "'  and Yyear='" & year(dtpStartDate) & "') " _
    & " update bankrecon SET  [OpeningBal]   = " & CDbl(txtOpeningBal) & ",[Unpresented]   = " & CDbl(txtUnpresentedChq) & ",[UnCredited]  = " & CDbl(txtDeposits) & ",[PreviousUncredited]   = " & CDbl(txtprevuncredited) & ",[PreviousUnpresented]   = " & CDbl(txtprevunpresented) & " ," _
    & " [DirectCredits]= " & CDbl(txtBankCredits) & ",[DirectDebits]  = " & CDbl(txtBankDebits) & ",[StatementBal]= " & CDbl(txtBankBal) & " ,[Receipts]= " & CDbl(txtReceipts) & ",[Payments]= " & CDbl(txtPayments) & " ,[CashBookBal]= " & CDbl(txtCBBalance) & " WHERE [AccNo] = '" & cboBank & "'  and mmonth='" & month(dtpStartDate) & "'  and Yyear='" & year(dtpStartDate) & "' else " _
    & "  INSERT INTO BankRecon  ( [AccNo], [OpeningBalDate], [OpeningBal],[Unpresented]," _
    & "[UnCredited],[DirectCredits],[DirectDebits],[StatementBal],[ReconDate],[Receipts],[Payments],[CashBookBal],[PreviousUncredited],[PreviousUnpresented])" _
    & " VALUES ( '" & cboBank & "','" & dtpStartDate & "'," & CDbl(txtOpeningBal) & ", " & CDbl(txtUnpresentedChq) & "," _
    & " " & CDbl(txtDeposits) & ", " & CDbl(txtBankCredits) & ", " & CDbl(txtBankDebits) & ", " & CDbl(txtBankBal) & ",'" & dtpFinishDate & "'," & CDbl(txtReceipts) & "," & CDbl(txtPayments) & "," & CDbl(txtCBBalance) & "," & CDbl(txtprevuncredited) & "," & CDbl(txtprevunpresented) & ")"

    oSaccoMaster.ExecuteThis (sql)
    
    MsgBox "Reconciliation Posted Successfully", vbInformation, Me.Caption
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdPostRecon_Click()
'    Dim rsRecon As New Recordset
'    On Error GoTo SysError
'    'XXXXXXXXXXX Update CashBook Entries XXXXXXXXXXXXXXXXXXXX'
'    For I = 1 To ListView1.ListItems.Count
'        mTransDate = ListView1.ListItems(I)
'        If ListView1.ListItems(I).Checked = True Then
'            Set rsRecon = OSACCOMASTER.GetRecordset("Set DateFormat DMY Update " _
'            & "CUSTOMERBALANCE Set Reconcile=1 where AccNo='A001' and TransDate='" _
'            & mTransDate & "' and VNo='" & ListView1.ListItems(I).SubItems(1) _
'            & "' and ChequeNo='" & ListView1.ListItems(I).SubItems(5) & "'")
'        End If
'    Next
'    Exit Sub
'SysError:
'    MsgBox Err.Description, vbInformation, Me.Caption
End Sub

Private Sub Generate_Reconciliatin_Report(CashBook_Balance As Double)
    
End Sub

Private Sub PrintRec()
    On Error GoTo sysError
    Dim RepPath As String, AccNo As String, credit As Double, debit As Double, balance As Double, _
    TransDescription As String
    
     If Trim(cboBank) = "" Then
        MsgBox "Please indicate the Bank Account to Reconcile", vbInformation, Me.Caption
        txtBankName.SetFocus
       ' SendKeys "{Home}+{End}"
        Exit Sub
    End If
    
    If Not Execute_Command("Exec Delete_Statement", ErrorMessage) Then
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
    End If

    
    Startdate = Format(dtpStartDate, "dd/mm/yyyy")
    FinishDate = Format(dtpFinishDate, "dd/mm/yyyy")
    AccNo = cboBank
    
    If ListView1.ListItems.Count > 0 Then
        For I = 1 To ListView1.ListItems.Count
            Set li = ListView1.ListItems(I)
            mTransDate = li
            If Startdate <= mTransDate Then
                If FinishDate >= mTransDate Then
                 If I = 1 Then
                    VoucherNo = "OpenBal"
                    credit = 0
                    debit = 0
                    balance = CDbl(txtOpeningBal)
                    TransDescription = "OpenBal"
                    If Not SAVE_STATEMENT(AccNo, VoucherNo, credit, debit, balance, _
                    TransDescription, Startdate, FinishDate, mTransDate, txtBankName, ErrorMessage) Then
                        If ErrorMessage <> "" Then
                            MsgBox ErrorMessage, vbInformation, Me.Caption
                            ErrorMessage = ""
                        End If
                    End If
                  End If
                    VoucherNo = li.SubItems(1)
                    credit = li.SubItems(4)
                    debit = li.SubItems(3)
                    balance = balance + IIf(li.SubItems(3) = 0, -li.SubItems(4), li.SubItems(3))
                    TransDescription = li.SubItems(2)
                    If Not SAVE_STATEMENT(AccNo, VoucherNo, credit, debit, balance, _
                    TransDescription, Startdate, FinishDate, mTransDate, txtBankName, ErrorMessage) Then
                        If ErrorMessage <> "" Then
                            MsgBox ErrorMessage, vbInformation, Me.Caption
                            ErrorMessage = ""
                        End If
                    End If
                End If
            End If
        Next I
     End If
     
     If lvwDebits.ListItems.Count > 0 Then
            For I = 1 To lvwDebits.ListItems.Count
            Set li = lvwDebits.ListItems(I)
            mTransDate = li
            If Startdate <= mTransDate Then
                If FinishDate >= mTransDate Then
                    VoucherNo = li.SubItems(4)
                    credit = 0
                    debit = li.SubItems(3)
                    balance = balance + IIf(li.SubItems(3) = "", 0, li.SubItems(3))
                    TransDescription = li.SubItems(2)
                    If Not SAVE_STATEMENT(AccNo, VoucherNo, credit, debit, balance, _
                    TransDescription, Startdate, FinishDate, mTransDate, txtBankName, ErrorMessage) Then
                        If ErrorMessage <> "" Then
                            MsgBox ErrorMessage, vbInformation, Me.Caption
                            ErrorMessage = ""
                        End If
                    End If
                End If
            End If
        Next I
    End If
    
    If lvwCredits.ListItems.Count > 0 Then
        For I = 1 To lvwCredits.ListItems.Count
        Set li = lvwCredits.ListItems(I)
        mTransDate = li
       
        If Startdate <= mTransDate Then
            If FinishDate >= mTransDate Then
                VoucherNo = li.SubItems(4)
                credit = li.SubItems(3)
                debit = 0
                balance = balance - IIf(li.SubItems(3) = "", 0, li.SubItems(3))
                TransDescription = li.SubItems(2)
                If Not SAVE_STATEMENT(AccNo, VoucherNo, credit, debit, balance, _
                TransDescription, Startdate, FinishDate, mTransDate, txtBankName, ErrorMessage) Then
                    If ErrorMessage <> "" Then
                        MsgBox ErrorMessage, vbInformation, Me.Caption
                        ErrorMessage = ""
                    End If
                End If
            End If
        End If
    Next I
   End If
   
    Set rs = oSaccoMaster.GetRecordset("Truncate table Statement1")
      
     sql = "insert into statement1(AccNo, VoucherNo, Credit, Debit, Balance, TransDescription, StartDate, FinishDate, TransDate, AccName)"
  
     sql = sql & "select AccNo, VoucherNo, Credit, Debit, Balance, TransDescription, StartDate, FinishDate, TransDate, AccName from  statement"
     
     Set rs = oSaccoMaster.GetRecordset(sql)
     
     
     
    If Not Execute_Command("Exec Delete_Statement", ErrorMessage) Then
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
    End If
    balance = 0
    Set rsCredits = oSaccoMaster.GetRecordset("select Balance from statement1  where voucherno='OpenBal'")
    If Not rsCredits.EOF Then
    balance = rsCredits.Fields(0)
    End If
    
   Set Rst = oSaccoMaster.GetRecordset("select * from  statement1 order by transdate")
     With Rst
      While Not .EOF
                    AccNo = cboBank
                    VoucherNo = Rst.Fields("VoucherNo")
                    credit = Rst.Fields("creDit")
                    debit = Rst.Fields("Debit")
                    mTransDate = Rst.Fields("Transdate")
                    balance = balance + IIf(Rst.Fields("Debit") > 0, Rst.Fields("Debit"), -Rst.Fields("creDit"))
                    TransDescription = IIf(IsNull(Rst.Fields("TransDescription")), "", Rst.Fields("TransDescription"))
                    If Not SAVE_STATEMENT(AccNo, VoucherNo, credit, debit, balance, _
                    TransDescription, Startdate, FinishDate, mTransDate, txtBankName, ErrorMessage) Then
                        If ErrorMessage <> "" Then
                            MsgBox ErrorMessage, vbInformation, Me.Caption
                            ErrorMessage = ""
                        End If
                    End If
                    
                    .MoveNext
      
      Wend
     End With
   
   
   Show_Sales_Crystal_Report "", "LedgerAcc Statement.rpt", ""
'Show_Sales_Crystal_Report "", "Member Statement2.rpt", "", True
   
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub Print_Report()
    On Error Resume Next
    PrintRec
    Exit Sub
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
Set xlBook = CreateObject("Excel.Application")
Set xlApp = CreateObject("Excel.Application")
Set xlSheet = CreateObject("Excel.ActiveSheet")
Set xlBook = CreateObject("Excel.Application")
    'Dim xlApp As excel.Application
    'set = xlBook As Excel.Workbook
    'set xlSheet As Excel.Worksheet
    Dim CreditPos As Long, DebitPos As Long, DirCreidts As Double, DirDebits As Double, _
    UnpresChqs As Double, UncrChqs As Double
'    Set xlApp = New excel.Application
    Set xlBook = xlApp.Workbooks.Add
    'xlApp.Run "Print_Report", dtpStartDate
    Exit Sub
    Set xlSheet = xlBook.ActiveSheet
    
    Dim NA As String, j As Long, RPos As Long, PPos As Long, PTotal As Double, RTotal As Double
    NA = transactions.Company_Name("name")
    xlApp.Visible = True
    DoEvents
    xlApp.ActiveSheet.Cells(1, 2).Select
    j = 4
    PPos = 6
    RPos = 6
    xlApp.ActiveSheet.Cells(RPos, 1) = "Date"
    xlApp.ActiveSheet.Cells(RPos, 2) = "Description"
    xlApp.ActiveSheet.Cells(RPos, 3) = "Chq No"
    xlApp.ActiveSheet.Cells(RPos, 4) = "Amount(Ksh)"
    xlApp.ActiveSheet.Cells(RPos, 5) = "Credited"
    xlApp.ActiveSheet.Cells(RPos, 6) = "Date"
    xlApp.ActiveSheet.Cells(RPos, 7) = "Description"
    xlApp.ActiveSheet.Cells(RPos, 8) = "Chq No"
    xlApp.ActiveSheet.Cells(RPos, 9) = "Amount(Ksh)"
    xlApp.ActiveSheet.Cells(RPos, 10) = "Presented"
    With xlApp.ActiveSheet.PageSetup
        .Orientation = xlApp.ActiveSheet.xlLandscape
    End With
    PPos = PPos + 1
    RPos = RPos + 1
    Startdate = DateSerial(year(dtpStartDate), month(dtpStartDate), Day(dtpStartDate) - 1)
    xlApp.ActiveSheet.Cells(PPos, 1) = Startdate
    xlApp.ActiveSheet.Cells(PPos, 6) = Startdate
    xlApp.ActiveSheet.Cells(PPos, 2) = "BAL B/FWD"
    xlApp.ActiveSheet.Cells(PPos, 7) = "BAL B/FWD"
    If CDbl(txtOpBalance) > 0 Then
        xlApp.ActiveSheet.Cells(PPos, 4) = txtOpBalance
        RTotal = RTotal + CDbl(txtOpBalance)
    Else
        xlApp.ActiveSheet.Cells(PPos, 9) = Format(CDbl(txtOpBalance) * (-1), "#,##0.00")
        PTotal = PTotal + CDbl(xlApp.ActiveSheet.Cells(PPos, 9))
    End If
    For I = 1 To ListView1.ListItems.Count
        j = j + 1
        If CDbl(ListView1.ListItems(I).SubItems(3)) = 0 Then 'Payments
            PPos = PPos + 1
            xlApp.ActiveSheet.Cells(PPos, 6) = ListView1.ListItems(I)
            xlApp.ActiveSheet.Cells(PPos, 7) = ListView1.ListItems(I).SubItems(2)
            If IsNumeric(ListView1.ListItems(I).SubItems(1)) Then
                xlApp.ActiveSheet.Cells(PPos, 8) = CStr(ListView1.ListItems(I).SubItems(1))
            Else
                xlApp.ActiveSheet.Cells(PPos, 8) = ""
            End If
            xlApp.ActiveSheet.Cells(PPos, 9) = ListView1.ListItems(I).SubItems(4)
            PTotal = PTotal + CDbl(xlApp.ActiveSheet.Cells(PPos, 9))
            If ListView1.ListItems(I).Checked Then
                xlApp.ActiveSheet.Cells(PPos, 10) = "/"
            End If
        Else 'Receipts
            RPos = RPos + 1
            xlApp.ActiveSheet.Cells(RPos, 1) = ListView1.ListItems(I)
            xlApp.ActiveSheet.Cells(RPos, 2) = ListView1.ListItems(I).SubItems(2)
            If IsNumeric(ListView1.ListItems(I).SubItems(1)) Then
                xlApp.ActiveSheet.Cells(RPos, 3) = CStr(ListView1.ListItems(I).SubItems(1))
            Else
                xlApp.ActiveSheet.Cells(RPos, 3) = ""
            End If
            xlApp.ActiveSheet.Cells(RPos, 4) = ListView1.ListItems(I).SubItems(3)
            RTotal = RTotal + CDbl(xlApp.ActiveSheet.Cells(RPos, 4))
            If ListView1.ListItems(I).Checked Then
                xlApp.ActiveSheet.Cells(RPos, 5) = "/"
            End If
        End If
    Next I
    If RPos > PPos Then
        RPos = RPos + 1
    Else
        RPos = PPos + 1
    End If
    If RTotal > PTotal Then
        xlApp.ActiveSheet.Cells(RPos, 4) = Format(RTotal, Cfmt)
        xlApp.ActiveSheet.Cells(RPos, 9) = Format(RTotal, Cfmt)
        xlApp.ActiveSheet.Cells(RPos - 1, 4) = Format(RTotal - PTotal, Cfmt)
        xlApp.ActiveSheet.Cells(RPos - 1, 3) = "BAL C/FWD"
    Else
        xlApp.ActiveSheet.Cells(RPos, 4) = Format(PTotal, Cfmt)
        xlApp.ActiveSheet.Cells(RPos, 9) = Format(PTotal, Cfmt)
        xlApp.ActiveSheet.Cells(RPos - 1, 4) = Format(PTotal - RTotal, Cfmt)
        xlApp.ActiveSheet.Cells(RPos - 1, 3) = "BAL C/FWD"
    End If
    xlApp.ActiveSheet.Range("A6").Select
'    Range(xlApp.xlSheet.selection, xlApp.xlSheet.selection.End(xlToRight)).Select
'    Range(xlApp.xlSheet.selection, xlApp.xlSheet.selection.End(xlDown)).Select
    xlApp.ActiveSheet.Range("A6:J" & RPos - 1).Select
'    With selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlHairline
'        .ColorIndex = xlAutomatic
'    End With
'    With selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .Weight = xlHairline
'        .ColorIndex = xlAutomatic
'    End With
'    With selection.Borders(xlEdgeBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlHairline
'        .ColorIndex = xlAutomatic
'    End With
'    With selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlHairline
'        .ColorIndex = xlAutomatic
'    End With
'    With selection.Borders(xlInsideVertical)
'        .LineStyle = xlContinuous
'        .Weight = xlHairline
'        .ColorIndex = xlAutomatic
'    End With
'    With selection.Borders(xlInsideHorizontal)
'        .LineStyle = xlContinuous
'        .Weight = xlHairline
'        .ColorIndex = xlAutomatic
'    End With
'    xlApp.ActiveSheet.Range("D" & RPos).Select
'    With selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'    End With
'    With selection.Borders(xlEdgeBottom)
'        .LineStyle = xlDouble
'        .Weight = xlThick
'    End With
'    xlApp.ActiveSheet.Range("I" & RPos).Select
'    With selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'    End With
'    With selection.Borders(xlEdgeBottom)
'        .LineStyle = xlDouble
'        .Weight = xlThick
'    End With
'    Columns("D:D").Select
'    selection.NumberFormat = "#,##0.00"
'    Columns("I:I").Select
'    selection.NumberFormat = "#,##0.00"
'    xlApp.ActiveSheet.Range("B1:B4").Select
'    selection.Font.Bold = True
'    Rows("6:6").Select
'    selection.Font.Bold = True
'    xlApp.ActiveSheet.Cells.Select
'    With selection.Font
'        .name = "Arial"
'        .Size = 8
'    End With
    xlApp.ActiveSheet.Cells.EntireColumn.AutoFit
    xlApp.ActiveSheet.Cells(2, 2) = NA
    xlApp.ActiveSheet.Cells(3, 2) = "Cash Book Details"
    xlApp.ActiveSheet.Cells(4, 2) = "For the Period Between " & dtpStartDate & " and " & dtpFinishDate
    xlApp.ActiveSheet.Range("A1").Select
    xlApp.Sheets("Sheet2").Select
    xlApp.ActiveSheet.Cells(6, 2) = "Balance as per cash book"
    xlApp.ActiveSheet.Cells(6, 3) = Format(RTotal - PTotal, Cfmt)
    xlApp.ActiveSheet.Cells(7, 2) = "Add: Direct Credits"
    xlApp.ActiveSheet.Cells(7, 3) = Format(CDbl(txtBankCredits), Cfmt)
    xlApp.ActiveSheet.Cells(8, 2) = "Less: Direct Debits"
    xlApp.ActiveSheet.Cells(8, 3) = Format(CDbl(txtBankDebits), Cfmt)
    xlApp.ActiveSheet.Cells(9, 2) = "Bal as per adjusted cash book"
    xlApp.ActiveSheet.Cells(9, 3) = Format(CDbl(xlApp.ActiveSheet.Cells(6, 3)) + CDbl(xlApp.ActiveSheet.Cells(7, 3)) - CDbl(xlApp.ActiveSheet.Cells(8, 3)), Cfmt)
    xlApp.ActiveSheet.Cells(10, 2) = "Add: Unpresented Cheques"
    xlApp.ActiveSheet.Cells(10, 3) = Format(CDbl(txtUnpresentedChq), Cfmt)
    xlApp.ActiveSheet.Cells(11, 2) = "Less: Uncredited Cheques"
    xlApp.ActiveSheet.Cells(11, 3) = Format(CDbl(txtDeposits), Cfmt)
    xlApp.ActiveSheet.Cells(12, 2) = "Bal as per bank statement"
    xlApp.ActiveSheet.Cells(12, 3) = Format(CDbl(xlApp.ActiveSheet.Cells(9, 3)) + CDbl(xlApp.ActiveSheet.Cells(10, 3)) - CDbl(xlApp.ActiveSheet.Cells(11, 3)), Cfmt)
   'Columns("C:C").Select
'    selection.NumberFormat = "#,##0.00"
'    Columns("E:E").Select
'    selection.NumberFormat = "#,##0.00"
'    xlApp.ActiveSheet.Range("B1:B4").Select
'    selection.Font.Bold = True
'    xlApp.ActiveSheet.Cells(15, 2) = "Notes"
'    xlApp.ActiveSheet.Range("B15").Select
'    selection.Font.Bold = True
'    CreditPos = 16
'    DebitPos = 16
'    xlApp.ActiveSheet.Cells(CreditPos, 2) = "Direct Credits"
'    xlApp.ActiveSheet.Range("B" & CreditPos).Select
'    selection.Font.Bold = True
'    xlApp.ActiveSheet.Cells(DebitPos, 4) = "Direct Debits"
'    xlApp.ActiveSheet.Range("D" & CreditPos).Select
'    selection.Font.Bold = True
'    xlApp.ActiveSheet.Range("A1").Select
'    'XXXXXXXXXXXXXXXXXXXXX Direct Credits XXXXXXXXXXXXXX
'    If lvwCredits.ListItems.Count > 0 Then
'        For I = 1 To lvwCredits.ListItems.Count
'            CreditPos = CreditPos + 1
'            xlApp.ActiveSheet.Cells(CreditPos, 2) = lvwCredits.ListItems(I).SubItems(1)
'            xlApp.ActiveSheet.Cells(CreditPos, 3) = lvwCredits.ListItems(I).SubItems(3)
'            DirCreidts = DirCreidts + CDbl(lvwCredits.ListItems(I).SubItems(3))
'        Next I
'    End If
'    'XXXXXXXXXXXXXXXXXXXXX Direct Debits XXXXXXXXXXXXXXX
'    If lvwDebits.ListItems.Count > 0 Then
'        For I = 1 To lvwDebits.ListItems.Count
'            DebitPos = DebitPos + 1
'            xlApp.ActiveSheet.Cells(DebitPos, 4) = lvwDebits.ListItems(I).SubItems(1)
'            xlApp.ActiveSheet.Cells(DebitPos, 5) = lvwDebits.ListItems(I).SubItems(3)
'            DirDebits = DirDebits + CDbl(lvwDebits.ListItems(I).SubItems(3))
'        Next I
'    End If
    xlApp.ActiveSheet.Cells(DebitPos + 1, 5) = Format(DirDebits, Cfmt)
    xlApp.ActiveSheet.Cells(CreditPos + 1, 3) = Format(DirCreidts, Cfmt)
    If DebitPos > CreditPos Then
        CreditPos = DebitPos + 1
        DebitPos = DebitPos + 1
    Else
        DebitPos = CreditPos + 1
        CreditPos = CreditPos + 1
    End If
    'XXXXXXXXXXXXXXXXXXXXX Unpresented Cheques XXXXXXXXX
    CreditPos = CreditPos + 2
    DebitPos = DebitPos + 2
    xlApp.ActiveSheet.Cells(CreditPos, 2) = "Unpresented Cheques"
    xlApp.ActiveSheet.Range("B" & CreditPos).Select
    'selection.Font.Bold = True
    xlApp.ActiveSheet.Cells(DebitPos, 4) = "Uncredited Cheques"
    xlApp.ActiveSheet.Range("D" & CreditPos).Select
    'selection.Font.Bold = True
    For I = 1 To ListView1.ListItems.Count
        If Not ListView1.ListItems(I).Checked Then
            If ListView1.ListItems(I).SubItems(3) = "0.00" Then 'Unpresented Cheques
                CreditPos = CreditPos + 1
                xlApp.ActiveSheet.Cells(CreditPos, 2) = ListView1.ListItems(I).SubItems(2)
                xlApp.ActiveSheet.Cells(CreditPos, 3) = ListView1.ListItems(I).SubItems(4)
                UnpresChqs = UnpresChqs + CDbl(ListView1.ListItems(I).SubItems(4))
            Else 'Uncredited Cheques
                DebitPos = DebitPos + 1
                xlApp.ActiveSheet.Cells(DebitPos, 4) = ListView1.ListItems(I).SubItems(2)
                xlApp.ActiveSheet.Cells(DebitPos, 5) = ListView1.ListItems(I).SubItems(3)
                UncrChqs = UncrChqs + CDbl(ListView1.ListItems(I).SubItems(3))
            End If
        End If
    Next I
   ' Cells(DebitPos + 1, 5) = Format(UncrChqs, Cfmt)
    xlApp.ActiveSheet.Range("E" & DebitPos + 1).Select
   ' selection.Font.Bold = True
    'Cells(CreditPos + 1, 3) = Format(UnpresChqs, Cfmt)
    xlApp.ActiveSheet.Range("C" & CreditPos + 1).Select
   ' selection.Font.Bold = True
    'XXXXXXXXXXXXXXXXXXXXX Uncredited Cheques XXXXXXXXXX
    xlApp.ActiveSheet.Cells.Select
'    With selection.Font
'        .name = "Arial"
'        .Size = 8
'    End With
    xlApp.ActiveSheet.Cells.EntireColumn.AutoFit
    xlApp.ActiveSheet.Cells(1, 2) = "KONOIN TEA GROWERS SACCO"
    xlApp.ActiveSheet.Cells(2, 2) = "BANK RECONCILIATION STATEMENT"
    xlApp.ActiveSheet.Cells(3, 2) = "FOR THE PERIOD BETWEEN " & dtpStartDate & " AND " & dtpFinishDate
    xlApp.ActiveSheet.Cells(4, 2) = txtBankName
    xlApp.Sheets("Sheet1").Select
    xlApp.Sheets("Sheet1").name = "CashBook Details"
    xlApp.Sheets("Sheet2").name = "Bank Reconciliation Statement"
    xlApp.Sheets("CashBook Details").Select
    Set xlBook = Nothing
    Set xlApp = Nothing
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
    xlBook.Close
    Set xlApp = Nothing
End Sub

Private Sub cmdprint_Click()
    'On Error Resume Next
'    PrintRec
'    Exit Sub
'    Dim xlApp As excel.Application, xlBook As excel.Workbook, xlSheet As excel.Worksheet
'    Set xlApp = New excel.Application
'    Set xlBook = xlApp.Workbooks.Add
'    Set xlSheet = xlBook.ActiveSheet
'    Dim NA As String, j As Long, RPos As Long, PPos As Long, PTotal As Double, RTotal As Double
'    NA = transactions.Company_Name("name")
'    xlApp.Visible = True
'    DoEvents
'    xlSheet.Cells(1, 2).Select
'    j = 4
'    PPos = 6
'    RPos = 6
'    xlSheet.Cells(RPos, 1) = "Date"
'    xlSheet.Cells(RPos, 2) = "Description"
'    xlSheet.Cells(RPos, 3) = "Chq No"
'    xlSheet.Cells(RPos, 4) = "Amount(Ksh)"
'    xlSheet.Cells(RPos, 5) = "Credited"
'    xlSheet.Cells(RPos, 6) = "Date"
'    xlSheet.Cells(RPos, 7) = "Description"
'    xlSheet.Cells(RPos, 8) = "Chq No"
'    xlSheet.Cells(RPos, 9) = "Amount(Ksh)"
'    xlSheet.Cells(RPos, 10) = "Presented"
'    PPos = PPos + 1
'    RPos = RPos + 1
'    Startdate = DateSerial(year(dtpStartDate), month(dtpStartDate), Day(dtpStartDate) - 1)
'    xlSheet.Cells(PPos, 1) = Startdate
'    xlSheet.Cells(PPos, 6) = Startdate
'    xlSheet.Cells(PPos, 2) = "BAL B/FWD"
'    xlSheet.Cells(PPos, 7) = "BAL B/FWD"
'    If CDbl(txtOpBalance) > 0 Then
'        xlSheet.Cells(PPos, 4) = txtOpBalance
'        RTotal = RTotal + CDbl(txtOpBalance)
'    Else
'        xlSheet.Cells(PPos, 9) = Format(CDbl(txtOpBalance) * (-1), "#,##0.00")
'        PTotal = PTotal + CDbl(xlSheet.Cells(PPos, 9))
'    End If
'    For I = 1 To ListView1.ListItems.Count
'        j = j + 1
'        If CDbl(ListView1.ListItems(I).SubItems(3)) = 0 Then 'Payments
'            PPos = PPos + 1
'            xlSheet.Cells(PPos, 6) = ListView1.ListItems(I)
'            xlSheet.Cells(PPos, 7) = ListView1.ListItems(I).SubItems(2)
'            If IsNumeric(ListView1.ListItems(I).SubItems(1)) Then
'                xlSheet.Cells(PPos, 8) = CStr(ListView1.ListItems(I).SubItems(1))
'            Else
'                xlSheet.Cells(PPos, 8) = ""
'            End If
'            xlSheet.Cells(PPos, 9) = ListView1.ListItems(I).SubItems(4)
'            PTotal = PTotal + CDbl(xlSheet.Cells(PPos, 9))
'            If ListView1.ListItems(I).Checked Then
'                xlSheet.Cells(PPos, 10) = "/"
'            End If
'        Else 'Receipts
'            RPos = RPos + 1
'            xlSheet.Cells(RPos, 1) = ListView1.ListItems(I)
'            xlSheet.Cells(RPos, 2) = ListView1.ListItems(I).SubItems(2)
'            If IsNumeric(ListView1.ListItems(I).SubItems(1)) Then
'                xlSheet.Cells(RPos, 3) = CStr(ListView1.ListItems(I).SubItems(1))
'            Else
'                xlSheet.Cells(RPos, 3) = ""
'            End If
'            xlSheet.Cells(RPos, 4) = ListView1.ListItems(I).SubItems(3)
'            RTotal = RTotal + CDbl(xlSheet.Cells(RPos, 4))
'            If ListView1.ListItems(I).Checked Then
'                xlSheet.Cells(RPos, 5) = "/"
'            End If
'        End If
'    Next I
'    If RPos > PPos Then
'        RPos = RPos + 1
'    Else
'        RPos = PPos + 1
'    End If
'    If RTotal > PTotal Then
'        xlSheet.Cells(RPos, 4) = Format(RTotal, Cfmt)
'        xlSheet.Cells(RPos, 9) = Format(RTotal, Cfmt)
'        xlSheet.Cells(RPos - 1, 4) = Format(RTotal - PTotal, Cfmt)
'        xlSheet.Cells(RPos - 1, 3) = "BAL C/FWD"
'    Else
'        xlSheet.Cells(RPos, 4) = Format(PTotal, Cfmt)
'        xlSheet.Cells(RPos, 9) = Format(PTotal, Cfmt)
'        xlSheet.Cells(RPos - 1, 4) = Format(PTotal - RTotal, Cfmt)
'        xlSheet.Cells(RPos - 1, 3) = "BAL C/FWD"
'    End If
'    xlSheet.Range("A6").Select
''    range(Selection, Selection.End(xlToRight)).Select
''    range(Selection, Selection.End(xlDown)).Select
''    range("A6:J" & RPos - 1).Select
'    With selection.Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlHairline
'        .ColorIndex = xlAutomatic
'    End With
'    With selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .Weight = xlHairline
'        .ColorIndex = xlAutomatic
'    End With
'    With selection.Borders(xlEdgeBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlHairline
'        .ColorIndex = xlAutomatic
'    End With
'    With selection.Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlHairline
'        .ColorIndex = xlAutomatic
'    End With
'    With selection.Borders(xlInsideVertical)
'        .LineStyle = xlContinuous
'        .Weight = xlHairline
'        .ColorIndex = xlAutomatic
'    End With
'    With selection.Borders(xlInsideHorizontal)
'        .LineStyle = xlContinuous
'        .Weight = xlHairline
'        .ColorIndex = xlAutomatic
'    End With
'    Range("D" & RPos).Select
'    With selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'    End With
'    With selection.Borders(xlEdgeBottom)
'        .LineStyle = xlDouble
'        .Weight = xlThick
'    End With
'    Range("I" & RPos).Select
'    With selection.Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'    End With
'    With selection.Borders(xlEdgeBottom)
'        .LineStyle = xlDouble
'        .Weight = xlThick
'    End With
'    Columns("D:D").Select
'    selection.NumberFormat = "#,##0.00"
'    Columns("I:I").Select
'    selection.NumberFormat = "#,##0.00"
'    Range("B1:B4").Select
'    selection.Font.Bold = True
'    Rows("6:6").Select
'    selection.Font.Bold = True
'    xlSheet.Cells.Select
'    With selection.Font
'        .name = "Arial"
'        .Size = 8
'    End With
'    xlSheet.Cells.EntireColumn.AutoFit
'    xlSheet.Cells(2, 2) = NA
'    xlSheet.Cells(3, 2) = "Cash Book Details"
'    xlSheet.Cells(4, 2) = "For the Period Between " & dtpStartDate & " and " & dtpFinishDate
'    Range("A1").Select
'    Sheets("Sheet2").Select
'    xlSheet.Cells(1, 2) = "KONOIN TEA GROWERS SACCO"
'    xlSheet.Cells(2, 2) = "BANK RECONCILIATION STATEMENT"
'    xlSheet.Cells(3, 2) = "FOR THE PERIOD BETWEEN " & dtpStartDate & " AND " & dtpFinishDate
'    xlSheet.Cells(4, 2) = txtBankName
'    xlSheet.Cells(6, 2) = "Balance as per cash book"
'    xlSheet.Cells(6, 3) = Format(RTotal - PTotal, Cfmt)
'    xlSheet.Cells(7, 2) = "Add: Direct Credits"
'    xlSheet.Cells(7, 3) = Format(CDbl(txtBankCredits), Cfmt)
'    xlSheet.Cells(8, 2) = "Less: Direct Debits"
'    xlSheet.Cells(8, 3) = Format(CDbl(txtBankDebits), Cfmt)
'    xlSheet.Cells(9, 2) = "Bal as per adjusted cash book"
'    xlSheet.Cells(9, 3) = Format(CDbl(xlSheet.Cells(6, 3)) + CDbl(xlSheet.Cells(7, 3)) - CDbl(xlSheet.Cells(8, 3)), Cfmt)
'    xlSheet.Cells(10, 2) = "Add: Unpresented Cheques"
'    xlSheet.Cells(10, 3) = Format(CDbl(txtUnpresentedChq), Cfmt)
'    xlSheet.Cells(11, 2) = "Less: Uncredited Cheques"
'    xlSheet.Cells(11, 3) = Format(CDbl(txtDeposits), Cfmt)
'    xlSheet.Cells(12, 2) = "Bal as per bank statement"
'    xlSheet.Cells(12, 3) = Format(CDbl(xlSheet.Cells(9, 3)) + CDbl(xlSheet.Cells(10, 3)) - CDbl(xlSheet.Cells(11, 3)), Cfmt)
'    Columns("C:C").Select
'    selection.NumberFormat = "#,##0.00"
'    Range("B1:B4").Select
'    selection.Font.Bold = True
'    xlSheet.Cells.Select
'    With selection.Font
'        .name = "Arial"
'        .Size = 8
'    End With
'    xlSheet.Cells.EntireColumn.AutoFit
'    Range("A1").Select
'    'XXXXXXXXXXXXXXXXXXXXX Direct Credits XXXXXXXXXXXXXX
'    'XXXXXXXXXXXXXXXXXXXXX Direct Debits XXXXXXXXXXXXXXX
'    'XXXXXXXXXXXXXXXXXXXXX Unpresented Cheques XXXXXXXXX
'    'XXXXXXXXXXXXXXXXXXXXX Uncredited Cheques XXXXXXXXXX
'    Sheets("Sheet1").Select
'    Sheets("Sheet1").name = "CashBook Details"
'    Sheets("Sheet2").name = "Bank Reconciliation Statement"
'    Sheets("Sheet3").Select
'    Sheets("Sheet3").name = "Notes"
'    xlSheet.Cells(1, 2) = "Notes"
'    xlApp.Quit
'    Set xlBook = Nothing
'    Set xlApp = Nothing
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdRemove_Click()
    On Error GoTo sysError
    Dim vno As String, transdate As Date, Amount As Double
    If ListView1.ListItems.Count > 0 Then
        If MsgBox("Do you want to remove the selected entry?", vbQuestion + _
        vbYesNo, "Remove Entry") = vbNo Then
            Exit Sub
        End If
        If CDbl(ListView1.SelectedItem.SubItems(4)) = 0 Then
            Amount = CDbl(ListView1.SelectedItem.SubItems(3))
        Else
            Amount = CDbl(ListView1.SelectedItem.SubItems(4))
        End If
        vno = ListView1.SelectedItem.SubItems(1)
        transdate = ListView1.SelectedItem
        If Amount > 0 Then
            If Not Delete_CustBal(cboBank.Text, transdate, vno, Amount, ErrorMessage) Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        ListView1.ListItems.Remove ListView1.SelectedItem.Index
    Else
        MsgBox "There are no entries to be removed", vbExclamation, "Remove entry"
    End If
    Calculate_Summaries
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdtransfer_Click()
    On Error GoTo sysError
    Dim FromAcc As String, ToAcc As String, DocumentNo As String, TransT As String, _
    transdate As Date, mCredit As String
    Dim rsTransfer As New Recordset
    If ListView1.ListItems.Count > 0 Then
        Set li = ListView1.SelectedItem
        transdate = li
        DocumentNo = CStr(li.SubItems(7))
        TransT = CStr(li.SubItems(1))
        Select Case CDbl(li.SubItems(3)) 'PAYMENTS
            Case 0 'RECEIPTS
            mCredit = "DR"
            Case Else 'PAYMENTS
            mCredit = "CR"
        End Select
        'Set rsTransfer = OSACCOMASTER.GetRecordset("Set DateFormat DMY Update CUSTOMERBALANCE" _
        & " Set AccNo='" & txtAccNo & "' where AccNo='" & cboBank & "' and VNo='" & TransT & _
        "' and ChequeNo='" & DocumentNo & "' and TransDate='" & TransDate & "' and TransType='" _
        & mCredit & "'")
        If Not Execute_Command("Set DateFormat DMY Update CUSTOMERBALANCE Set AccNo='" & _
        txtAccNo & "' where AccNo='" & cboBank & "' and VNo='" & TransT & "' and CustomerbalanceID='" _
        & DocumentNo & "' and TransDate='" & transdate & "'", ErrorMessage) Then
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
        fraTransfer.Visible = False
        'cmdLoad_Click
        ListView1.ListItems.Remove ListView1.SelectedItem.Index
    End If
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdTransferFunds_Click()
    On Error GoTo sysError
    'frmTransferFunds.Show , Me
    fraTransfer.Visible = True
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub



Private Sub Command2_Click()
reportname = "ReconReport.rpt"
STRFORMULA = "" '"{BankRecon.ReconDate}='" & dtpFinishDate & "'"
Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
End Sub

Private Sub dtpStartDate_Change()
'If cboBank <> "" Then
'    sql = " set dateformat dmy   select (select ISNULL(sum(Amount),0) from CustomerBalance where AccNO='" & cboBank & "' and TransDate>='" & dtpStartDate & "' "
'    sql = sql & " and Transdate<='" & dtpFinishDate & "' and transType='DR')-"
'    sql = sql & " (select ISNULL(sum(Amount),0) from CustomerBalance where AccNO='" & cboBank & "' and TransDate>='" & dtpStartDate & "'  "
'    sql = sql & "  and Transdate<='" & dtpFinishDate & "' and transType='CR') as OpeningBal "
'     Set rst = OSACCOMASTER.GetRecordset(sql)
'     If Not rst.EOF Then
'     txtOpeningBal = rst.Fields("OpeningBal")
'     Else
'     txtOpeningBal = 0
'     End If
'Else
'  txtOpeningBal = 0
'  End If
'  txtOpBalance = txtOpeningBal
End Sub

Private Sub dtpstartdate_Click()
dtpStartDate_Change
End Sub

Private Sub Form_Load()
    On Error GoTo sysError
    dtpReconciliation = Format(Get_Server_Date, " dd-MM-yyyy")
    dtpStartDate = Get_Server_Date
    dtpFinishDate = Get_Server_Date
    dtpDrTransDate = Format(Get_Server_Date, " dd-MM-yyyy")
    dtpCrTransDate = Format(Get_Server_Date, " dd-MM-yyyy")
    dtpStatement = Format(Get_Server_Date, " dd-MM-yyyy")
    
    ''' load Banks
    cboBank.Clear
    Set Rst = oSaccoMaster.GetRecordset("Select distinct GlAccNo From Banks where GlAccNo<>''")
    With Rst
    While Not .EOF
    cboBank.AddItem (Rst.Fields(0))
     .MoveNext
    Wend
    End With
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub ListView1_Click()
    Load_Statement
    Calculate_Summaries
End Sub

Private Sub ListView1_DblClick()
'    If ListView1.ListItems.Count > 0 Then
'        Set li = ListView1.SelectedItem
'        mTransDate = CDate(li)
'        TransNo = li.SubItems(6)
'        mDocNo = li.SubItems(5)
'    End If
'    frmLedgers.Show vbModal, Me
    ListView1.SelectedItem.SubItems(6) = dtpStatement
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Calculate_Summaries
End Sub

Private Sub ListView2_Click()
    On Error GoTo sysError
    If ListView2.ListItems.Count > 0 Then
        Set li = ListView2.SelectedItem
        txtAccNo = li
        txtAccName = ListView2.SelectedItem.SubItems(1)
        ListView2.ListItems.Clear
        ListView2.Visible = False
    End If
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub lvwAccounts_Click()
    On Error GoTo sysError
    Dim BankName As String
    If lvwAccounts.ListItems.Count > 0 Then
        Set li = lvwAccounts.SelectedItem
        BankName = lvwAccounts.SelectedItem.SubItems(1)
        cboBank = li
        txtBankName = BankName
        lvwAccounts.ListItems.Clear
        lvwAccounts.Visible = False
    End If
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub lvwCrAccounts_Click()
    On Error GoTo sysError
    If lvwCrAccounts.ListItems.Count > 0 Then
        txtCrAccName = lvwCrAccounts.SelectedItem
        txtCrAccNo = lvwCrAccounts.SelectedItem.SubItems(1)
    End If
    lvwCrAccounts.Visible = False
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub lvwDrAccounts_Click()
    On Error GoTo sysError
    If lvwDrAccounts.ListItems.Count > 0 Then
        txtDrAccName = lvwDrAccounts.SelectedItem
        txtDrAccNo = lvwDrAccounts.SelectedItem.SubItems(1)
    End If
    lvwDrAccounts.Visible = False
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub optTransfer_Click()
    If optTransfer.value = vbChecked Then
        Frame1.Visible = True
    Else
       Frame1.Visible = False
    End If
End Sub

Private Sub txtAccName_Change()
    On Error GoTo sysError
    Dim rsAccount As New Recordset
    ListView2.ListItems.Clear
    If Trim$(txtAccName) <> "" Then
        If Not Editing Then
            Set rsAccount = oSaccoMaster.GetRecordset("Select * From GLSETUP where " _
            & "GLAccName like '%" & txtAccName & "%'")
            With rsAccount
                If .State = adStateOpen Then
                    If Not .EOF Then
                        ListView2.Visible = True
                        While Not .EOF
                            Set li = ListView2.ListItems.Add(, , IIf(IsNull(!AccNo), "", !AccNo))
                            li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
                            .MoveNext
                        Wend
                        If ListView2.ListItems.Count = 1 Then
                            txtAccNo = li
                            txtAccName = li.SubItems(1)
                            ListView2.ListItems.Clear
                            ListView2.Visible = False
                        End If
                    Else
                        ListView2.Visible = False
                    End If
                End If
            End With
        End If
    Else
        ListView2.Visible = False
    End If
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtBankBal_Change()
If txtBankCredits = "" Then txtBankCredits = 0
If txtBankDebits = "" Then txtBankDebits = 0
If txtBankBal = "" Then txtBankBal = 0
If txtCBBalance = "" Then txtCBBalance = 0
If txtUnpresentedChq = "" Then txtUnpresentedChq = 0
If txtDeposits = "" Then txtDeposits = 0
If txtBankBal = "" Then txtBankBal = 0
 txtDifference = Format(CDbl(txtCBBalance) + CDbl(txtUnpresentedChq) - CDbl(txtDeposits) + CDbl(txtBankCredits) - CDbl(txtBankDebits) - CDbl(txtBankBal), Cfmt)
End Sub



Private Sub txtBankBal_Click()
 txtBankBal_Change
End Sub

Private Sub txtbankbalance_Change()
    On Error Resume Next
    txtBankBal = txtbankbalance
    txtBankBal = Format(txtbankbalance, Cfmt)
End Sub

Private Sub txtBankName_Change()
    'On Error GoTo SysError
    Dim rsAccount As New Recordset
    lvwAccounts.ListItems.Clear
    If Trim$(txtBankName) <> "" Then
        If Not Editing Then
            Set rsAccount = oSaccoMaster.GetRecordset("Select * From GLSETUP where " _
            & "GLAccName like '%" & txtBankName & "%'")
            With rsAccount
                If .State = adStateOpen Then
                    If Not .EOF Then
                        lvwAccounts.Visible = True
                        lvwAccounts.Height = 1560
                        While Not .EOF
                            Set li = lvwAccounts.ListItems.Add(, , IIf(IsNull(!AccNo), "", !AccNo))
                            li.SubItems(1) = IIf(IsNull(!GlAccName), "", !GlAccName)
                            .MoveNext
                        Wend
                        If lvwAccounts.ListItems.Count = 1 Then
                            cboBank = li
                            txtBankName = li.SubItems(1)
                            lvwAccounts.ListItems.Clear
                            lvwAccounts.Visible = False
                        End If
                    Else
                        lvwAccounts.Visible = False
                    End If
                End If
            End With
        End If
    Else
        lvwAccounts.Visible = False
    End If
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtCBBalance_Change()
txtbankbalance_Change
End Sub

Private Sub txtCrAccName_Change()
    On Error GoTo sysError
    Dim rsAcc As New Recordset
    lvwCrAccounts.ListItems.Clear
    If Trim$(txtCrAccName) <> "" Then
        Set rsAcc = oSaccoMaster.GetRecordset("Select AccNo,GLAccName From GLSETUP " _
        & "where GLACCName like '%" & txtCrAccName & "%'")
        With rsAcc
            If Not .EOF Then
                lvwCrAccounts.Visible = True
                While Not .EOF
                    Set li = lvwCrAccounts.ListItems.Add(, , !GlAccName)
                    li.SubItems(1) = IIf(IsNull(!AccNo), "", !AccNo)
                    .MoveNext
                Wend
            Else
                lvwCrAccounts.Visible = False
            End If
        End With
    Else
        lvwCrAccounts.Visible = False
    End If
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtCrAmount_KeyPress(KeyAscii As Integer)
    On Error GoTo sysError
    Select Case KeyAscii
        Case 48 To 57
        Case Is = 46
        Case Is = 8
        Case 13
        If Trim$(txtCrAmount) Then
            Set li = lvwCredits.ListItems.Add(, , dtpCrTransDate)
            li.SubItems(1) = txtCrNarration
            li.SubItems(2) = txtDrAccNo
            li.SubItems(3) = Format(txtCrAmount, Cfmt)
            li.SubItems(4) = txtCrDocumentNo
        End If
        Case Else
        KeyAscii = 0
    End Select
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub


   




Private Sub txtDrAccName_Change()
    On Error GoTo sysError
    Dim rsAcc As New Recordset
    lvwDrAccounts.ListItems.Clear
    If Trim$(txtDrAccName) <> "" Then
        Set rsAcc = oSaccoMaster.GetRecordset("Select AccNo,GLAccName From GLSETUP " _
        & "where GLACCName like '%" & txtDrAccName & "%'")
        With rsAcc
            If Not .EOF Then
                lvwDrAccounts.Visible = True
                While Not .EOF
                    Set li = lvwDrAccounts.ListItems.Add(, , !GlAccName)
                    li.SubItems(1) = IIf(IsNull(!AccNo), "", !AccNo)
                    .MoveNext
                Wend
            Else
                lvwDrAccounts.Visible = False
            End If
        End With
    Else
        lvwDrAccounts.Visible = False
    End If
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtDrAmount_KeyPress(KeyAscii As Integer)
    On Error GoTo sysError
    Select Case KeyAscii
        Case 48 To 57
        Case Is = 46
        Case Is = 8
        Case 13
'        If Trim(txtDrAmount) <> "" Then
'            Set li = lvwCredits.ListItems.Add(, , dtpDrTransDate)
'            li.SubItems(1) = txtDrNarration
'            li.SubItems(2) = txtDrAccNo
'            li.SubItems(3) = Format(txtDrAmount, CfMt)
'            li.SubItems(4) = txtDrDocumentNo
'        End If
        Case Else
        KeyAscii = 0
    End Select
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtOpBalance_Change()
    Calculate_Summaries
End Sub

Private Sub txtOpeningBal_Change()
    Calculate_Summaries
End Sub
