VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPaymentPosting 
   BackColor       =   &H00808000&
   Caption         =   "GENERAL TRANSACTIONS"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9975
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPaymentPosting.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7395
      Left            =   0
      TabIndex        =   0
      Top             =   660
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   13044
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   8421376
      TabCaption(0)   =   "RECEIPTS"
      TabPicture(0)   =   "frmPaymentPosting.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdclose"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "PAYMENTS"
      TabPicture(1)   =   "frmPaymentPosting.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C000&
         Height          =   6720
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   525
         Width           =   9510
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Print"
            Height          =   375
            Index           =   3
            Left            =   1920
            TabIndex        =   60
            Top             =   6840
            Width           =   1455
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Print Voucher"
            Height          =   375
            Index           =   0
            Left            =   4200
            TabIndex        =   59
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtPayee 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   0
            Left            =   1800
            TabIndex        =   47
            Top             =   2280
            Width           =   3225
         End
         Begin VB.TextBox txtParticulars 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   0
            Left            =   1800
            TabIndex        =   46
            Top             =   1725
            Width           =   3225
         End
         Begin VB.TextBox txtDistributed 
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
            Height          =   285
            Index           =   0
            Left            =   7800
            MaxLength       =   15
            TabIndex        =   45
            Text            =   "0"
            Top             =   2295
            Width           =   1380
         End
         Begin VB.TextBox txtBalance 
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
            Height          =   315
            Index           =   0
            Left            =   7830
            MaxLength       =   15
            TabIndex        =   44
            Text            =   "0"
            Top             =   2655
            Width           =   1380
         End
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
            Height          =   375
            Index           =   0
            Left            =   1800
            TabIndex        =   43
            Top             =   720
            Width           =   1860
         End
         Begin VB.TextBox txtAmountPaid 
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
            Height          =   285
            Index           =   0
            Left            =   6255
            MaxLength       =   15
            TabIndex        =   42
            Text            =   "0"
            Top             =   2310
            Width           =   1380
         End
         Begin VB.ComboBox cboMode 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            ItemData        =   "frmPaymentPosting.frx":0342
            Left            =   1800
            List            =   "frmPaymentPosting.frx":0355
            TabIndex        =   41
            Text            =   "Cash"
            Top             =   1275
            Width           =   1425
         End
         Begin VB.CommandButton cmdReceipt 
            Caption         =   "<>"
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
            Index           =   0
            Left            =   3705
            TabIndex        =   40
            Top             =   720
            Width           =   345
         End
         Begin VB.TextBox txtmode 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   3300
            TabIndex        =   39
            Top             =   1290
            Width           =   1380
         End
         Begin VB.CommandButton cmdupdatereceipt 
            Caption         =   "&Post"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   38
            Top             =   6240
            Width           =   1425
         End
         Begin VB.CommandButton cmdBank 
            Caption         =   "<>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   3705
            TabIndex        =   37
            Top             =   285
            Width           =   345
         End
         Begin VB.ComboBox cboBanks 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            ItemData        =   "frmPaymentPosting.frx":0378
            Left            =   1800
            List            =   "frmPaymentPosting.frx":037A
            TabIndex        =   36
            Top             =   270
            Width           =   1830
         End
         Begin VB.ComboBox cboAccno 
            Height          =   330
            Index           =   0
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   2730
            Width           =   1200
         End
         Begin VB.TextBox txtAccNames 
            Height          =   315
            Index           =   0
            Left            =   1815
            TabIndex        =   34
            Top             =   2730
            Width           =   3225
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add>>"
            Height          =   345
            Index           =   0
            Left            =   5520
            TabIndex        =   33
            Top             =   3120
            Width           =   1170
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "<<Remove"
            Height          =   345
            Index           =   0
            Left            =   6720
            TabIndex        =   32
            Top             =   3120
            Width           =   1170
         End
         Begin VB.CommandButton cmdAcctsSearch 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   1485
            Picture         =   "frmPaymentPosting.frx":037C
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   2730
            Width           =   330
         End
         Begin MSComctlLib.ListView lvwNtrans 
            Height          =   2340
            Index           =   0
            Left            =   150
            TabIndex        =   48
            Top             =   3795
            Width           =   9060
            _ExtentX        =   15981
            _ExtentY        =   4128
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
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Trans Description"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Amount"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Accno"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Payee"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Payee"
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
            Index           =   1
            Left            =   465
            TabIndex        =   57
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Particulars"
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
            Index           =   1
            Left            =   375
            TabIndex        =   56
            Top             =   1875
            Width           =   885
         End
         Begin VB.Label Label11 
            Caption         =   "Distributed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   7845
            TabIndex        =   55
            Top             =   2085
            Width           =   1005
         End
         Begin VB.Label Amount 
            Caption         =   "Balance"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   6945
            TabIndex        =   54
            Top             =   2850
            Width           =   780
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Receipt No"
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
            Index           =   1
            Left            =   480
            TabIndex        =   53
            Top             =   802
            Width           =   870
         End
         Begin VB.Label Label13 
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
            Left            =   6240
            TabIndex        =   52
            Top             =   2085
            Width           =   660
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Payment Mode"
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
            Left            =   420
            TabIndex        =   51
            Top             =   1320
            Width           =   1230
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Bank (DR)"
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
            Left            =   450
            TabIndex        =   50
            Top             =   330
            Width           =   780
         End
         Begin VB.Label lblbankname 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   4200
            TabIndex        =   49
            Top             =   285
            Width           =   4095
         End
      End
      Begin VB.CommandButton cmdclose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8565
         TabIndex        =   28
         Top             =   7560
         Width           =   1005
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFF00&
         Height          =   6720
         Index           =   0
         Left            =   -74880
         TabIndex        =   1
         Top             =   525
         Width           =   9510
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Print Voucher"
            Height          =   375
            Index           =   1
            Left            =   4200
            TabIndex        =   58
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "<<Remove"
            Height          =   345
            Index           =   1
            Left            =   6720
            TabIndex        =   29
            Top             =   3120
            Width           =   1170
         End
         Begin VB.TextBox txtPayee 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   1
            Left            =   1800
            TabIndex        =   26
            Top             =   2280
            Width           =   3225
         End
         Begin VB.TextBox txtParticulars 
            Appearance      =   0  'Flat
            Height          =   375
            Index           =   1
            Left            =   1800
            TabIndex        =   24
            Top             =   1725
            Width           =   3225
         End
         Begin VB.TextBox txtDistributed 
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
            Height          =   285
            Index           =   1
            Left            =   7800
            MaxLength       =   9
            TabIndex        =   22
            Text            =   "0"
            Top             =   2295
            Width           =   1380
         End
         Begin VB.TextBox txtBalance 
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
            Height          =   315
            Index           =   1
            Left            =   7830
            TabIndex        =   20
            Text            =   "0"
            Top             =   2655
            Width           =   1380
         End
         Begin VB.TextBox txtVoucherNo 
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
            Height          =   375
            Index           =   1
            Left            =   1800
            TabIndex        =   16
            Top             =   720
            Width           =   1860
         End
         Begin VB.TextBox txtAmountPaid 
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
            Height          =   285
            Index           =   1
            Left            =   6255
            MaxLength       =   9
            TabIndex        =   15
            Text            =   "0"
            Top             =   2310
            Width           =   1380
         End
         Begin VB.ComboBox cboMode 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            ItemData        =   "frmPaymentPosting.frx":047E
            Left            =   1800
            List            =   "frmPaymentPosting.frx":0491
            TabIndex        =   14
            Text            =   "Cash"
            Top             =   1275
            Width           =   1425
         End
         Begin VB.CommandButton cmdVoucher 
            Caption         =   "<>"
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
            Index           =   1
            Left            =   3705
            TabIndex        =   13
            Top             =   720
            Width           =   345
         End
         Begin VB.TextBox txtmode 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   3300
            TabIndex        =   12
            Top             =   1290
            Width           =   1380
         End
         Begin VB.CommandButton cmdupdatereceipt 
            Caption         =   "&Post"
            Height          =   375
            Index           =   1
            Left            =   255
            TabIndex        =   11
            Top             =   6225
            Width           =   1425
         End
         Begin VB.CommandButton cmdBank 
            Caption         =   "<>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   3705
            TabIndex        =   8
            Top             =   285
            Width           =   345
         End
         Begin VB.ComboBox cboBanks 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            ItemData        =   "frmPaymentPosting.frx":04B4
            Left            =   1800
            List            =   "frmPaymentPosting.frx":04B6
            TabIndex        =   7
            Top             =   270
            Width           =   1830
         End
         Begin VB.ComboBox cboAccno 
            Height          =   330
            Index           =   1
            Left            =   315
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2730
            Width           =   1200
         End
         Begin VB.TextBox txtAccNames 
            Height          =   315
            Index           =   1
            Left            =   1815
            TabIndex        =   4
            Top             =   2730
            Width           =   3225
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add>>"
            Height          =   345
            Index           =   1
            Left            =   5520
            TabIndex        =   3
            Top             =   3120
            Width           =   1170
         End
         Begin VB.CommandButton cmdAcctsSearch 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   1485
            Picture         =   "frmPaymentPosting.frx":04B8
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   2730
            Width           =   330
         End
         Begin MSComctlLib.ListView lvwNtrans 
            Height          =   2340
            Index           =   1
            Left            =   270
            TabIndex        =   6
            Top             =   3795
            Width           =   9060
            _ExtentX        =   15981
            _ExtentY        =   4128
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
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Trans Description"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Amount"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Accno"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Payee"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Payee"
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
            Index           =   0
            Left            =   465
            TabIndex        =   27
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Particulars"
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
            Left            =   375
            TabIndex        =   25
            Top             =   1875
            Width           =   885
         End
         Begin VB.Label Label7 
            Caption         =   "Distributed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   7845
            TabIndex        =   23
            Top             =   2085
            Width           =   1005
         End
         Begin VB.Label Amount 
            Caption         =   "Balance"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   6945
            TabIndex        =   21
            Top             =   2850
            Width           =   780
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Voucher No"
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
            Index           =   0
            Left            =   480
            TabIndex        =   19
            Top             =   802
            Width           =   960
         End
         Begin VB.Label Label11 
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
            Index           =   0
            Left            =   6240
            TabIndex        =   18
            Top             =   2085
            Width           =   660
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Payment Mode"
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
            Index           =   0
            Left            =   420
            TabIndex        =   17
            Top             =   1320
            Width           =   1230
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Bank (CR)"
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
            Left            =   450
            TabIndex        =   10
            Top             =   330
            Width           =   795
         End
         Begin VB.Label lblbankname 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   4200
            TabIndex        =   9
            Top             =   285
            Width           =   4095
         End
      End
   End
   Begin MSComDlg.CommonDialog dlg9 
      Left            =   10455
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DTPDatedeposited 
      Height          =   375
      Left            =   8115
      TabIndex        =   61
      Top             =   120
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   119013377
      CurrentDate     =   40421
   End
   Begin VB.Label Label8 
      Caption         =   "Transaction Date"
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
      Left            =   5880
      TabIndex        =   62
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmPaymentPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim totalamount As Currency
Dim pushed As Currency
Dim objLabelEdit As LabelEdit
Dim objLabelEdit2 As LabelEdit
Dim objLabelEdit3 As LabelEdit
Dim interestAcc As String, LoanAcc As String
Dim j As Integer
Dim balance As Double
Dim amt As Double
Dim WhatFor As String
Dim DRaccno As String, Craccno As String
Dim detamt As Double
Dim OtherDeductions As Double, TotalBalance As Double

Private Sub cboAccno_Change(Index As Integer)
    Dim AccNo As String
    AccNo = cboAccno(Index).Text
    sql = "select GLACCNAME from glsetup where accno='" & AccNo & "'"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If Not rs.EOF Then
        txtAccNames(Index).Text = rs(0)
    End If
End Sub

Private Sub cboAccno_Click(Index As Integer)
    cboAccno_Change (Index)
End Sub



Private Sub cboAccno_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBalance(Index).SetFocus
        txtBalance(Index).Text = 0
    End If
End Sub

Private Sub cboBanks_Change(Index As Integer)
    On Error GoTo sysError
        Dim Account As Account_Details
        'Account = Get_Account_Details(cboBanks, "BOSA", ErrorMessage)
        sql = "select GlAccName from Glsetup where accno='" & cboBanks(Index).Text & "'"
        Set Rst = oSaccoMaster.GetRecordset(sql)
        If Not Rst.EOF Then
            lblbankname(Index).Caption = Rst(0)
        Else
            lblbankname(Index).Caption = ""
        End If
        Exit Sub
sysError:
        MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cboBanks_Click(Index As Integer)
    cboBanks_Change (Index)
End Sub



Private Sub cmdAcctsSearch_Click(Index As Integer)
    frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            cboAccno(Index) = SearchValue
            SearchValue = ""
            Continue = False
        End If
    End If
End Sub

Private Sub cmdAdd_Click(Index As Integer)
    On Error GoTo sysError
    If txtParticulars(Index).Text = "" Then
        MsgBox "Write the particulars", vbInformation
        txtParticulars(Index).SetFocus
        Exit Sub
    End If
    If txtPayee(Index).Text = "" Then
        MsgBox "Input the Payee", vbInformation
         txtPayee(Index).SetFocus
        Exit Sub
    End If
    If txtAmountPaid(Index).Text = "" Or txtAmountPaid(Index).Text = 0 Then
        MsgBox "Input the Amount", vbInformation
        txtAmountPaid(Index).SetFocus
        Exit Sub
    End If
    If cboAccno(Index).Text = "" Then
        MsgBox "Select Respective Contra Transactions Gl Ledger", vbInformation, Me.Caption
         cboAccno(Index).SetFocus
        Exit Sub
    End If
    Set li = lvwNtrans(Index).ListItems.Add(, , cboAccno(Index))
        li.SubItems(1) = txtParticulars(Index)
        li.SubItems(2) = txtAmountPaid(Index) '"0.00"
        li.SubItems(3) = cboBanks(Index) '"0.00"
        li.SubItems(4) = txtPayee(Index)
      
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub



Private Sub cmdBank_Click(Index As Integer)
    On Error GoTo sysError
    frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        'SearchValue = sel
        If SearchValue <> "" Then
            cboBanks(Index) = SearchValue
            SearchValue = ""
        End If
    End If
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdbookedreceipts_Click()
'//bookedreceipts
    reportname = "bookedreceipts.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub NewVoucher()
    On Error GoTo sysError
    txtVoucherNo(1) = Generate_ReceiptNoT("Voucher", DTPDatedeposited)
    txtReceiptsno(0) = Generate_ReceiptNoT("Receipt", DTPDatedeposited)
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub
Public Function Generate_ReceiptNoT(Choice As String, Transdatee As Date) As String
    On Error GoTo sysError
    Dim rsNos As New Recordset, strmemno As String, lngMemNo As Double, rcount As Double, Rno As String
    Select Case Choice
        Case "Voucher"
         Set rsNos = oSaccoMaster.GetRecordset("select isnull(MAX(RIGHT(ReceiptNo,3)),0)+1 ccount from TellerTrans  where ReceiptNo like '%VNO%' and year(TransDate)='" & year(Transdatee) & "' and month(TransDate)='" & month(Transdatee) & "' and day(TransDate)= '" & year(Transdatee) & "'")
        Case "Receipt"
         Set rsNos = oSaccoMaster.GetRecordset("select isnull(MAX(RIGHT(ReceiptNo,3)),0)+1 ccount from TellerTrans  where ReceiptNo like '%RCP%' and year(TransDate)='" & year(Transdatee) & "' and month(TransDate)='" & month(Transdatee) & "' and day(TransDate)= '" & Day(Transdatee) & "'")
       End Select
       
    With rsNos
    rcount = rsNos.Fields(0)
    Rno = Format(CStr(rcount), "000")
    Generate_ReceiptNoT = IIf(Choice = "Voucher", "VNO", "RCP") & CStr(Day(thisDay)) & CStr(month(thisDay)) & Right(CStr(year(thisDay)), 2) & "-" & CStr(Rno)
    End With
    
    Exit Function
sysError:
    ErrorMessage = err.description
    Generate_ReceiptNoT = ""
End Function

Private Sub cmdprint_Click(Index As Integer)
    On Error GoTo sysError

    If Trim$(IIf(Index = 0, txtReceiptsno(Index), txtVoucherNo(Index))) = "" Then
        MsgBox "Please enter the receipt number", vbInformation, Me.Caption
        Exit Sub
    End If
    Select Case Index
        Case 0
            reportname = "CashReceiptvoucher.rpt"
            STRFORMULA = "{TellerTrans.ReceiptNo}='" & txtReceiptsno(Index) & "'"
        Case 1
            reportname = "CashPaymenttvoucher.rpt"
            STRFORMULA = "{TellerTrans.ReceiptNo}='" & txtVoucherNo(Index) & "'"

        Case Else
    End Select
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
'
'    STRFORMULA = "{paymentBooking.VoucherNo}='" & txtVoucherNo(Index) & "'"
'    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
'    lvwTrans().ListItems.Clear
'    'txtMemberNo = ""
'    txtBalance(0) = "0.00"
'    txtBalance(1) = "0.00"
'    'txtVoucherNo = ""
'    txtAmountPaid(0) = "0.00"
'    txtAmountPaid(1) = "0.00"
'    Exit Sub
'    ''//---------------------printing option-----------------------------------
'    Set rs = oSaccoMaster.GetRecordset("SELECT     CompanyName  FROM         SYSPARAM")
'    Set rst = oSaccoMaster.GetRecordset("SELECT     Memberno, Ccode, Name, Transdate, Amount, Chequeno, Ptype  FROM ReceiptBooking  WHERE     (ReceiptNo = '" & txtVoucherNo(1) & "')")
'    With Adodc1
'        .RecordSource = rst.Source
'        .ConnectionString = cn
'        .Refresh
'    End With
'    Dim pay, tot, disc As Currency
'    Dim Z, x As Integer
'    'for number of copies
'    Dim a As Integer
'    Dim b As Integer
'    Dim c As Integer
'    dlg9.CancelError = True
'    dlg9.FontName = "Garamond"
'    Dim j As Printer
'    a = dlg9.Copies
'    Adodc1.Recordset.MoveFirst
'
'    Printer.CurrentY = 500
'    Printer.CurrentX = 9000
'    Printer.FontSize = 11
'    Printer.CurrentY = 600
'    Printer.CurrentX = 1000
'    Printer.Print Tab(8); "TRANSACTION VOUCHER"
'    Printer.Print Tab(0); "---------------------------------------"
'    Printer.Print Tab(0); rs.Fields(0)
'    Printer.Print Tab(0); "Utumishi HOUSE"
'    Printer.Print Tab(0); "P.o.Box 10454,Tel 020 2733603 "
'    Printer.Print Tab(0); "NAIROBI"
'    Printer.Print Tab(0); "---------------------------------------"
''    Printer.Print Tab(0); "Receipt No"; Tab(10); txtreceiptnoreprint
'    'If optmemberno = True Then
'    Printer.Print Tab(0); "Member No"; Tab(10); rst.Fields(0)
'    Printer.Print Tab(0); "Name"; Tab(10); rst.Fields(2)
''    Else
''    Printer.Print Tab(0); "Company Code"; Tab(10); rst.Fields(1)
''    Printer.Print Tab(0); "Name"; Tab(10); rst.Fields(1)
''    End If
'    Printer.Print
'
'    Printer.CurrentX = 500#
'    Printer.FontSize = 10
'    Printer.CurrentX = 500
'    Printer.FontSize = 8
'    Printer.Print "DATE :  "; Get_Server_Date
'    Printer.CurrentX = 500
'    Printer.FontSize = 8
'    Printer.Print
'    'iNFO TO PRINT
'    Adodc1.Recordset.MoveFirst
'    Printer.CurrentX = 500
'    Printer.FontSize = 8
'    Printer.Print
'    Printer.Print Tab(2); "Item Description"; Tab(30); "AMOUNT"
'    Printer.Print
'    Printer.Print Tab(2); rst.Fields(2); Tab(30); rst.Fields(4)
'    Printer.Print Tab(0); "---------------------------------------"
'    Printer.Print "Amount Received :"; Tab(20); rst.Fields(4)
'
'    'Printer.Print "Your Balance is :"; Tab(20); asa
'    Printer.Print Tab(0); "---------------------------------------"
'    Printer.Print Tab(2); "You were Served by: " & User
'    Printer.Print Tab(2); "Signature"
'    Printer.Print
'    Printer.Print
'    Printer.Print
'    Printer.Print Tab(0); "---------------------------------------"
'    Printer.Print Tab(0); "TRANSFORMING LIVES THROUGH SAVINGS AND CREDIT"
'    Printer.Print
'     'Printer.Print Tab(0); "***********THIS IS A REPRINT***********"
'    Printer.Print
'    Printer.Print Tab(0); "POWERED BY EASYSACCO ENTERPRISE SOLUTIONS"
'    Printer.Print
'    Printer.EndDoc
'    Adodc1.Recordset.MoveNext
'    '//-------------------------
'    'mysql = ""
'    'mysql = "set dateformat dmy Insert into Receiptno(Receiptno,Auditdate,auditid,memberno)values('" & txtvoucherNo(0)& "','" & Format(Get_Server_Date, "dd/MM/yyyy") & "','" & User & "','" & txtMemberno & "')"
'    'oSaccoMaster.ExecuteThis (mysql)
'    lvwTrans.ListItems.Clear
'    MsgBox "You have successfully saved the record", vbInformation
'    txtAmountPaid(0) = 0
'    lblfullnames = ""
'    txtmemberno = ""
'    txtVoucherNo(0) = ""
    Exit Sub
sysError:
    MsgBox err.description, vbInformation
'txtothers = ""
End Sub


'Private Sub cmdRemoveKitu_Click()
' With lvwOtherDeductions
'    If .ListItems.Count > 0 Then
'        OtherDeductions = OtherDeductions - CDbl(.ListItems(.SelectedItem.Index).ListSubItems(3))
'        lvwTrans.ListItems(1).ListSubItems(3).Text = IIf(CDbl(lvwTrans.ListItems(1).ListSubItems(3).Text) > 0, CDbl(lvwTrans.ListItems(1).ListSubItems(3).Text) + CDbl(.ListItems(1).ListSubItems(3)), CDbl(lvwTrans.ListItems(1).ListSubItems(3)) + CDbl(.ListItems(1).ListSubItems(3)))
'        .ListItems.Remove .SelectedItem.Index
'
'    End If
'    txtAmountPaid(0) = 0
'
'
'    txtBalance(0) = CDbl(txtDistributed(0)) * -1
' End With
'End Sub

'Private Sub cmdVoucher_Click(Index As Integer)
'    On Error GoTo SysError
'    frmSearchVoucher.Show vbModal
'    mno = SearchValue
'    'mno1 = SearchValue
'    If mno <> "" Then
'     txtVoucherNo(0).Text = SearchValue
'    txtVoucherNo(1).Text = SearchValue
'        mno = txtVoucherNo(0)
'        If MyRecord <> "" Then
'            txtmemberno = MyRecord
'            MyRecord = ""
'        End If
'    End If
'    Exit Sub
'SysError:
'    MsgBox err.description, vbInformation, Me.Caption
'End Sub

Private Sub cmdRemove_Click(Index As Integer)
    On Error GoTo sysError
    With lvwNtrans(Index)
        If .ListItems.Count > 0 Then
            If MsgBox("Do you want to remove " & .SelectedItem.SubItems(1) & _
            " From the list?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                Exit Sub
            End If
            pushed = pushed - .SelectedItem.ListSubItems(2)
            .ListItems.Remove (.SelectedItem.Index)
        End If
    End With
    Recalculate (Index)
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

'Private Sub cmdupdatereceipt_Click(Index As Integer)
'    On Error GoTo SysError
'    Dim j As Integer
'    Dim crge As Boolean
'    Dim amt As Double, interest As Double
'    Dim code As String, rsLoan As New Recordset
'    Dim k As Long, repaymethod As String
'    Dim shareBal As Double, balance As Double
'    Dim sharebals As Double, Amount As Double
'    Dim ptype As String, Principal As Double, ChequeNo As String
'    Dim postTrans As ADODB.Connection
'    'Recalculate (0)
'
'    If cboBanks(0).Text = "" Then
'        MsgBox "Choose the Receiving Control Account", vbCritical
'        Exit Sub
'    End If
'
'    If txtVoucherNo(0) = "" Then
'        MsgBox "Please generate the receiptno", vbCritical
'        Exit Sub
'    End If
'
'    If txtAmountPaid(0) <= 0 Then
'        MsgBox "Amount should be greater than zero", vbCritical
'        Exit Sub
'    End If
'    If CDbl(txtBalance(0).Text) > 0 Then
'        MsgBox "That Amount is greater than the Remaining Balance, please revise", vbCritical
'        Exit Sub
'    ElseIf CDbl(txtBalance(0).Text) < 0 Then
'        If MsgBox("Is this a part Payments?", vbQuestion + vbYesNo, vbCrLf & "CONFIRMATION") = vbNo Then
'            Exit Sub
'        End If
'    Else
'
'    End If
'    'MsgBox ""
'    If txtAmountPaid(0) = "" Then
'        MsgBox "Amount should be have a figure on it", vbCritical
'        Exit Sub
'    End If
'    If cboMode(Index) = "Cheque" Then
'        If txtmode(Index) = "" Then
'            MsgBox ("VoucherNo or ChequeNo Required"), vbCritical
'            Exit Sub
'        End If
'    End If
'    If cboMode(Index) = "Cash" Then
'        If txtVoucherNo(0) = "" Then
'            MsgBox ("Cash Receipt No Required"), vbInformation
'            Exit Sub
'        End If
'    End If
'    If cboMode(Index) = "EFT" Then
'        If txtmode = "" Then
'            MsgBox ("EFT Receipt No Required"), vbInformation
'            Exit Sub
'        End If
'    End If
'    If cboMode(Index) = "Mpesa" Then
'        If txtmode = "" Then
'            MsgBox ("Mpesa Receipt No Required"), vbInformation
'            Exit Sub
'        End If
'    End If
'    If cboMode(Index) = "Zap" Then
'        If txtmode = "" Then
'            MsgBox ("Zap Receipt No Required"), vbInformation
'            Exit Sub
'        End If
'    End If
'
'
'    sql = "select voucherno from PaymentBooking where voucherno='" & txtVoucherNo(0) & "'"
'    Set rst = oSaccoMaster.GetRecordset(sql)
'    If Not rst.EOF Then
'        MsgBox "That receiptno is already used, Get another One!"
'        Exit Sub
'    End If
'
'    If MsgBox("Do you want to post this Voucher?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
'        Exit Sub
'    End If
'
'
'    Select Case cboMode.Text
'        Case "Cheque"
'        ChequeNo = txtmode
'        Case Else
'        ChequeNo = cboMode.Text
'    End Select
'
'    '/////Check Kama Analibia Vitu Vingine,
'    '/////and also we don't forget to Pay for the refinance if applicable-----Cosmas
'    Dim myItem As String
'
'    'WE START OUR TRANSACTION HERE
'
'    On Error GoTo TransError
'    'oSaccoMaster.goConn.RollbackTrans
'    Set postTrans = New ADODB.Connection
'    postTrans.Open SelectedDsn
'    postTrans.BeginTrans
'
'        'save TransactionNo
'        transactionTotal = CDbl(txtAmountPaid(0).Text)
'        NewTransaction transactionTotal, DTPicker1, "Cheque Issue"
'        If Not SaveTransaction(transactionno, transactionTotal, User, DTPicker1, "Payment Posting - " & txtVoucherNo(0)) Then
'            GoTo TransError
'        End If
'
'        With lvwTrans
'            If .ListItems.Count < 0 Then
'                Exit Sub
'            End If
'            j = 0
'            i = .ListItems.Count
'            For j = 1 To i
'                code = .ListItems(j).ListSubItems(7)
'                If code = "Shares" Then
'                    amt = CDbl(txtAmountPaid(0).Text)
'                    sharebals = CDbl(txtBalance(0).Text) * (-1)
'                    ContraAcc = cboBanks(0).Text
'                    If Not SaveContrib(mMemberNo, DTPDatedeposited.value, lvwTrans.ListItems(j).SubItems(1), CDbl(amt) * (-1), cboBanks(0).Text, ChequeNo, ChequeNo, User, "Share Withdrawal", transactionno) Then
'                        GoTo TransError
'                    End If
'
'                    sql = ""
'                    sql = "set dateformat dmy INSERT INTO PaymentBooking (VoucherNo,Memberno,Ccode,Name,Transdate," _
'                    & "Amount, Chequeno, Ptype, auditid,datedeposited,draccno,craccno,PayeeDesc,Transactionno) VALUES ('" & txtVoucherNo(0) & "','" & _
'                    txtmemberno & "','" & cboCompany & "','" & .ListItems(j).ListSubItems(1) & "--" & ptype _
'                    & "','" & DTPDatedeposited & "'," & amt & ",'" & _
'                    txtmode.Text & "','" & ptype & "','" & User & "','" & DTPDatedeposited & "','" & shareAcc & "','" & ContraAcc & "','" & lblfullnames.Caption & "','" & transactionno & "')"
'                    oSaccoMaster.ExecuteThis (sql)
'
'                    oSaccoMaster.ExecuteThis ("Update shareWithdrawal set withdrawn=1 where memberno='" & txtmemberno & "' and sharescode='" & lvwTrans.ListItems(j).SubItems(1) & "' and withdrawn=0")
'                    If Not success Then GoTo TransError
'                ElseIf code = "Loan" Then
'                    LoanAmount = .ListItems(j).ListSubItems(2)
'                    amt = txtAmountPaid(0) 'IIf(.ListItems(j).ListSubItems(3) > 0, .ListItems(j).ListSubItems(3), .ListItems(j).ListSubItems(2))
'                    balance = CDbl(txtBalance(0).Text) * (-1) 'amt - CDbl(txtAmountPaid(0))
'                    LoanNo = .ListItems(j).Text
'
'                    If Not EffectLoanPayment(LoanNo, txtVoucherNo(0).Text, amt, balance, cboBanks(0), DTPDatedeposited, User, transactionno, txtmode) = True Then
'                        GoTo TransError
'                        Exit Sub
'                    Else
'                         Set rst = oSaccoMaster.GetRecordset("select loanAcc from loantype lt inner join loans l on l.loancode=lt.loancode where l.loanno='" & LoanNo & "'")
'                        If Not rst.EOF Then
'                            LoanAcc = rst(0)
'                            ContraAcc = cboBanks(0).Text
'                        End If
'
'                        'If there was charges, pay
'                        Set rs2 = oSaccoMaster.GetRecordset("Select  DISTINCT dd.* from disbursementdeduction dd inner join bridgingloan bl on dd.loanno=bl.loanno where dd.loanno='" & LoanNo & "' and bl.paid=0 and dd.rloanno='FEE'")
'                        'rst.Open "Select * from DisbursementDeduction where loanno='Charge'", oSaccoMaster.goConn, adOpenKeyset, adLockReadOnly
'                        While Not rs2.EOF 'He/She is
'                            CRaccno = Trim(rs2!Accno)
'                            If Save_GLTRANSACTION(dateissued, rs2("Amount"), LoanAcc, CRaccno, txtVoucherNo(0).Text _
'                            , txtmemberno, User, "", "Loan Disbursement Deduction", 0, 1, txtmode.Text, transactionno, "BOSA") = False Then
'                                MsgBox ErrorMessage, vbOKOnly + vbInformation
'                                GoTo TransError
'                            End If
'                            rs2.MoveNext
'                        Wend
'
'                        If Not oSaccoMaster.Execute("Update loans set status=4 where loanno='" & LoanNo & "'") Then
'                            GoTo TransError
'                        End If
'
'
'
'                        sql = ""
'                        sql = "set dateformat dmy INSERT INTO PaymentBooking (VoucherNo,Memberno,PayeeDesc,Ccode,Name,Transdate," _
'                        & "Amount, Chequeno, Ptype, auditid,datedeposited,draccno,craccno) VALUES ('" & txtVoucherNo(0) & "','" & _
'                        txtmemberno & "','" & lblfullnames.Caption & "','" & cboCompany & "','" & .ListItems(j).ListSubItems(1) & "--" & cboMode.Text _
'                        & "','" & DTPDatedeposited & "'," & amt & ",'" & _
'                        txtmode.Text & "','" & cboMode.Text & "','" & User & "','" & DTPDatedeposited & "','" & LoanAcc & "','" & ContraAcc & "')"
'                        oSaccoMaster.ExecuteThis (sql)
'                        If success = True Then
'
'                        Else
'                            GoTo TransError
'                        End If
'
'                        'if this was a refinance loan, release the older loan
'                        Dim myloanno As String
'                        myloanno = ""
'                        Set RsLoans = oSaccoMaster.GetRecordset("select dd.rloanno,dd.amount from disbursementdeduction dd inner join bridgingloan bl on dd.rloanno=bl.brgloanno where dd.loanno='" & LoanNo & "' and dd.description='LR' and bl.paid=0")
'                        If Not RsLoans.EOF Then
'                            While Not RsLoans.EOF
'                                If myloanno = "" Then
'                                    myloanno = RsLoans!RLOANNO
'                                Else
'                                    myloanno = myloanno & "','" & RsLoans!RLOANNO
'                                End If
'                                'get the loanacc for the refinancing loan to use as the paying account
'
'                                Set rst = oSaccoMaster.GetRecordset("select lt.loanacc from loantype lt inner join loans l on l.loancode=lt.loancode where l.loanno='" & LoanNo & "'")
'                                If Not rst.EOF Then
'                                    BankAcc = rst!LoanAcc
'                                End If
'
'                                crge = True
''                                If RsLoans!interest > 0 Then
''                                    crge = True
''                                Else
''                                    crge = False
''                                End If
'                                If Not SaveRepay(RsLoans!RLOANNO, DTPDatedeposited, RsLoans!Amount, BankAcc, txtVoucherNo(0), 0, 1, "Loan Refinanced ", User, "Refinancing", transactionno, "", False, crge) Then
'                                    GoTo TransError
'                                Else
'                                End If
'                            RsLoans.MoveNext
'                            Wend
'
'                            If Not oSaccoMaster.Execute("Update bridgingloan set paid=1 where brgloanno in ('" & myloanno & "')") Then
'                                GoTo TransError
'                            End If
'                        Else
'
'                        End If
'                    End If
'                ElseIf code = "LoanRefund" Then
'                    LoanAmount = .ListItems(j).ListSubItems(2)
'                    amt = txtAmountPaid(0) 'IIf(.ListItems(j).ListSubItems(3) > 0, .ListItems(j).ListSubItems(3), .ListItems(j).ListSubItems(2))
'                    balance = CDbl(txtBalance(0).Text) * (-1) 'amt - CDbl(txtAmountPaid(0))
'                    LoanNo = .ListItems(j).Text
'
'                    If Not SaveRepay(LoanNo, DTPDatedeposited.value, LoanAmount * (-1), cboBanks(0), txtVoucherNo(0), 0, 0, "Loan Refund", User, "Refund", transactionno, , False) Then
'                        GoTo TransError
'                    End If
'
'                    sql = ""
'                    sql = "set dateformat dmy INSERT INTO PaymentBooking (VoucherNo,Memberno,Ccode,Name,Transdate," _
'                    & "Amount, Chequeno, Ptype, auditid,datedeposited,draccno,craccno,PayeeDesc) VALUES ('" & txtVoucherNo(0) & "','" & _
'                    txtmemberno & "','" & cboCompany & "','" & .ListItems(j).ListSubItems(1) & "--" & ptype _
'                    & "','" & DTPDatedeposited & "'," & amt & ",'" & _
'                    txtmode.Text & "','" & ptype & "','" & User & "','" & DTPDatedeposited & "','" & LoanAcc & "','" & ContraAcc & "','" & lblfullnames.Caption & "')"
'                    oSaccoMaster.ExecuteThis (sql)
'                    If success = True Then
'                        oSaccoMaster.ExecuteThis ("Update Refunds set done=1 where refcode='" & LoanNo & "' and done=0")
'                    Else
'                        GoTo TransError
'                    End If
'                Else
'                    ErrorMessage = "The item under this transaction is not enlisted in the system. Therefore, transactions never saved"
'                    GoTo TransError
'                End If
'            Next j
'        End With
'    'So, if there was no error, Commit the transaction
'    postTrans.CommitTrans
'
'    lvwTrans.ListItems.Clear
'    PendingPayments
'    If MsgBox("Voucher Posted successfully. Print Voucher?", vbYesNo + vbQuestion) = vbYes Then
'        cmdPrint_Click (0)
'    End If
'    NewVoucher
'    Exit Sub
'SysError:
'    If ErrorMessage = "" Then ErrorMessage = err.description
'    MsgBox ErrorMessage, vbInformation, Me.Caption
'    Exit Sub
'TransError:
'    If ErrorMessage = "" Then ErrorMessage = err.description
'    MsgBox ErrorMessage, vbInformation
'    postTrans.RollbackTrans
'End Sub



Private Sub cmdupdatereceipt_Click(Index As Integer)
    Dim balance As Double, Cr As Double, Dr As Double, TransSource As String, TransDescription As String
    Dim TransPPosting As ADODB.Connection, AccNo As String, chequeno As String
    If txtBalance(1).Text <> 0 Then
        MsgBox "This payment is not fully allocated,please Check", vbCritical
        Exit Sub
    End If
    If cboMode(Index) = "" Then
        MsgBox "Please select the payment Mode", vbInformation, Me.Caption
        Exit Sub
    End If


    NewVoucher
    
    Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & DTPDatedeposited & "'")
    If Not rs.EOF Then
        MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
        Exit Sub
    End If
    
    If MsgBox("Post The transaction?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    
    
    If Index = 1 Then
'    If MsgBox("Authorize the Reversal?", vbQuestion + vbOKCancel) = vbOK Then
'        frmAuthorize.Show vbModal
'        If user = Authority Or (UCase(Authority) = "ADMIN") Or Authority = "" Then
'            MsgBox "Transaction denied/differed", vbOKOnly + vbExclamation
'            Exit Sub
'        End If
'    Else
'        Exit Sub
'    End If
    End If
    If lvwNtrans(Index).ListItems.Count > 0 Then
    'save TransactionNo
    GetTransactionNo
    transactionTotal = CDbl(txtAmountPaid(Index).Text)

    Set TransPPosting = New ADODB.Connection
    TransPPosting.Open oSaccoMaster.goConn

    On Error GoTo TransError
    TransPPosting.BeginTrans
    
        If Not SaveTransaction(transactionNo, transactionTotal, User, DTPDatedeposited, IIf(Index = 0, "Receipt Posting -RCP", "Payment Posting -VNo") & IIf(Index = 0, txtReceiptsno(Index), txtVoucherNo(Index))) Then
            GoTo TransError
        End If

        With lvwNtrans(Index)
            If .ListItems.Count < 1 Then
            Exit Sub
            End If

            For I = 1 To .ListItems.Count
            
                If Index = 0 Then
                    DRaccno = cboBanks(Index)
                    Craccno = lvwNtrans(Index).ListItems(I)
                    AccNo = DRaccno
                Else
                    DRaccno = lvwNtrans(Index).ListItems(I)
                    Craccno = cboBanks(Index)
                    AccNo = Craccno
                End If
              
              Dr = IIf(Index = 0, CDbl(.ListItems(I).ListSubItems(2).Text), 0)
              Cr = IIf(Index = 1, CDbl(.ListItems(I).ListSubItems(2).Text), 0)
              TransDescription = .ListItems(I).ListSubItems(1).Text
              TransSource = .ListItems(I).ListSubItems(4).Text
              chequeno = txtmode(Index)
'            sql = ""
'            sql = "set dateformat dmy INSERT INTO Payments (VoucherNo,RefNo,Transdate," _
'            & "Amount, Chequeno, Ptype, auditid,particulars,payee) VALUES" _
'            & " ('" & IIf(Index = 0, txtReceiptsno(Index), txtVoucherNo(Index)) & "','" & IIf(Index = 0, txtReceiptsno(Index), txtVoucherNo(Index)) & "','" & DTPDatedeposited & "'," & CDbl(.ListItems(I).ListSubItems(2).Text) & ",'" & _
'            IIf(Index = 0, txtReceiptsno(Index), txtVoucherNo(Index)) & "','" & cboMode(Index) & "','" & user & "','" & txtParticulars(Index) & "','" & txtPayee(Index) & "')"
'            oSaccoMaster.ExecuteThis (sql)
        
            sql = "Insert into tellertrans (transdate,AccNo,userloginid,DR,CR,Receiptno,Description,Auditid,refno,TransName,TransactionNo)"
            sql = sql & " Values('" & Format(DTPDatedeposited, "DD/MM/YYYY") & "','" & AccNo & "','" & User & "'," & Dr & "," & Cr & ",'" & IIf(Index = 0, txtReceiptsno(Index), txtVoucherNo(Index)) & "','" & TransDescription & "','" & User & "','" & cboMode(Index) & "','" & TransSource & "','" & transactionNo & "')"
            oSaccoMaster.ExecuteThis (sql)

'            If Not Save_GLTRANSACTION(Format(DTPDatedeposited, "DD/MM/YYYY"), lvwNtrans(Index).ListItems(I).SubItems(2), DRaccno, Craccno, _
'            IIf(Index = 0, txtReceiptsno(Index), txtVoucherNo(Index)), "Source", User, "", txtParticulars(Index), 0, txtPayee(Index), 0, transactionNo, "") = False Then
'                MsgBox ErrorMessage, vbOKOnly + vbInformation
'                Exit Sub
'            End If
            
            If Not Save_GLTRANSACTION(Format(DTPDatedeposited, "DD/MM/YYYY"), lvwNtrans(Index).ListItems(I).SubItems(2), DRaccno, Craccno, _
            IIf(Index = 0, txtReceiptsno(Index), txtVoucherNo(Index)), TransSource, User, "", TransDescription, 0, 0, txtmode(Index), transactionNo, "") Then
              GoTo TransError
            End If

            Next I

        End With

        TransPPosting.CommitTrans

        If MsgBox("Voucher Posted successfully. Print Voucher?", vbYesNo + vbQuestion) = vbYes Then
            cmdprint_Click (Index)
        End If
        lvwNtrans(Index).ListItems.Clear
        txtDistributed(Index).Text = 0
        txtAmountPaid(Index).Text = 0
        txtBalance(Index) = ""
        txtBalance(Index).SetFocus
         txtParticulars(Index).Text = ""
         txtPayee(Index).Text = ""
        NewVoucher
    End If
    Exit Sub
TransError:
    MsgBox err.description
    TransPPosting.RollbackTrans
End Sub

Private Sub Recalculate(Index As Integer)
    txtBalance(Index).Text = 0
    balance = 0
    If lvwNtrans(Index).ListItems.Count > 0 Then
        For I = 1 To lvwNtrans(Index).ListItems.Count
            balance = balance + CDbl(lvwNtrans(Index).ListItems(I).SubItems(2))
        Next I
    End If
    txtDistributed(Index) = Format(balance, Cfmt)
    txtBalance(Index) = Format(CDbl(txtAmountPaid(Index)) - CDbl(txtDistributed(Index)), Cfmt)

End Sub



Private Sub Form_Load()
    'dtpTransDate = Format(Get_Server_Date, "dd/mm/yyyy")
    DTPDatedeposited = Format(Get_Server_Date, "dd/mm/yyyy")
    thisDay = Format(Get_Server_Date, "dd/mm/yyyy")
    'Load Gl's
    sql = "Select accno from glsetup order by accno asc"
    Set Rst = oSaccoMaster.GetRecordset(sql)
    While Not Rst.EOF
        cboAccno(0).AddItem (Rst(0))
        cboAccno(1).AddItem (Rst(0))
        Rst.MoveNext
    Wend
    'load the banks
    cboBanks(0).Clear
    cboBanks(1).Clear
    sql = "select AssignGl from UserAccounts where AssignGl in (Select accno from glsetup )"
    Set Rst = oSaccoMaster.GetRecordset(sql)
    While Not Rst.EOF
        cboBanks(0).AddItem Rst(0)
        cboBanks(1).AddItem Rst(0)
        Rst.MoveNext
    Wend
   
    cboBanks(0).List(0) = current_user.tellerGlAcc
    cboBanks(0).Text = cboBanks(0).List(0)
    cboBanks(1).List(0) = current_user.tellerGlAcc
    cboBanks(1).Text = cboBanks(1).List(0)
'    For I = 0 To 1
'        cboBanks_Change (I)
'    Next I
    
    
    'initialization
    totalamount = 0
    pushed = 0
    
'    InitSubClass
'    'Enable label editing for listview2
'    Set objLabelEdit = New LabelEdit
'    objLabelEdit.Init Me, lvwNtrans(0)
'    Set objLabelEdit2 = New LabelEdit
'    objLabelEdit2.Init Me, lvwNtrans(1)
'   InitSubClass
' Set objLabelEdit = New LabelEdit
'    objLabelEdit.Init Me, lvwNtrans(0)
'    Set objLabelEdit2 = New LabelEdit
'    objLabelEdit2.Init Me, lvwNtrans(1)
    
    NewVoucher

End Sub



Private Sub Form_Unload(Cancel As Integer)
    'Stop subclassing
    CloseSubClass
    'Clean up by setting the classes to Nothing
    Set objLabelEdit = Nothing
    'Set objLabelEdit2 = Nothing
End Sub



Private Sub lvwNtrans_Click(Index As Integer)
    Dim Total As Double
    Dim ccount As Integer
    With lvwNtrans(Index)
        If .ListItems.Count > 0 Then
            ccount = .ListItems.Count
            For I = 1 To ccount
                With .ListItems(I)
                        amt = .ListSubItems(2)
                        Total = Total + amt
                End With
            Next I

        Else
            Total = 0
        End If
        TotalBalance = Total
    End With
    txtDistributed(Index).Text = Format(Total, Cfmt)
    txtAmountPaid_Change (Index)
End Sub

Private Sub lvwNtrans_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    With lvwNtrans(Index)
        Dim samount As Double
        
        If .ListItems.Count = 0 Then
            Exit Sub
        End If

        If KeyAscii = 13 Then
            samount = InputBox("Enter the Amount ", 0) '
            If Not IsNumeric(samount) Then
                samount = Format(0, Cfmt)
            End If
            .ListItems(.SelectedItem.Index).SubItems(2) = samount
            Recalculate (Index)
        End If
    End With
End Sub



Private Sub txtAmountPaid_Change(Index As Integer)
On Error Resume Next
    If txtAmountPaid(Index).Text = "" Or IsNull(txtAmountPaid(Index)) Then
        txtAmountPaid(Index) = Format(0, Cfmt)
    End If


    If CDbl(txtAmountPaid(Index)) > CDbl(TotalBalance) Then
        'MsgBox "You should not exceed this Balance (" & TotalBalance & ")", vbCritical
        txtAmountPaid_KeyPress 0, 8
        'Exit Sub
        'KeyAscii = 0
    End If


    If txtAmountPaid(Index).Text = "" Then txtAmountPaid(Index).Text = 0
    totalamount = CDbl(txtAmountPaid(Index).Text)
    pushed = 0
    txtBalance(Index).Text = Format((totalamount) - CDbl(txtDistributed(Index).Text), Cfmt)
    'txtAmountPaid(index).Text = CDbl(txtBalance(index).Text)

End Sub


Private Sub txtAmountPaid_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii > 47 And KeyAscii < 58 Then

    ElseIf KeyAscii = 8 Then

    ElseIf KeyAscii = 46 Then

    Else
        Beep
        KeyAscii = 0
    End If

End Sub

Private Sub txtAmountPaid_LostFocus(Index As Integer)
    txtAmountPaid(Index) = Format(txtAmountPaid(Index), Cfmt)
End Sub


