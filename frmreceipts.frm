VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmreceipts 
   BackColor       =   &H00FF8080&
   Caption         =   "SALES"
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmreceipts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   11295
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cbocategory 
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
      ItemData        =   "frmreceipts.frx":0442
      Left            =   8400
      List            =   "frmreceipts.frx":044F
      TabIndex        =   62
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox txtcomment 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   2160
      TabIndex        =   60
      Top             =   8880
      Width           =   6255
   End
   Begin VB.CommandButton cmdSearchMember 
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
      Left            =   9240
      Picture         =   "frmreceipts.frx":0479
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   4200
      Width           =   315
   End
   Begin VB.TextBox txtstaffname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   58
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CommandButton cmdsalesreturn 
      Caption         =   "Sales Return"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   57
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox TXTTOTAL 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   56
      Text            =   "0"
      Top             =   6120
      Width           =   2535
   End
   Begin VB.TextBox TXTCHANGE 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   54
      Text            =   "0"
      Top             =   8280
      Width           =   2415
   End
   Begin VB.TextBox txtamtreceived 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   52
      Text            =   "0"
      Top             =   7080
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   50
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Add New Product"
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
      Left            =   1800
      TabIndex        =   49
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdsagroded 
      Caption         =   "Staff Agrovet Deductions"
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
      Left            =   120
      TabIndex        =   48
      Top             =   7320
      Width           =   1935
   End
   Begin VB.TextBox txtstaffno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7920
      TabIndex        =   47
      Top             =   4200
      Width           =   1095
   End
   Begin VB.OptionButton optstaff 
      Caption         =   "Credit"
      Height          =   255
      Left            =   6840
      TabIndex        =   45
      Top             =   2040
      Width           =   1335
   End
   Begin VB.OptionButton Optbranch 
      Caption         =   "Dispatch To Station"
      Height          =   375
      Left            =   6840
      TabIndex        =   44
      Top             =   1320
      Width           =   1455
   End
   Begin VB.ComboBox Cmbstation 
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
      ItemData        =   "frmreceipts.frx":057B
      Left            =   8400
      List            =   "frmreceipts.frx":057D
      TabIndex        =   43
      Top             =   1320
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTPto 
      Height          =   255
      Left            =   9480
      TabIndex        =   41
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   146472961
      CurrentDate     =   40588
   End
   Begin VB.TextBox txttranscode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      TabIndex        =   38
      Top             =   4800
      Width           =   1335
   End
   Begin VB.OptionButton opttransport 
      Caption         =   "Transporters"
      Height          =   255
      Left            =   6840
      TabIndex        =   35
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CheckBox chkPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Print Receipt"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   34
      Top             =   4560
      Value           =   2  'Grayed
      Width           =   1935
   End
   Begin VB.ComboBox cboproductname 
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
      Left            =   1680
      TabIndex        =   33
      Top             =   1200
      Width           =   4215
   End
   Begin VB.TextBox txtSNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   23
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton lblCheckOff 
      Caption         =   "Check Off"
      Height          =   255
      Left            =   1560
      TabIndex        =   21
      Top             =   3360
      Width           =   1335
   End
   Begin VB.OptionButton optCash 
      Caption         =   "Cash"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4200
      Picture         =   "frmreceipts.frx":057F
      ScaleHeight     =   360
      ScaleWidth      =   240
      TabIndex        =   16
      Top             =   720
      Width           =   240
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4200
      Picture         =   "frmreceipts.frx":0701
      ScaleHeight     =   360
      ScaleWidth      =   240
      TabIndex        =   15
      Top             =   240
      Width           =   240
   End
   Begin VB.CommandButton cmdnextitem 
      Caption         =   "Next item"
      Default         =   -1  'True
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
      Left            =   4080
      TabIndex        =   14
      Top             =   2640
      Width           =   1455
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
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
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
      Left            =   120
      TabIndex        =   4
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
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
      Left            =   120
      TabIndex        =   12
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox txtpcode 
      Appearance      =   0  'Flat
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
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox txtquantity 
      Appearance      =   0  'Flat
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
      Left            =   1680
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtamount 
      Appearance      =   0  'Flat
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
      Left            =   1680
      TabIndex        =   3
      Top             =   2640
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker txtransdate 
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   146472961
      CurrentDate     =   40265
   End
   Begin VB.TextBox txtrno 
      Appearance      =   0  'Flat
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
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin MSComctlLib.ListView Lvwitems 
      Height          =   3255
      Left            =   2160
      TabIndex        =   19
      Top             =   5520
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5741
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   4
      MouseIcon       =   "frmreceipts.frx":0883
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ITEM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "QNTY"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "PRICE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "AMOUNT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cash"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlg9 
      Left            =   10560
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label12 
      Caption         =   "Comment"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   61
      Top             =   9000
      Width           =   855
   End
   Begin VB.Label Label18 
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   55
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label17 
      Caption         =   "CHANGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8400
      TabIndex        =   53
      Top             =   7920
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "AMOUNT RECEIVED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8400
      TabIndex        =   51
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label14 
      Caption         =   "Staff No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   46
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label16 
      Caption         =   "Period Ending"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   42
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lbltransnetpay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      TabIndex        =   40
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Net Pay:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   39
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label lbltransportername 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   37
      Top             =   4800
      Width           =   165
   End
   Begin VB.Label Label5 
      Caption         =   "Transport Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   36
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblSNames 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   4680
      TabIndex        =   32
      Top             =   3360
      Width           =   60
   End
   Begin VB.Label Label13 
      Caption         =   "Total Kgs :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblTKgs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   30
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Gross Pay:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   29
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblGPay 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3480
      TabIndex        =   28
      Top             =   3720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Deductions :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   27
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblDed 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7320
      TabIndex        =   26
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblNPay 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4200
      TabIndex        =   25
      Top             =   4200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblNetPay 
      BackColor       =   &H0000FF00&
      Caption         =   "NetPay:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   24
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblSNo 
      Caption         =   "SNo :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   22
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   18
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblbalance 
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Product Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Trans_Date"
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Product Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Receipt No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmreceipts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Provider As String
Dim SelectedDsn As String
Dim DIA
Dim Amount As Double
Dim Total As Currency

Private Sub cboproductname_Change()
        sql = "select p_code, S_No, Qout, sprice from ag_products where p_name='" & cboproductname & "'"
       
        Set rs = oSaccoMaster.GetRecordset(sql)
        If Not rs.EOF Then
            txtpcode = rs.Fields(0)
            lblbalance = IIf(IsNull(rs.Fields(2)), 0, rs.Fields(2))
            txtamount = IIf(IsNull(rs.Fields(3)), 0, rs.Fields(3))
        
        End If
 Exit Sub
End Sub

Private Sub cboproductname_Click()
cboproductname_Change
 Exit Sub
End Sub

Private Sub cboproductname_KeyPress(KeyAscii As Integer)
'KeyAscii = 0
'cboproductname_Validate (True)
cboproductname_Change

End Sub

Private Sub cboproductname_Validate(Cancel As Boolean)
cmdnew_Click

Provider = cn
Set cn = New ADODB.Connection
Dim p As Integer
'cn.Open Provider, "bi"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
'Dim rst As New ADODB.Recordset
sql = ""
'SELECT p_code, p_name, S_No, Qout, sprice FROM   ag_Products
sql = "select p_code, S_No, Qout, sprice from ag_products where p_name='" & cboproductname & "'"
Set rs = New ADODB.Recordset
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
txtpcode = rs.Fields(0)
lblbalance = rs.Fields(2)
'txtserialno = rs.Fields(1)
txtamount = rs.Fields(3)

End If

End Sub


Private Sub Cmbstation_Change()
lblCheckOff.Visible = False
lblSNo.Visible = False
txtSNo.Visible = False
Label13.Visible = False
lblTKgs.Visible = False
Label11.Visible = False
lblGPay.Visible = False
Label8.Visible = False
lblDed.Visible = False
lblNetPay.Visible = False
lblNPay.Visible = False
Label5.Visible = False
txttranscode.Visible = False
Label10.Visible = False
lbltransnetpay.Visible = False


End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdnew_Click()
Set rs = oSaccoMaster.GetRecordset("d_sp_NextReceipt")
If Not (rs.EOF) Then
txtrno = rs.Fields(0) + 1
Else
txtrno = 1
End If

 txtpcode = ""
 'txtserialno = ""
 txtquantity = 1
 txtamount = 0
 txtamtreceived = 0
 TXTCHANGE = 0
 TXTTOTAL = 0
End Sub

Private Sub cmdnextitem_Click()
Dim cash As Integer
Dim Total As Double
    If Trim(txtquantity) = "" Then
        MsgBox "Quantity cannot be Zero", vbInformation
        Exit Sub
    End If
    
    If CDbl(txtquantity) > CDbl(lblbalance) Then
        MsgBox "You Cannot Sale more than Stock Balance", vbInformation, Me.Caption
        Exit Sub
    End If

    If Optbranch = True Then
        If Trim(Cmbstation.Text) = "" Then
            MsgBox "Please enter the Agrovet Station."
                Cmbstation.SetFocus
        Exit Sub
    End If
    End If
    
    
    
    If opttransport = True Then
        If Trim(txttranscode) = "" Then
            MsgBox "Please enter the Transporter"
        
        Exit Sub
        End If
    End If
    
    If txtpcode = "" Then
        MsgBox "Please Enter the Product CODE before You Proceed!", vbCritical
        Exit Sub
    End If
    If txtrno = "" Then
        MsgBox "Please Enter Receipt Number before you Proceed!", vbCritical
        Exit Sub
    End If
    
If txtamount = "" Or txtamount = 0 Then
 MsgBox "Update The Product Selling Price first", vbInformation
  Exit Sub
End If
'Provider = "maziwa"
'Set cn = New ADODB.Connection
'cn.Open Provider, "bi"
'// check if they are in stock.
Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qout from ag_products where p_code='" & txtpcode & "'"
Set rsinstock = oSaccoMaster.GetRecordset(sql)

'// check the stock if it is less than zero
If rsinstock.Fields(1) <= 0 Then
    MsgBox "Sorry Stock is Zero please re-stock before your proceed", vbInformation
    Exit Sub
End If
'// check the quanttity being sold versus the balance
Dim piu As Double
piu = rsinstock.Fields(1) - CInt(txtquantity)

'If piu < 0 Then
'MsgBox "Stock will be negative please re-stock before you proceed", vbInformation
'Exit Sub
'End If

If optCash.Value = True Then
cash = 1
Else
cash = 0
End If

Dim j, Coun As Integer
j = 1




'Check if same item is in the list
   Do While Not j > (Coun)
         Lvwitems.ListItems.Item(j).selected = True
            
    If Lvwitems.SelectedItem = txtpcode Then
        txtquantity = (CCur(txtquantity) + CCur(Lvwitems.SelectedItem.ListSubItems(2)))
        Lvwitems.ListItems.Remove (Lvwitems.SelectedItem.Index)
                        
        Set li = Lvwitems.ListItems.Add(, , txtpcode)
                        li.SubItems(1) = cboproductname & ""
                        li.SubItems(2) = txtquantity & ""
                        li.SubItems(3) = txtamount & ""
                        li.SubItems(4) = CCur(txtamount) * CCur(txtquantity) & ""
                        li.SubItems(5) = cash
                        'Total = CCur(Total + li.SubItems(4))
                        TXTTOTAL = Total
                                                
        j = Coun + 1
        
        lblbalance = CCur(lblbalance) - CCur(txtquantity)

        txtpcode = ""
        txtquantity = ""
       ' txtserialno = ""
        txtpcode.SetFocus
        Exit Sub
         
    
   
'   lvwItems.ListItems.Item(J).selected = True
   End If
   j = j + 1
    Loop
    
     If j > 1 Then
   
    Set li = Lvwitems.ListItems.Add(, , txtpcode)
                        li.SubItems(1) = cboproductname & ""
                        li.SubItems(2) = txtquantity & ""
                        li.SubItems(3) = txtamount & ""
                        li.SubItems(4) = CCur(txtamount) * (CCur(txtquantity)) & ""
                        li.SubItems(5) = cash
                        'Total = CCur(Total + li.SubItems(4))
                        TXTTOTAL = Total
                        
        lblbalance = CCur(lblbalance) - CCur(txtquantity)
        txtpcode = ""
        txtquantity = ""
        'txtserialno = ""
        txtpcode.SetFocus
        Exit Sub
    End If
     If Coun = 0 Then
     Set li = Lvwitems.ListItems.Add(, , txtpcode)
                        li.SubItems(1) = cboproductname & ""
                        li.SubItems(2) = txtquantity & ""
                        li.SubItems(3) = txtamount & ""
                        li.SubItems(4) = CCur(txtamount) * (CCur(txtquantity)) & ""
                        li.SubItems(5) = cash
                        'Total = CCur(Total + li.SubItems(4))
                        TXTTOTAL = Total
    End If

lblbalance = CCur(lblbalance) - CCur(txtquantity)
TXTTOTAL = 0
'Coun = Lvwitems.ListItems.Count
'For j = 1 To Lvwitems.ListItems.Count
'    Total = CCur(Total + li.SubItems(4))
'    txttotal = Total
'
'Next j
Do While Not j > (Lvwitems.ListItems.Count)
'For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True
 Total = Total + CCur(Lvwitems.SelectedItem.SubItems(4))
 TXTTOTAL = Total
j = j + 1
Loop

txtpcode = ""
txtquantity = ""
'txtserialno = ""
txtpcode.SetFocus
End Sub

Private Sub cmdsagroded_Click()
'//staffagrovetdeductions
'//d_payroll\
'//call the companyname

 reportname = "staffagrovetdeductions.rpt"

 
 Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub cmdsalesreturn_Click()
Unload Me
frmsalesreturn.Show vbModal

End Sub

Private Sub cmdSave_Click()
On Error GoTo HEREEE

If opttransport = True Then
savetransporters
Exit Sub
End If

If Optbranch = True Then
savestation
Exit Sub
End If


If lblCheckOff = True Then
   If txtSNo = "" Then
    MsgBox "Enter the SupplierNo  ", vbInformation, "CheckOff"
     Exit Sub
    End If
savesno
Exit Sub
End If

If optCash = True Then
    If TXTCHANGE < 0 Then
        If MsgBox("Insufficient Amount Received,do you want to transfer balance to check off ", vbYesNo) = vbYes Then
            lblCheckOff_Click
            lblCheckOff.Value = True
            optCash.Value = False
           Exit Sub
        Else
           Exit Sub
         End If
    End If
    savecash
   Exit Sub
End If
If optstaff = True Then
savestaff
Exit Sub
End If
HEREEE:
MsgBox err.description & " error occured."

End Sub

Private Sub savesno()
On Error GoTo ebraim

If lblCheckOff = True Then

Set rs = New ADODB.Recordset
sql = "d_sp_SelectSuppliers '" & txtSNo & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(2)) Then lblSNames.Caption = rs.Fields(2)
'load_deduc
Else
lblSNames.Caption = ""
End If

If Trim(lblSNames.Caption) = "" Then
MsgBox "Please enter a valid supplier number."
txtSNo.SetFocus
Exit Sub
End If
Dim a, b, X
DIA = 0
Dim U As Double, S As Double
Dim cn As Connection
Dim j As Integer
'txtserialno_LostFocus
'If DIA = 1 Then
'Exit Sub
'End If
If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items sold."
Exit Sub
End If
j = 1


Total = 0
Do While Not j > (Lvwitems.ListItems.Count)
'For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True
 Total = Total + CCur(Lvwitems.SelectedItem.SubItems(4))
j = j + 1
Loop

If optCash.Value = False Then

Startdate = DateSerial(year(txtransdate), month(txtransdate), 1)
Enddate = DateSerial(year(txtransdate), month(txtransdate) + 1, 1 - 1)


Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If
'End If
End If
j = 1
For j = 1 To Lvwitems.ListItems.Count
'Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
If Trim(txtSNo) = "" Then
txtSNo = "0"
End If
'// check if they are in stock.

Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qout from ag_products where p_code='" & Lvwitems.SelectedItem & "'"

Set rsinstock = oSaccoMaster.GetRecordset(sql)

Dim Remain As Double
Remain = rsinstock.Fields(1) - CInt(Lvwitems.SelectedItem.SubItems(2))


sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash,S_No, SNo,Remarks,Description) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & txtSNo & "','" & txtSNo & "','Checkoff','" & txtcomment & "')"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "'")
'//XXXXXXXXXXXXXXX
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'A047','I013','" & Lvwitems.SelectedItem & "','" & Lvwitems.SelectedItem.SubItems(1) & "' ,' CHECK OFF SALES ','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)




'XXXXXXXXXXXXXXXXXXXXXX
Next j
'j = j + 1
'Loop

If optCash.Value = False Then
Set cn = New ADODB.Connection

sql = "d_sp_SupplierDeduct '" & txtSNo & "','" & txtransdate & "','Agrovet'," & Total & ",'" & Startdate & "','" & Enddate & "'," & year(txtransdate) & ",'" & User & "','Receipt " & txtrno & "'"
oSaccoMaster.ExecuteThis (sql)
End If

If CDbl(txtamtreceived) > 0 Then
    '******Deduct Amount paid in cash
   
    Amount = 0
    Amount = CDbl(txtamtreceived)
    sql = "d_sp_SupplierDeduct '" & txtSNo & "','" & txtransdate & "','Agrovet'," & Amount * -1 & ",'" & Startdate & "','" & Enddate & "'," & year(txtransdate) & ",'" & User & "','Cash'"
oSaccoMaster.ExecuteThis (sql)

 ' REFLECT CASH PARTLY SALES HERE
sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Amount & ",'A046','A047','" & Lvwitems.SelectedItem & "','" & Lvwitems.SelectedItem.SubItems(1) & "' ,'PARTLY CASH CHECK OFF SALES ','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)
    ' REDUCE CREDIT SALES HERE
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Amount & ",'I013','I014','" & Lvwitems.SelectedItem & "','" & Lvwitems.SelectedItem.SubItems(1) & "' ,'PARTLY CASH CHECK OFF SALES ','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)
End If

'//Update deductions
If chkPrint.Value = vbChecked Then
PrintReceipt
PrintReceipt
End If
End If

Lvwitems.ListItems.Clear
txtpcode.Text = ""
txtquantity = ""
txtamount = ""
cboproductname = ""
txtrno = ""
txtSNo = ""
lblTKgs = ""
lblGPay = ""
lblDed = ""
lblNPay = ""
lblSNames = ""
txtcomment = ""
cmdnew_Click
MsgBox "Records saved"

Exit Sub
ebraim:
MsgBox err.description & " error occured."

End Sub
Private Sub savetransporters()
On Error GoTo kiparu2
Set Rst = New Recordset
Dim a, b, X
DIA = 0
Dim U As Double, S As Double
Dim cn As Connection
Dim j As Integer

If opttransport = True Then
If txttranscode = "" Then
MsgBox "Please enter the transporter"
Exit Sub
End If
sql = "d_sp_TransEnquiry  '" & txttranscode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
lbltransportername = rs.Fields(0)
Else
lbltransportername = ""
End If

If Trim(lbltransportername.Caption) = "" Then
MsgBox "Please enter a valid transporter number."
txttranscode.SetFocus
Exit Sub
End If

If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items sold."
Exit Sub
End If
j = 1


Total = 0
Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
 Total = Total + CCur(Lvwitems.SelectedItem.SubItems(4))
j = j + 1
Loop

If optCash.Value = False Then
If Total > CCur(lbltransnetpay) Then
If MsgBox("Transporter number " & txttranscode & " has a netpay of " & lblNPay & " do you wish to proceed?", vbYesNo) = vbYes Then
Else
Exit Sub
End If
End If


Startdate = DateSerial(year(txtransdate), month(txtransdate), 1)
Enddate = DateSerial(year(txtransdate), month(txtransdate) + 1, 1 - 1)


Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If
'End If
End If
j = 1
For j = 1 To Lvwitems.ListItems.Count
'Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
If Trim(txttranscode) = "" Then
txttranscode = "0"
End If
'// check if they are in stock.

Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qout from ag_products where p_code='" & Lvwitems.SelectedItem & "'"
'//Set rsinstock = New ADODB.Recordset
Set rsinstock = oSaccoMaster.GetRecordset(sql)
'// check the stock if it is less than zero
If rsinstock.Fields(1) <= 0 Then
MsgBox "Sorry Stock is Zero for item " & Lvwitems.SelectedItem.SubItems(1) & " please re-stock before your proceed", vbInformation
Exit Sub
End If
'// check the quanttity being sold versus the balance
Dim Remain As Double
Remain = rsinstock.Fields(1) - CInt(Lvwitems.SelectedItem.SubItems(2))
If Remain < 0 Then
MsgBox "Stock will be negative " & Remain & " please re-stock before you proceed", vbInformation
Exit Sub
End If

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Remarks,Description) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & txttranscode & "','checkoff','" & txtcomment & "')"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "'")
'j = j + 1
'Loop
    
    sql = ""
sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'A047','I013','" & Lvwitems.SelectedItem & "','" & Lvwitems.SelectedItem.SubItems(1) & "' ,' CHECK OFF SALES ','" & User & "',0,0)"
oSaccoMaster.ExecuteThis (sql)

Next j
'//Update deductions
If optCash.Value = False Then
Set cn = New ADODB.Connection
sql = "d_sp_TransDeduct '" & txttranscode & "','" & txtransdate & "','Agrovet'," & Total & ",'" & Startdate & "','" & Enddate & "','" & User & "','Receiptno " & txtrno & "'"
oSaccoMaster.ExecuteThis (sql)
End If

If CDbl(txtamtreceived) > 0 Then
Amount = 0
Amount = CDbl(txtamtreceived) * 1
Set cn = New ADODB.Connection
sql = "d_sp_TransDeduct '" & txttranscode & "','" & txtransdate & "','Agrovet'," & Amount & ",'" & Startdate & "','" & Enddate & "','" & User & "','Cash'"
oSaccoMaster.ExecuteThis (sql)

 ' REFLECT CASH PARTLY SALES HERE
sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Amount & ",'A046','A047','" & Lvwitems.SelectedItem & "','" & Lvwitems.SelectedItem.SubItems(1) & "' ,'PARTLY CASH CHECK OFF SALES ','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)
    ' REDUCE CREDIT SALES HERE
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Amount & ",'I013','I014','" & Lvwitems.SelectedItem & "','" & Lvwitems.SelectedItem.SubItems(1) & "' ,'PARTLY CASH CHECK OFF SALES ','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)
End If

If chkPrint.Value = vbChecked Then
PrintReceipt
PrintReceipt
End If

Lvwitems.ListItems.Clear
txttranscode = ""
txtrno.Text = ""
txtpcode.Text = ""
txtcomment = ""
lbltransnetpay = ""
txtquantity = 1
txtamount = ""
 
MsgBox "Records saved"
Exit Sub
kiparu2:
MsgBox err.description & " error occured."
End If
End Sub
Private Sub PrintReceipt()
    On Error GoTo sysError
    Dim strReceipts As String
    Dim pay, tot, disc As Currency
    Dim Z, X As Integer
    Dim a As Integer
    Dim b As Integer
    Dim mode As String
    
    If lblCheckOff.Value = True Or opttransport.Value = True Then
    mode = "CHECKOFF"
    ElseIf optCash.Value = True Then
    mode = "CASH SALES"
    ElseIf optstaff.Value = True Then
    mode = "CREDIT SALES"
    ElseIf Optbranch.Value = True Then
    mode = "DISPATCH TO STATION"
    End If
    
    dlg9.CancelError = True
    dlg9.FontName = "Garamond"
    Dim j As Printer
    a = dlg9.Copies
    Printer.CurrentY = 500
    Printer.CurrentX = 9000
    Printer.FontSize = 8
    Printer.CurrentY = 500
    Printer.CurrentX = 1000
    Printer.Print Tab(0); "     " & cname & ""
    Printer.Print Tab(0); "      " & paddress & ""
    Printer.Print Tab(0); "      " & Phone & ""
    Printer.Print Tab(0); "Email:" & Email & ""
    Printer.Print Tab(0); "--------------------------------------------------------------"
    Printer.Print Tab(0); "    AGROVET RECEIPT"
    Printer.Print Tab(0); "    " & mode & ""
    If lblCheckOff = True Then
    Printer.Print Tab(2); "SNo:" & txtSNo
    Printer.Print Tab(2); "Name:" & lblSNames
    ElseIf opttransport Then
    Printer.Print Tab(2); "TCODE:" & txttranscode
    Printer.Print Tab(2); "Name:" & lbltransportername
    ElseIf Optbranch Then
     Printer.Print Tab(2); "BRANCH:" & Cmbstation
    End If
    Printer.CurrentX = 500#
    Printer.FontSize = 10
    Printer.CurrentX = 500
    Printer.FontSize = 8
    Printer.CurrentX = 500
    Printer.FontSize = 8
    Printer.CurrentX = 500
    Printer.FontSize = 8
        a = 1
        strReceipts = ""
        Do While Not a > (Lvwitems.ListItems.Count)
            Lvwitems.ListItems.Item(a).selected = True
            strReceipts = strReceipts & Lvwitems.SelectedItem.SubItems(1) & vbNewLine & Lvwitems.SelectedItem.SubItems(2) & vbTab & vbTab
            strReceipts = strReceipts & Format(Lvwitems.SelectedItem.SubItems(3), "#,##0.00") & vbTab & vbTab & Format(Lvwitems.SelectedItem.SubItems(4), "#,##0.00") & vbNewLine
            a = a + 1
        Loop
        strReceipts = strReceipts & vbNewLine & "--------------------------------------------------------------" & vbNewLine
        strReceipts = strReceipts & "TOTAL" & vbTab & vbTab & vbTab & vbTab & Format(Total, "#,##0.00") & vbNewLine
        strReceipts = strReceipts & "==================================="
        
    Printer.Print Tab(2); "Item Description"
    Printer.Print Tab(0); "--------------------------------------------------------------"
    Printer.Print Tab(0); "QNTY" & vbTab & vbTab & "PRICE" & vbTab & vbTab & "AMOUNT"
    Printer.Print Tab(0); "........................................................................"
    Printer.Print Tab(0); strReceipts
    Printer.Print Tab(0); "AMOUNT RECEVED" & vbTab & vbTab & txtamtreceived
    Printer.Print
    Printer.Print Tab(0); "CHANGE" & vbTab & vbTab & vbTab & IIf(CDbl(TXTCHANGE) < 0, 0, CDbl(TXTCHANGE))
    Printer.Print
    Printer.Print Tab(0); "----------------------------------------------------------------"
    Printer.Print Tab(2); "Customer Signature   /Thumb Print"
    Printer.Print
    Printer.Print Tab(0); "----------------------------------------------------------------"
    Printer.Print Tab(2); "You were Served by: " & UCase(username)
    Printer.Print
    Printer.Print Tab(2); "----------------------------------------------------------------"
    Printer.Print Tab(0); "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    Printer.Print Tab(2); "     THANK YOU AND WELCOME "
    Printer.Print Tab(2); "POWERED BY EASYMA 6.0"
    Printer.Print Tab(2); "DEVELOP BY: AMTECH TECHNOLOGIES LTD"
    Printer.Print Tab(2); "******************************************************"
    Printer.Print
    Printer.EndDoc

    Exit Sub
sysError:
    MsgBox err.description, vbInformation
End Sub


Private Sub savestation()
On Error GoTo kiparu

If Optbranch = True Then
Dim centre As String
centre = Cmbstation.Text
If Trim(Cmbstation.Text) = "" Then
 MsgBox "Please enter the Agrovet Station."
   Cmbstation.SetFocus
Exit Sub
End If


Dim j As Integer
If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items sold."
Exit Sub
End If
j = 1

Dim pprice As Currency
Total = 0
Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
 Total = Total + CCur(Lvwitems.SelectedItem.SubItems(4))
j = j + 1
Loop


Startdate = DateSerial(year(txtransdate), month(txtransdate), 1)
Enddate = DateSerial(year(txtransdate), month(txtransdate) + 1, 1 - 1)


Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If
'End If
'End If
j = 1
For j = 1 To Lvwitems.ListItems.Count

 Lvwitems.ListItems.Item(j).selected = True


Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qout,PPrice from ag_products where p_code='" & Lvwitems.SelectedItem & "'"

Set rsinstock = oSaccoMaster.GetRecordset(sql)


Dim Remain As Double
Remain = rsinstock.Fields(1) - CInt(Lvwitems.SelectedItem.SubItems(2))

Amount = 0
Amount = Lvwitems.SelectedItem.SubItems(3) * Lvwitems.SelectedItem.SubItems(2)
sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Remarks,Description) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Amount
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & centre & "','Dispatch to station','" & txtcomment & "')"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "'")

Dim DRaccno As String
Dim Craccno As String

'XXXXXXXXXXX SAVE TO GL
Set rs2 = oSaccoMaster.GetRecordset("SELECT Agro_Debtor,Agro_Sales FROM Ag_Station WHERE Station='" & centre & "'")
If Not rs2.EOF Then
DRaccno = IIf(IsNull(rs2.Fields(0)), centre, rs2.Fields(0))
Craccno = IIf(IsNull(rs2.Fields(1)), centre, rs2.Fields(1))
Else
DRaccno = centre
Craccno = centre
End If

    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Amount & ",'" & DRaccno & "','" & Craccno & "','" & Lvwitems.SelectedItem & "','" & Lvwitems.SelectedItem.SubItems(1) & "' ,' DISPATCH STATION ','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)


'XXXXXXXXXXXXXXXXXXXXXX


Next j

    If chkPrint.Value = vbChecked Then
        PrintReceipt
        PrintReceipt
    End If
End If
'//Update deductions
'If optCash.value = False Then
'Set cn = New ADODB.Connection
'sql = "d_sp_TransDeduct '" & txttranscode & "','" & txtransdate & "','Agrovet'," & Total & ",'" & Startdate & "','" & Enddate & "','" & User & "'"
'oSaccoMaster.ExecuteThis (sql)

'Set cn = New ADODB.Connection
'sql = "d_sp_TransDeduct '" & txtTCode & "','" & DTPDDate & "','" & cboDeductionType & "'," & txtamount & ",'" & DTPStartDate & "','" & DTPEndDate & "','" & User & "'"
'oSaccoMaster.ExecuteThis (sql)


''XXXXXXXXXXXXXXXXXXXXXXXxx
''\\ save to gl
'
'
'    sql = ""
'    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & txtquantity & " *" & txtPPrice & ",'" & lbldracc & "','" & lblcracc & "','stock intake','" & cbosupplier & "' ,'stock intake','" & User & "',0,0)"
'    oSaccoMaster.ExecuteThis (sql)
''
'
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Lvwitems.ListItems.Clear
txtrno = ""
txtpcode.Text = ""
txtquantity = 1
txtamount = ""
Cmbstation.Text = ""
txtcomment = ""
MsgBox "Record saved Successfully"
Exit Sub
kiparu:
MsgBox err.description & " error occured."
End Sub
Private Sub savestaff()
On Error GoTo olkalou

If optstaff = True Then
Dim C As String

Dim centre As String
C = cbocategory.Text
If Trim(cbocategory.Text) = "" Then
 MsgBox "Please Select Credit Category", vbInformation
   cbocategory.SetFocus
Exit Sub
End If

Dim j As Integer
If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items sold."
Exit Sub
End If
''If txtstaffno = "" Then
''MsgBox "Enter Staff Number before you continue", vbCritical, "Maziwa"
''Exit Sub
''End If
''
''If txtstaffno <> "" Then
''   Set rs = oSaccoMaster.Get_Payroll_Recordset("Select empno from Employees where empno='" & Trim$(txtstaffno) & "'")
''             If rs.EOF Then
''                MsgBox "Staff No is Not Valid or Does Not Exceed,Enter Staffno correctly", vbInformation
''               Exit Sub
''             End If
''
''End If

j = 1


Total = 0
Do While Not j > (Lvwitems.ListItems.Count)
'For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True
 Total = Total + CCur(Lvwitems.SelectedItem.SubItems(4))
j = j + 1
Loop

If optCash.Value = False Then

Startdate = DateSerial(year(txtransdate), month(txtransdate), 1)
Enddate = DateSerial(year(txtransdate), month(txtransdate) + 1, 1 - 1)


Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If
'End If
End If
j = 1
For j = 1 To Lvwitems.ListItems.Count
'Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
'If Trim(txtstaffno) = "" Then
'txtstaffno = "0"
'End If
'// check if they are in stock.

Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qout from ag_products where p_code='" & Lvwitems.SelectedItem & "'"

Set rsinstock = oSaccoMaster.GetRecordset(sql)

Dim Remain As Double
Remain = rsinstock.Fields(1) - CInt(Lvwitems.SelectedItem.SubItems(2))

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Remarks,Description) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','Credit','" & txtcomment & "')"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "'")
'//XXXXXXXXXXXXXXX
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'A048','I013','" & Lvwitems.SelectedItem & "','" & Lvwitems.SelectedItem.SubItems(1) & "' ,'CREDIT SALES ','" & User & "',0,0)"
    oSaccoMaster.ExecuteThis (sql)

'XXXXXXXXXXXXXXXXXXXXXX
Next j


''sql = "set dateformat dmy insert into Oded (EmpNo, DCode, DDesc, Amount, DDate, EnterDate, FDate, EDate, UUser) " _
''       & " VALUES ('" & Trim$(txtstaffno) & "','D05','AGROVET'," & -1 * Amount & ",'" & txtransdate & "','" & txtransdate & "','" & Startdate & "','" & Enddate & "','" & User & "')"
''             oSaccoMaster.ExecuteThis_Payroll (sql)
''           '  oSaccoMaster.Get_Payroll_Recordset (sql)
''End If


If chkPrint.Value = vbChecked Then
PrintReceipt
PrintReceipt
End If
'//Update deductions
'If optCash.value = False Then
'Set cn = New ADODB.Connection
'sql = "d_sp_TransDeduct '" & txttranscode & "','" & txtransdate & "','Agrovet'," & Total & ",'" & Startdate & "','" & Enddate & "','" & User & "'"
'oSaccoMaster.ExecuteThis (sql)

'Set cn = New ADODB.Connection
'sql = "d_sp_TransDeduct '" & txtTCode & "','" & DTPDDate & "','" & cboDeductionType & "'," & txtamount & ",'" & DTPStartDate & "','" & DTPEndDate & "','" & User & "'"
'oSaccoMaster.ExecuteThis (sql)





Lvwitems.ListItems.Clear
txtrno = ""
txtpcode.Text = ""
txtquantity = 1
txtamount = ""
txtcomment = ""
cbocategory.Text = ""
MsgBox "Record saved Successfully"
Exit Sub
End If
olkalou:
MsgBox err.description & " error occured."

End Sub

Private Sub savecash()
On Error GoTo olkalou

If optCash = True Then
Dim C As String
C = "cash"

Dim j As Integer
If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items sold."
Exit Sub
End If
j = 1


Total = 0
Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
 Total = Total + CCur(Lvwitems.SelectedItem.SubItems(4))
j = j + 1
Loop



Startdate = DateSerial(year(txtransdate), month(txtransdate), 1)
Enddate = DateSerial(year(txtransdate), month(txtransdate) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If


'// check if they are in stock.
For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True


Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qout from ag_products where p_code='" & Lvwitems.SelectedItem & "'"
'//Set rsinstock = New ADODB.Recordset
Set rsinstock = oSaccoMaster.GetRecordset(sql)
'// check the stock if it is less than zero
If rsinstock.Fields(1) <= 0 Then
MsgBox "Sorry Stock is Zero for item " & Lvwitems.SelectedItem.SubItems(1) & " please re-stock before your proceed", vbInformation
Exit Sub
End If
'// check the quanttity being sold versus the balance
Dim Remain As Double
Remain = rsinstock.Fields(1) - CInt(Lvwitems.SelectedItem.SubItems(2))
If Remain < 0 Then
MsgBox "Stock will be negative " & Remain & " please re-stock before you proceed", vbInformation
Exit Sub
End If

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Remarks,Description) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & User & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','Cash','" & txtcomment & "')"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "'")

'\\ save to gl


    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'A046','I014','" & Lvwitems.SelectedItem & "','" & Lvwitems.SelectedItem.SubItems(1) & "' ,'cash sales','" & User & "',1,0)"
    oSaccoMaster.ExecuteThis (sql)
'

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
Next j


If chkPrint.Value = vbChecked Then
PrintReceipt
PrintReceipt
End If
End If
Lvwitems.ListItems.Clear
txtrno = ""
txtcomment = ""
txtpcode.Text = ""
txtquantity = 1
txtamount = ""
MsgBox "Record saved Successfully"
Exit Sub
olkalou:
MsgBox err.description & " error occured."

End Sub

Private Sub cmdSearchMember_Click()
 frmsearchEmployee.Show vbModal
    txtstaffno.Text = SearchValue
End Sub

Private Sub Command1_Click()
frmproduct1s.Show vbModal
End Sub

Private Sub Command2_Click()
Dim Total As Double
Dim j, Coun As Integer
j = 1
On Error GoTo ErrorHandler
TXTTOTAL = 0
'If Lvwitems.ListItems.Count > 0 Then
''Total = CCur(txttotal - li.SubItems(4))
Lvwitems.ListItems.Remove (Lvwitems.SelectedItem.Index)  '// removes the selected item

Do While Not j > (Lvwitems.ListItems.Count)
'For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True
 Total = Total + CCur(Lvwitems.SelectedItem.SubItems(4))
 TXTTOTAL = Total
j = j + 1
Loop

'End If
Exit Sub
ErrorHandler:
MsgBox err.description

End Sub

Private Sub Form_Load()
Label5.Visible = False
txttranscode.Visible = False
lbltransportername.Visible = False
Label10.Visible = False
lbltransnetpay.Visible = False
txtransdate = Format(Date, "dd/mm/yyyy")

Provider = "MAZIWA"
Set cn = New ADODB.Connection
cn.Open Provider, "bi"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_NAME  from ag_products ORDER BY P_NAME ASC"
Set rs = New ADODB.Recordset
rs.Open sql, cn

While Not rs.EOF
cboproductname.AddItem rs.Fields(0)
rs.MoveNext
Wend

    
    sql = "SELECT Station FROM Ag_Station"
    Set rs = oSaccoMaster.GetRecordset(sql)
    If rs.EOF Then Exit Sub
    With rs
        While Not .EOF
         Cmbstation.AddItem rs.Fields(0)
         .MoveNext
        Wend
    End With


cboproductname.Enabled = True
chkPrint.Value = vbChecked
End Sub
Private Sub cboname()
'Provider = cn
'Set cn = New ADODB.Connection
''cn.Open Provider, "bi"
''If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
'sql = "select P_NAME from ag_products where p_Name='" & cboproductname.Text & "'"
'Set rs = New ADODB.Recordset
'Set rs = oSaccoMaster.GetRecordset(sql)
'If Not rs.EOF Then
'If Not IsNull(rs.Fields(0)) Then cboproductname.Text = (rs.Fields(0))
'If Not IsNull(rs.Fields(1)) Then lblbalance = rs.Fields(1)
'End If
End Sub

Private Sub lblCheckOff_Click()
lblSNo.Visible = True
txtSNo.Visible = True
lblNetPay.Visible = True
lblNPay.Visible = True
lblDed.Visible = True
lblTKgs.Visible = True
lblGPay.Visible = True
Label11.Visible = True
Label13.Visible = True
Label8.Visible = True
txttranscode.Visible = False
Label5.Visible = False
Label10.Visible = False
lbltransnetpay.Visible = False
lbltransportername.Visible = False
txtstaffname.Visible = False
End Sub

Private Sub Optbranch_Click()
lblSNo.Visible = False
txtSNo.Visible = False
Label13.Visible = False
lblTKgs.Visible = False
Label11.Visible = False
lblGPay.Visible = False
Label8.Visible = False
lblDed.Visible = False
txtstaffname.Visible = False
lblNetPay.Visible = False
lblNPay.Visible = False
Label5.Visible = False
txttranscode.Visible = False
Label10.Visible = False
lbltransnetpay.Visible = False
lbltransportername.Visible = False
End Sub

Private Sub Optcash_Click()
lblSNo.Visible = False
txtSNo.Visible = False
lblNetPay.Visible = False
lblNPay.Visible = False

lblDed.Visible = False
lblTKgs.Visible = False
lblGPay.Visible = False
Label11.Visible = False
Label13.Visible = False
Label8.Visible = False

End Sub

Private Sub optstaff_Click()
lblSNo.Visible = False
txtSNo.Visible = False
lblNetPay.Visible = False
lblNPay.Visible = False

lblDed.Visible = False
lblTKgs.Visible = False
lblGPay.Visible = False
Label11.Visible = False
Label13.Visible = False
Label8.Visible = False
txtstaffname.Visible = True
End Sub

Private Sub opttransport_Click()
If opttransport = True Then
Label5.Visible = True
txttranscode.Visible = True
lbltransportername.Visible = True
Label10.Visible = True
lbltransnetpay.Visible = True
lblSNames.Visible = False
txtstaffname.Visible = False

lblSNo.Visible = False
txtSNo.Visible = False
Label13.Visible = False
lblTKgs.Visible = False
Label11.Visible = False
lblGPay.Visible = False
Label8.Visible = False
lblDed.Visible = False
lblNetPay.Visible = False
lblNPay.Visible = False
lblSNames.Visible = False

Else
Label5.Visible = False
txttranscode.Visible = False
lbltransportername.Visible = False
Label10.Visible = False
lbltransnetpay.Visible = False
End If
End Sub

Private Sub opttransport_Validate(Cancel As Boolean)
opttransport_Click
End Sub

Private Sub Cmbstation_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Picture1_Click()
frmSearch.Show vbModal
Dim Y As String
Y = sel
Dim p As Integer
If Y <> "" Then
'Provider = cn
Set cn = New ADODB.Connection
'cn.Open Provider, "bi"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,P_NAME,S_NO,QOUT,seria,s_no from ag_products where p_code='" & Y & "'"
Set rs = New ADODB.Recordset
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtpcode.Text = (rs.Fields(0))
If Not IsNull(rs.Fields(4)) Then p = (rs.Fields(4))
If p = 1 Then
If Not IsNull(rs.Fields(5)) Then 'txtserialno = (rs.Fields(5))
'lblserialno.Visible = True
'txtserialno.Visible = True
Else
'lblserialno.Visible = False
'txtserialno.Visible = False
End If
End If

If Not IsNull(rs.Fields(1)) Then cboproductname = (rs.Fields(1))

'If Not IsNull(rs.Fields(3)) Then lblbalance = (rs.Fields(3))
'// check if it has the serial numbers
'get_serialno Y
End If

'// check if the product have the serial then show the ag_receipts details
cboproductname_Validate True

End If
End Sub
Private Sub get_serialno(Pcode As String)
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, , "pius12"
Dim RSSE As Recordset
sql = ""
sql = "select top 1 serialno,p_code,used from serialno where p_code='" & txtpcode & "'  order by serialid desc"
Set RSSE = New ADODB.Recordset

RSSE.Open sql, cn, adOpenKeyset, adLockOptimistic
If RSSE.Fields(2) = 1 Then
MsgBox "Serial Number and receipt no used please check again before posting", vbCritical
Exit Sub
End If
End Sub
Private Sub Picture2_Click()
On Error Resume Next
frmsearchre.Show vbModal
Dim Y As String
Y = sel

If Y <> "" Then
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, , "pius12"
sql = ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select r_no,P_CODE,S_NO,Qua,amount from ag_receipts where r_no=" & Y & ""
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtrno = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtpcode = (rs.Fields(1))
'If Not IsNull(rs.Fields(2)) Then txtserialno = (rs.Fields(2))
If Not IsNull(rs.Fields(3)) Then txtquantity = (rs.Fields(3))
If Not IsNull(rs.Fields(4)) Then txtamount = (rs.Fields(4))
If Not IsNull(rs.Fields(3)) Then lblbalance = (rs.Fields(3))
Call cboname
End If
End If
End Sub

Private Sub txtpassword_LostFocus()
'fra1.Visible = True
Dim rsp As Recordset
Set cn = CreateObject("adodb.connection")
Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider, , "pius12"
Set rsp = CreateObject("adodb.recordset")
sql = "select *  from useraccounts where UserLoginID='" & User & "' and usergroup='administrator'"
rsp.Open sql, cn
Dim pass As String


txtransdate = Format(Date, "DD/MM/YYYY")
'fra1.Visible = True
'End If
End Sub
Private Sub Picture5_Click()
        Me.MousePointer = vbHourglass
        frmSearchSupplier.Show vbModal
        txtSNo = sel
        txtSNo_Validate True
        Me.MousePointer = 0
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtamtreceived_Change()
On Error Resume Next
TXTCHANGE = txtamtreceived - TXTTOTAL
End Sub

Private Sub txtpcode_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierid from ag_products where p_code='" & txtpcode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
 
If Not IsNull(rs.Fields(1)) Then cboproductname = (rs.Fields(1))



End If
End If
'// check with serial no if it exist
End Sub



Private Sub txtquantity_Validate(Cancel As Boolean)
If Not IsNumeric(txtquantity) Then
MsgBox "Enter a value please", vbCritical
txtquantity = ""
txtquantity.SetFocus
Exit Sub
End If
End Sub

Private Sub txtransdate_click()
'fra1.Visible = True
End Sub

Private Sub txtransdate_KeyPress(KeyAscii As Integer)
'fra1.Visible = True
End Sub

Private Sub txtransdate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'fra1.Visible = True
End Sub
Private Sub txtpassword_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
txtpassword_LostFocus
End Sub

Private Sub txtpcode_LostFocus()
Call cboname

End Sub
Private Sub txtserialno_LostFocus()
Dim rss As ADODB.Recordset
Dim rsproduct As ADODB.Recordset
sql = ""
sql = "select * from ag_products where seria=1 AND P_CODE='" & txtpcode & "'"
Set rsproduct = New ADODB.Recordset
rsproduct.Open sql, cn, adOpenKeyset, adLockOptimistic
If Not rsproduct.EOF Then
sql = ""
sql = "select serialno  from serialno "
Set rss = New ADODB.Recordset
rss.Open sql, cn, adOpenKeyset, adLockOptimistic
If Not rss.EOF Then
'// check if gth
While Not rss.EOF
Dim ser As String
ser = rss.Fields(0)

'If ser = txtserialno Then GoTo hererere

rss.MoveNext
Wend
Else
MsgBox "Serial no not in our database", vbInformation

DIA = 1
Exit Sub
End If
End If
hererere:
End Sub

Private Sub txtSNo_Validate(Cancel As Boolean)
On Error GoTo ErrorHandler
Set rs = New ADODB.Recordset
sql = "d_sp_SelectSuppliers '" & txtSNo & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then lblSNames = rs.Fields(2)
Else
lblSNames = ""
End If

Startdate = DateSerial(year(txtransdate), month(txtransdate), 1)
Enddate = DateSerial(year(txtransdate), month(txtransdate) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & Startdate & "','" & Enddate & "', 0")

If Not rs.EOF Then
lblTKgs = IIf(IsNull(rs.Fields(0)), 0, rs.Fields(0))
lblGPay = IIf(IsNull(rs.Fields(1)), 0, rs.Fields(1))
Else
lblTKgs = "0.00"
End If

'If Not IsNull(rs.Fields(1)) Then
'lblGPay = rs.Fields(1)
'Else
'lblGPay = "0.00"
'End If

Set rs = oSaccoMaster.GetRecordset("d_sp_SupNet " & txtSNo & ",'" & Startdate & "','" & Enddate & "', 1")
If Not IsNull(rs.Fields(0)) Then
lblDed = rs.Fields(0)
Else
lblDed = "0.00"
End If

lblNPay = Format((CCur(lblGPay) - CCur(lblDed)), "#,##0.00")

Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub txtstaffno_Change()
'If txtstaffno <> "" Then
'   Set rs = oSaccoMaster.Get_Payroll_Recordset("Select EMPNO,isnull(surname,'') +'   '+ isnull(OtherNames,'')as name  from Employees where EMPNO='" & Trim$(txtstaffno) & "'")
'             If Not rs.EOF Then
'               txtstaffname = IIf(IsNull(rs.Fields(1)), "", rs.Fields(1))
'               Else
'               txtstaffname = ""
'             End If
'
'End If
End Sub

Private Sub txttotal_Change()
On Error Resume Next
TXTCHANGE = txtamtreceived - TXTTOTAL
End Sub

Private Sub txttranscode_Change()
Set rs = New ADODB.Recordset
Dim dtpFrom As Date
sql = "d_sp_TransEnquiry  '" & txttranscode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then lbltransportername = rs.Fields(0)
End If
dtpFrom = DateSerial(year(txtransdate), month(txtransdate), 1)
DTPto = DateSerial(year(txtransdate), month(txtransdate) + 1, 1 - 1)
'oSaccoMaster.ExecuteThis ("d_sp_UpdateTranstmpEnquery '" & txttranscode & "','" & DTPto & "'")
'oSaccoMaster.ExecuteThis ("d_sp_UpdateTranstmpEnqueryDed '" & txttranscode & "','" & DTPfrom & "','" & DTPto & "'")
'
'sql = ""
'sql = "SELECT     TOP 1 Bal  FROM         d_tmpTransEnquery WHERE     (Code = '" & txttranscode & "') ORDER BY Bal DESC"
'Set Rst = oSaccoMaster.GetRecordset(sql)
'If Not Rst.EOF Then
'lbltransnetpay = IIf(IsNull(Rst.Fields(0)), 0, Rst.Fields(0))
'End If
' get transporter netpay
   Dim mMonth, yyear As Integer
   mMonth = month(txtransdate)
   yyear = year(txtransdate)
   
  sql = " Select(Select isnull(SUM(Amount + Subsidy),0) from d_TransDetailed where Trans_Code='" & txttranscode & "' and MMonth= " & mMonth & " and YYear=" & yyear & "),"
  sql = sql & " (Select isnull(SUM(Amount),0) from d_Transport_Deduc where TransCode='" & txttranscode & "' and MONTH(TDate_Deduc)=" & mMonth & " and YEAR(TDate_Deduc)= " & yyear & ")"
   Set rs2 = oSaccoMaster.GetRecordset(sql)
   If Not rs2.EOF Then
   lbltransnetpay = Format(rs2.Fields(0) - rs2.Fields(1), Cfmt)
  
   Else
   lbltransnetpay = "0.00"
   
   End If
End Sub
