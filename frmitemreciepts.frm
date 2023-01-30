VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmitemreciepts 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF8080&
   Caption         =   "Item Reciepts"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10170
   ScaleWidth      =   10665
   StartUpPosition =   1  'CenterOwner
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
      Left            =   1920
      TabIndex        =   40
      Top             =   9000
      Width           =   1425
   End
   Begin VB.PictureBox Picture21 
      Height          =   285
      Left            =   3405
      Picture         =   "frmitemreciepts.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   39
      Top             =   9015
      Width           =   300
   End
   Begin VB.PictureBox Picture4 
      Height          =   285
      Left            =   3315
      Picture         =   "frmitemreciepts.frx":02C2
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   36
      Top             =   8400
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
      Left            =   1920
      TabIndex        =   35
      Top             =   8400
      Width           =   1410
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
      Left            =   6120
      TabIndex        =   32
      Top             =   9600
      Value           =   2  'Grayed
      Width           =   1935
   End
   Begin VB.ComboBox ports 
      Height          =   315
      ItemData        =   "frmitemreciepts.frx":0584
      Left            =   9480
      List            =   "frmitemreciepts.frx":0594
      TabIndex        =   30
      Text            =   "COM1"
      Top             =   9720
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      Caption         =   "Close"
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   9600
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2880
      TabIndex        =   16
      Top             =   9600
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   9600
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   4210688
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmitemreciepts.frx":05B0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label11"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label12"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "dtprecDate"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtRef"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cboVendor"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtDelNo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cboStore"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtRemarks"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtQnty"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtOrdered"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "LPO Items"
      TabPicture(1)   =   "frmitemreciepts.frx":05CC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtcomment"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdRemove"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdAdd"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "LvwselectedItems"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lvwItems"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label15"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label14"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "LBLTOTAL"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label10"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label9"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      Begin VB.TextBox txtcomment 
         Height          =   615
         Left            =   -67680
         TabIndex        =   34
         Top             =   4080
         Width           =   2895
      End
      Begin VB.TextBox txtOrdered 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         Height          =   375
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "0"
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtQnty 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2640
         TabIndex        =   25
         Text            =   "0"
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   -69840
         TabIndex        =   23
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   -71280
         TabIndex        =   22
         Top             =   4200
         Width           =   1215
      End
      Begin MSComctlLib.ListView LvwselectedItems 
         Height          =   2535
         Left            =   -74760
         TabIndex        =   19
         Top             =   4680
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "LPO NO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "LPO Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Ordered Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Delivery Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Rejected Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Buying Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Balance"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvwItems 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   18
         Top             =   360
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   6588
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ReceiptNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ord Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ProductName"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Ordered Qyt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Delivery Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Balance"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtRemarks 
         Height          =   1815
         Left            =   240
         TabIndex        =   13
         Top             =   4200
         Width           =   6975
      End
      Begin VB.ComboBox cboStore 
         Height          =   315
         Left            =   2640
         TabIndex        =   11
         Text            =   "<Select Store>"
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox txtDelNo 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   1920
         Width           =   2175
      End
      Begin VB.ComboBox cboVendor 
         Height          =   315
         Left            =   2640
         TabIndex        =   8
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtRef 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         Height          =   375
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtprecDate 
         Height          =   375
         Left            =   2640
         TabIndex        =   10
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   109707265
         CurrentDate     =   40110
      End
      Begin VB.Label Label15 
         Caption         =   "Comment"
         Height          =   255
         Left            =   -68400
         TabIndex        =   33
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "TOTAL"
         Height          =   375
         Left            =   -67560
         TabIndex        =   29
         Top             =   7320
         Width           =   615
      End
      Begin VB.Label LBLTOTAL 
         Caption         =   "0"
         Height          =   255
         Left            =   -66840
         TabIndex        =   28
         Top             =   7320
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "Quantity Ordered :"
         Height          =   255
         Left            =   3720
         TabIndex        =   26
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Quantity Delivered :"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "This Delivery"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   21
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "LPO ITEM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69600
         TabIndex        =   20
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Store"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Recieved Date"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Delivery No"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Vendor"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Reference"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Delivery Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
   End
   Begin MSComDlg.CommonDialog dlg9 
      Left            =   8880
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Suppliers Creditors ACC"
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
      Left            =   120
      TabIndex        =   42
      Top             =   9000
      Width           =   1755
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
      Left            =   3600
      TabIndex        =   41
      Top             =   9000
      Width           =   4215
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
      Left            =   3615
      TabIndex        =   38
      Top             =   8400
      Width           =   4170
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Agrovet Purchase ACC"
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
      Left            =   120
      TabIndex        =   37
      Top             =   8445
      Width           =   1695
   End
   Begin VB.Label Label13 
      Caption         =   "Printer Port"
      Height          =   375
      Left            =   8400
      TabIndex        =   31
      Top             =   9720
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Item Reciept"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmitemreciepts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objLabelEdit As LabelEdit
Dim objLabelEdit2 As LabelEdit
Dim objLabelEdit3 As LabelEdit
Dim lenght As Double

Private Sub cboStore_KeyPress(KeyAscii As Integer)
KeyAscii = 0
Beep
End Sub

Private Sub cboVendor_Change()
lvwItems.ListItems.Clear
Set rs = oSaccoMaster.GetRecordset("d_sp_SupplierOrderedGoods '" & cboVendor & "'")
'Set rs = oSaccoMaster.GetRecordset("d_sp_IOrdered")
While Not rs.EOF
Set li = lvwItems.ListItems.Add(, , rs.Fields(0))
                li.SubItems(1) = rs.Fields(1) & ""
                li.SubItems(2) = rs.Fields(2) & ""
                li.SubItems(3) = rs.Fields(3) & ""
                li.SubItems(4) = rs.Fields(4) & ""
                li.SubItems(5) = "0" & ""
                li.SubItems(6) = rs.Fields(4) & ""
                
rs.MoveNext
Wend
End Sub

Private Sub cboVendor_Click()
cboVendor_Change
End Sub

Private Sub cboVendor_KeyPress(KeyAscii As Integer)
KeyAscii = 0
Beep
End Sub

Private Sub cmdAdd_Click()
 
If lvwItems.ListItems.Count = 0 Then
    MsgBox "There is no records to add"
        cmdAdd.SetFocus
    Exit Sub
End If
If cboVendor = "" Then
    MsgBox "Please select Vendor First ", vbInformation, Me.Caption
    SSTab1.Visible = True
        cboVendor.SetFocus
    Exit Sub
End If

Set li = LvwselectedItems.ListItems.Add(, , lvwItems.SelectedItem)
    
    Dim pp As Double
    
    Set Rst = oSaccoMaster.GetRecordset("select pprice from ag_products where p_name='" & lvwItems.SelectedItem.ListSubItems(2) & "'")
    If Not Rst.EOF Then
    pp = Rst.Fields(0)
    End If
                        li.SubItems(1) = lvwItems.SelectedItem.ListSubItems(1) & ""
                        li.SubItems(2) = lvwItems.SelectedItem.ListSubItems(2) & ""
                        li.SubItems(3) = lvwItems.SelectedItem.ListSubItems(4) & ""
                        li.SubItems(4) = lvwItems.SelectedItem.ListSubItems(5) & ""
                        li.SubItems(5) = "0" & ""
                        li.SubItems(6) = pp
                        li.SubItems(7) = IIf(IsNull(lvwItems.SelectedItem.ListSubItems(6)), 0, lvwItems.SelectedItem.ListSubItems(5)) & ""
lvwItems.ListItems.Remove (lvwItems.SelectedItem.Index)
Calculate_Total
Exit Sub

End Sub
Sub Calculate_Total()

    Dim Total As Double, amt As Double, Price As Double, qnty As Integer
    Dim ccount As Integer
    On Error Resume Next
    Total = 0
    With LvwselectedItems
        If .ListItems.Count > 0 Then
            ccount = .ListItems.Count
            For I = 1 To ccount
                With .ListItems(I)
                        Price = CDbl(.ListSubItems(6))
                        qnty = CDbl(.ListSubItems(4))
                        amt = Price * qnty
                        Total = Total + amt
                End With
            Next I

        Else
            Total = 0
        End If
    End With
    LBLTOTAL.Caption = Total
End Sub

Private Sub cmdClear_Click()
Form_Load
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub Command4_Click()

End Sub

Private Sub cmdRemove_Click()
If LvwselectedItems.ListItems.Count = 0 Then
    MsgBox "There is no records to Remove"
        cmdRemove.SetFocus
    Exit Sub
End If


Set li = lvwItems.ListItems.Add(, , LvwselectedItems.SelectedItem)
                        li.SubItems(1) = LvwselectedItems.SelectedItem.ListSubItems(1) & ""
                        li.SubItems(2) = LvwselectedItems.SelectedItem.ListSubItems(2) & ""
                        li.SubItems(3) = LvwselectedItems.SelectedItem.ListSubItems(3) & ""
                        li.SubItems(4) = LvwselectedItems.SelectedItem.ListSubItems(4) & ""
                        li.SubItems(5) = "0" & ""
                        li.SubItems(6) = LvwselectedItems.SelectedItem.ListSubItems(7) & ""
LvwselectedItems.ListItems.Remove (LvwselectedItems.SelectedItem.Index)
Calculate_Total


End Sub

Private Sub cmdSave_Click()
Dim Pcode, ReceiptNo As String, lenght As Integer
'If txtRef = "" Then
'    MsgBox "Please enter the reference number."
'        txtRef.SetFocus
'    Exit Sub
'End If
'
If TxtOtherPAcc = "" Then
    MsgBox "Please enter the Suppliers Creditors GL Account", vbInformation
        TxtOtherPAcc.SetFocus
    Exit Sub
End If

If txtcontra = "" Then
    MsgBox "Please enter the Agrovet Stock Purchase Gl Account", vbInformation
       txtcontra.SetFocus
    Exit Sub
End If
Dim j As Integer
dtprecDate = Format(dtprecDate, "dd/mm/yyyy")

j = 1
For j = 1 To LvwselectedItems.ListItems.Count
' LvwselectedItems.ListItems.Item(j).selected = True

Set li = LvwselectedItems.ListItems(j)
sql = "select P_CODE,qout,unserialized,pprice,sprice,o_bal from ag_products where p_name='" & LvwselectedItems.ListItems(j).SubItems(2) & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
Pcode = rs.Fields("p_code")
sql = ""
sql = "set dateformat DMY update ag_products set qin=" & LvwselectedItems.ListItems(j).SubItems(4) & ",qout=" & LvwselectedItems.ListItems(j).SubItems(4) + rs.Fields("qout") & ",o_bal=" & LvwselectedItems.ListItems(j).SubItems(4) + rs.Fields("qout") & ",last_d_updated='" & dtprecDate & "',user_id='" & user & "',audit_date='" & Get_Server_Date & "',unserialized=0,SERIA=0,pprice=" & rs.Fields("pprice") & ",sprice=" & rs.Fields("sprice") & " where p_code='" & Pcode & "'"
cn.Execute sql

Dim rsst As Recordset
'save inventory received
ReceiptNo = Trim$(li)
sql = "Insert into Ag_Received(pcode ,Rno,Pprice, sprice, Amount, Supplier, QTY, transdate, auditid, description)" _
      & " Values('" & Pcode & "','" & ReceiptNo & "','" & LvwselectedItems.ListItems(j).SubItems(4) & "',0,'" & LvwselectedItems.ListItems(j).SubItems(4) * LvwselectedItems.ListItems(j).SubItems(6) & "','" & cboVendor & "','" & LvwselectedItems.ListItems(j).SubItems(4) & "','" & dtprecDate & "','" & user & "','" & txtcomment & "')"
oSaccoMaster.ExecuteThis (sql)

sql = "Update d_Requisition set [status]='Receipt' ,Qnty=" & LvwselectedItems.ListItems(j).SubItems(4) & "  where RNo='" & ReceiptNo & "'"
  oSaccoMaster.ExecuteThis (sql)

Next j

  ' Save to Gl Stock Delivered
   sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & dtprecDate & "'," & CDbl(LBLTOTAL) & ",'" & txtcontra & "','" & TxtOtherPAcc & "','" & ReceiptNo & "','" & cboVendor & "' ,'Agrovet Stock Purchased','" & user & "',1,0)"
    oSaccoMaster.ExecuteThis (sql)


If chkPrint.value = vbChecked Then
PrintReceipt
PrintReceipt
End If
LvwselectedItems.ListItems.Clear
txtcomment = ""
MsgBox "saved successfully"
Form_Load
End Sub
Private Sub PrintReceipt()
    On Error GoTo sysError
    Dim strReceipts As String
    Dim pay, tot, disc As Currency
    Dim Z, X As Integer, Total As Double
    Dim a As Integer
    Dim b As Integer
    Dim mode As String
    
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
    Printer.Print Tab(0); "   RECEIVING VOUCHER"
    Printer.Print Tab(0); "  INVOICE NO:" & "-------------------"""
    Printer.Print Tab(0); "Vendor" & cboVendor
    Printer.Print Tab(0); "--------------------------------------------------------------"
    Printer.Print Tab(0); "ITEM" & vbTab & vbTab & "QNTY" & vbTab & "PRICE" & vbTab & "AMOUNT"
    Printer.Print Tab(0); "........................................................................"
           a = 1
        strReceipts = ""
        Do While Not a > (LvwselectedItems.ListItems.Count)
            LvwselectedItems.ListItems.Item(a).selected = True
            lenght = Len(LvwselectedItems.SelectedItem.SubItems(2))
            strReceipts = Mid(LvwselectedItems.SelectedItem.SubItems(2), 5, lenght - 5)
            If Len(strReceipts) > 14 Then
            strReceipts = strReceipts & "-"
            Else
            strReceipts = strReceipts & vbTab
            End If
            strReceipts = strReceipts & CDbl(LvwselectedItems.SelectedItem.SubItems(4)) & vbTab & Format(LvwselectedItems.SelectedItem.SubItems(6), "#,##0.00") & vbTab & Format((LvwselectedItems.SelectedItem.SubItems(4) * LvwselectedItems.SelectedItem.SubItems(6)), "#,##0.00") & vbNewLine
             Printer.Print Tab(0); strReceipts
            a = a + 1
        Loop
    Printer.CurrentX = 500#
    Printer.FontSize = 10
    Printer.CurrentX = 500
    Printer.FontSize = 8
    Printer.CurrentX = 500
    Printer.FontSize = 8
    Printer.CurrentX = 500
    Printer.FontSize = 8
    Printer.Print Tab(0); "RECEIPT TOTAL" & vbTab & vbTab & Format(LBLTOTAL, "#,##0.00") & vbNewLine
    Printer.Print Tab(0); "======================================="
    Printer.Print Tab(0); "Remarks" & vbTab & txtcomment
    Printer.Print Tab(0); "----------------------------------------------------------------"
    Printer.Print Tab(2); "Customer Signature   /Thumb Print"
    Printer.Print
    Printer.Print Tab(0); "----------------------------------------------------------------"
    Printer.Print Tab(2); "You were Served by: " & UCase(username)
    Printer.Print
    Printer.Print Tab(2); "----------------------------------------------------------------"
    Printer.Print Tab(0); "    Date :" & Format(Get_Server_Date, "dd/mm/yyyy HH:MM:SS AM/PM")
    Printer.Print Tab(2); "     THANK YOU AND WELCOME "
    Printer.Print Tab(0); "Powered By EasyMa Amtech Technologies"
    Printer.Print Tab(2); "******************************************************"
    Printer.Print
    Printer.EndDoc

    Exit Sub
sysError:
    MsgBox err.description, vbInformation
End Sub



Private Sub Form_Load()
chkPrint.value = vbChecked

LvwselectedItems.Enabled = True
txtDelNo = ""
txtRef = ""
txtRemarks = ""
txtQnty = ""
txtcomment = ""

cboStore.Clear
cboVendor.Clear
dtprecDate = Format(Get_Server_Date, "dd/mm/yyyy")

Set rs = oSaccoMaster.GetRecordset("SELECT  CompanyName  FROM ag_Supplier1 order by companyname")
While Not rs.EOF
cboVendor.AddItem rs.Fields(0)
rs.MoveNext
Wend

    cboStore.Clear
    sql = "Select companyname from ag_Supplier1"
    Set rs = oSaccoMaster.GetRecordset(sql)
    While Not rs.EOF
    cboStore.AddItem rs.Fields(0)
    rs.MoveNext
    Wend

lvwItems.ListItems.Clear
Set rs = oSaccoMaster.GetRecordset("d_sp_loadOrderedGoods")
'Set rs = oSaccoMaster.GetRecordset("d_sp_IOrdered")
While Not rs.EOF
Set li = lvwItems.ListItems.Add(, , rs.Fields(0))
                li.SubItems(1) = rs.Fields(1) & ""
                li.SubItems(2) = rs.Fields(2) & ""
                li.SubItems(3) = rs.Fields(3) & ""
                li.SubItems(4) = rs.Fields(4) & ""
                li.SubItems(5) = "0" & ""
                li.SubItems(6) = rs.Fields(4) & ""
                
rs.MoveNext
Wend

InitSubClass

'    Set objLabelEdit = New LabelEdit
'    objLabelEdit.Init Me, lvwItems
'    Set objLabelEdit2 = New LabelEdit
'    objLabelEdit2.Init Me, lvwselecteditems
'        InitSubClass
  
    'Enable label editing for listview2
    Set objLabelEdit = New LabelEdit
    objLabelEdit.Init Me, lvwItems
    Set objLabelEdit2 = New LabelEdit
    objLabelEdit2.Init Me, lvwItems



End Sub

Private Sub lvwItems_DblClick()
cmdAdd_Click
End Sub

Private Sub LvwselectedItems_Click()
  Dim Total As Double, amt As Double, Price As Double, qnty As Integer
    Dim ccount As Integer
    On Error Resume Next
    Total = 0
    With LvwselectedItems
        If .ListItems.Count > 0 Then
            ccount = .ListItems.Count
            For I = 1 To ccount
                With .ListItems(I)
                        Price = CDbl(.ListSubItems(6))
                        qnty = CDbl(.ListSubItems(4))
                        amt = Price * qnty
                        Total = Total + amt
                            Dim objLabelEdit2 As LabelEdit
                        Set objLabelEdit2 = New LabelEdit
                    objLabelEdit2.Init Me, .ListSubItems(4)

                End With
            Next I

        Else
            Total = 0
        End If
    End With
    LBLTOTAL.Caption = Total
    

End Sub

Private Sub lvwSelectedItems_DblClick()
Set rs = oSaccoMaster.GetRecordset("SELECT d_LPO.RefNo,d_Requisition.CostCentre,d_LPO.Vendor FROM d_LPO,d_Requisition WHERE d_Requisition.RNo=d_LPO.RefNo AND  d_LPO.PNo=" & LvwselectedItems.SelectedItem)
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then
txtRef = rs.Fields(0)
End If
If Not IsNull(rs.Fields(1)) Then
cboStore = rs.Fields(1)
End If
If Not IsNull(rs.Fields(2)) Then
cboVendor = rs.Fields(2)
End If

txtOrdered = LvwselectedItems.SelectedItem.ListSubItems(3)
End If
SSTab1.Tab = 0
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

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
cmdSave.Enabled = True
'cmdClear.Enabled = True
Else
cmdSave.Enabled = True
'cmdClear.Enabled = False
End If
End Sub

Private Sub txtcontra_Change()
 On Error GoTo sysError
    Dim Account As Acc_Details
    Editing = True
    Account = Get_Acc_Details(txtcontra, ErrorMessage)
    If Account.accno <> "" Then
        lblcontra = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblcontra = ""
    End If
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtDelNo_Validate(Cancel As Boolean)
    If Trim(txtQnty) = "" Then
        txtQnty = "0"
    End If
End Sub


Private Sub TxtOtherPAcc_Change()
On Error GoTo sysError
    Dim Account As Acc_Details
    Editing = True
    Account = Get_Acc_Details(TxtOtherPAcc, ErrorMessage)
    If Account.accno <> "" Then
        lblOtherPaymentAcc = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblOtherPaymentAcc = ""
    End If
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub txtQnty_KeyPress(KeyAscii As Integer)
If (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "Please enter a number "
End If
End Sub

