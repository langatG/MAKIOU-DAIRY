VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmapprovelpo 
   Caption         =   "Approve Local Purchase Order"
   ClientHeight    =   8580
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   11730
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print LPO"
      Height          =   375
      Left            =   1440
      TabIndex        =   32
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   8040
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      Caption         =   "Close"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   8040
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   8421631
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmapprovelpo.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtRef"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cboVendor"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtDelNo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboStore"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtRemarks"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtQnty"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtOrdered"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "dtprecDate"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label5"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label6"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label7"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label8"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label11"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label12"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "LPO Items"
      TabPicture(1)   =   "frmapprovelpo.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label15"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label9"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lvwItems"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "LvwselectedItems"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtcomment"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdRemove"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdAdd"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      Begin VB.TextBox txtRef 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         Height          =   375
         Left            =   -72360
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox cboVendor 
         Height          =   315
         Left            =   -72360
         TabIndex        =   14
         Text            =   "<Select Vendor>"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtDelNo 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -72360
         TabIndex        =   13
         Top             =   1920
         Width           =   2175
      End
      Begin VB.ComboBox cboStore 
         Height          =   315
         Left            =   -72360
         TabIndex        =   12
         Text            =   "<Select Store>"
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox txtRemarks 
         Height          =   1815
         Left            =   -74760
         TabIndex        =   11
         Top             =   4200
         Width           =   6975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   3720
         TabIndex        =   8
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   5160
         TabIndex        =   7
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox txtQnty 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -72360
         TabIndex        =   6
         Text            =   "0"
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txtOrdered 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         Height          =   375
         Left            =   -69840
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0"
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtcomment 
         Height          =   495
         Left            =   7680
         TabIndex        =   4
         Top             =   4200
         Width           =   3495
      End
      Begin MSComctlLib.ListView LvwselectedItems 
         Height          =   2895
         Left            =   120
         TabIndex        =   9
         Top             =   4680
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5106
         View            =   3
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "LPO NO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "RefNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "LPO Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Ordered Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Delivery Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Rejected Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Vendor"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvwItems 
         Height          =   3375
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   5953
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777190
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "LPO NO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "RefNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "LPO Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Ordered Qyt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Delivery Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Balance"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Vendor"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtprecDate 
         Height          =   375
         Left            =   -72360
         TabIndex        =   16
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Format          =   49545217
         CurrentDate     =   40110
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
         Left            =   -74760
         TabIndex        =   28
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Reference"
         Height          =   255
         Left            =   -74640
         TabIndex        =   27
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Vendor"
         Height          =   255
         Left            =   -74640
         TabIndex        =   26
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Delivery No"
         Height          =   255
         Left            =   -74640
         TabIndex        =   25
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Recieved Date"
         Height          =   255
         Left            =   -74640
         TabIndex        =   24
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Store"
         Height          =   255
         Left            =   -74640
         TabIndex        =   23
         Top             =   3480
         Width           =   1455
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
         Left            =   -74760
         TabIndex        =   22
         Top             =   3960
         Width           =   1335
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
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1095
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
         Left            =   240
         TabIndex        =   20
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Quantity Delivered :"
         Height          =   255
         Left            =   -74640
         TabIndex        =   19
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Quantity Ordered :"
         Height          =   255
         Left            =   -71280
         TabIndex        =   18
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Comment"
         Height          =   255
         Left            =   6960
         TabIndex        =   17
         Top             =   4320
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Goods Ordered"
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
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   7575
   End
   Begin VB.Label Label14 
      Caption         =   "TOTAL"
      Height          =   375
      Left            =   7560
      TabIndex        =   30
      Top             =   8160
      Width           =   615
   End
   Begin VB.Label LBLTOTAL 
      Caption         =   "0"
      Height          =   255
      Left            =   8400
      TabIndex        =   29
      Top             =   8160
      Width           =   2175
   End
End
Attribute VB_Name = "frmapprovelpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objLabelEdit As LabelEdit
Dim objLabelEdit2 As LabelEdit
Dim objLabelEdit3 As LabelEdit
Private Sub cboStore_KeyPress(KeyAscii As Integer)
KeyAscii = 0
Beep
End Sub

Private Sub cboVendor_Change()
lvwItems.ListItems.Clear
LvwselectedItems.ListItems.Clear

Set rs = oSaccoMaster.GetRecordset("spIOrderedVendor '" & cboVendor & "'")

While Not rs.EOF
Set li = lvwItems.ListItems.Add(, , rs.Fields(0))
                li.SubItems(1) = rs.Fields(1) & ""
                li.SubItems(2) = rs.Fields(2) & ""
                li.SubItems(3) = rs.Fields(3) & ""
                 li.SubItems(4) = rs.Fields(5) & ""
                li.SubItems(5) = rs.Fields(4) & ""
                li.SubItems(6) = "0" & ""
                li.SubItems(7) = rs.Fields(4) & ""
                li.SubItems(8) = rs.Fields(6) & ""
                
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

Set li = LvwselectedItems.ListItems.Add(, , lvwItems.SelectedItem)
                        li.SubItems(1) = lvwItems.SelectedItem.ListSubItems(1) & ""
                        li.SubItems(2) = lvwItems.SelectedItem.ListSubItems(2) & ""
                        li.SubItems(3) = lvwItems.SelectedItem.ListSubItems(3) & ""
                        li.SubItems(4) = lvwItems.SelectedItem.ListSubItems(4) & ""
                        li.SubItems(5) = lvwItems.SelectedItem.ListSubItems(5) & ""
                        li.SubItems(6) = lvwItems.SelectedItem.ListSubItems(6) & ""
                        li.SubItems(7) = lvwItems.SelectedItem.ListSubItems(7) & ""
                        li.SubItems(8) = lvwItems.SelectedItem.ListSubItems(8) & ""
                        li.SubItems(7) = "0" & ""

lvwItems.ListItems.Remove (lvwItems.SelectedItem.Index)

cmdSave.Enabled = True

Calculate_Total
Exit Sub
End Sub

Private Sub cmdClear_Click()
Form_Load
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Command4_Click()

End Sub

Private Sub cmdprint_Click()
          On Error GoTo TransError

   Set rs = oSaccoMaster.GetRecordset("SELECT top 1 pno FROM lpo order by pno desc")
    reportname = "LPO.rpt"
    If Not IsNull(rs.Fields(0)) Then
    STRFORMULA = "{LPO.pno}=" & Trim$(rs.Fields(0)) & "  and {Requisition.Status}='Receipt'"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
    End If
    Exit Sub
TransError:
        
        MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
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
                        li.SubItems(5) = LvwselectedItems.SelectedItem.ListSubItems(5) & ""
                        li.SubItems(6) = LvwselectedItems.SelectedItem.ListSubItems(6) & ""
                        li.SubItems(7) = LvwselectedItems.SelectedItem.ListSubItems(6) & ""
                        li.SubItems(8) = LvwselectedItems.SelectedItem.ListSubItems(8) & ""
                        'li.SubItems(5) = "0" & ""

LvwselectedItems.ListItems.Remove (LvwselectedItems.SelectedItem.Index)


End Sub

Private Sub cmdsave_Click()
Dim Refno As String, I As Long, lpoNo As Integer, LPoDate As Date, duedate As Date, itemName As String, QTY As Integer, Pprice As Double, SerialNo As String, _
Vendor As String, REMARKS As String

Dim SaveLpo As New ADODB.Connection
'If txtRef = "" Then
'    MsgBox "Please enter the reference number."
'        txtRef.SetFocus
'    Exit Sub
'End If

'If txtDelNo = "" Then
'    MsgBox "Please enter the delivery number."
'        txtDelNo.SetFocus
'    Exit Sub
'End If

'If CCur(txtQnty) = 0 Then
'    MsgBox "Please enter the quantity received."
'        txtQnty.SetFocus
'    Exit Sub
'End If

'If LvwselectedItems.ListItems.Count > 0 Then
'   LvwselectedItems.ListItems.Remove (LvwselectedItems.SelectedItem.index)
'End If
If LvwselectedItems.ListItems.Count = 0 Then
    MsgBox "There is no records to Save"
        cmdAdd.SetFocus
    Exit Sub
End If

On Error GoTo TransError
With SaveLpo
    If .State = adStateClosed Then
    .Open SelectedDsn
    End If
    .BeginTrans
      
I = 0
For I = 1 To LvwselectedItems.ListItems.Count
 Set li = LvwselectedItems.ListItems(I)
     lpoNo = Trim$(li)
     Refno = Trim$(li.ListSubItems(1))
     LPoDate = li.ListSubItems(2)
     itemName = li.ListSubItems(3)
     QTY = CDbl(li.ListSubItems(6))
     Pprice = CDbl(li.ListSubItems(4))
     'duedate = li.ListSubItems(5)
     SerialNo = Refno
     Vendor = li.ListSubItems(8)
     'Remarks = li.ListSubItems(7)

'd_sp_Receipts @R varchar(35), @V varchar(80), @D varchar(35), @Q float, @T varchar(12), @re varchar(85), @A varchar(35) AS
oSaccoMaster.ExecuteThis ("d_sp_Receipts '" & Refno & "','" & Vendor & "'," & QTY & "," & Pprice & "," & QTY * Pprice & ",'" & dtprecDate & "','" & txtcomment & "','" & User & "'," & lpoNo & " ")

'//add it to the items

'//get the name available
'Dim rsg As New ADODB.Recordset, rsh As New ADODB.Recordset, pcode As String
'Dim namee As String
'SQL = "SELECT     IName  FROM         d_Requisition  WHERE     (RNo = '" & txtRef & "')"
'Set rsg = oSaccoMaster.GetRecordset(SQL)
'If Not rsg.EOF Then
'namee = IIf(IsNull(rsg.Fields(0)), "", rsg.Fields(0))
'If namee <> "" Then
'        SQL = ""
'        SQL = "select P_CODE,p_name from ag_products where p_name like '" & namee & "%'"
'        Set rsh = oSaccoMaster.GetRecordset(SQL)
'        If Not rsh.EOF Then
'        pcode = rsh.Fields(0)
'        End If
'End If
'Else
'MsgBox "Product not available in the database, key in first in the agrovet module"
'Exit Sub
'End If
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)





'SQL = "select P_CODE,qout,unserialized,pprice,sprice,o_bal from ag_products where p_code='" & pcode & "'"
'Set rs = New ADODB.Recordset
'
'Set rs = oSaccoMaster.GetRecordset(SQL)
'
'
'SQL = ""
'SQL = "insert into i_stock_Purchases(itemId,stockBalance,mcode,openingStock,changeInStock,transDate,grnId,purchasePrice,sellingPrice,auditId,itemCode)values()"
''sql = ""
''sql = "set dateformat DMY update ag_products set qin=" & txtQnty & ",qout=" & txtQnty.Text + Rs.Fields("qout") & ",o_bal=" & txtQnty.Text + Rs.Fields("qout") & ",last_d_updated='" & dtprecDate & "',user_id='" & User & "',audit_date='" & Get_Server_Date & "',unserialized=0,SERIA=0,pprice=" & Rs.Fields("pprice") & ",sprice=" & Rs.Fields("sprice") & " where p_code='" & pcode & "'"
'oSaccoMaster.Execute (SQL)
'
'Dim rsst As Recordset
'SQL = ""
'SQL = "set dateformat DMY select top 1 * from ag_stockbalance where p_code='" & pcode & "' order by trackid desc "
'Set rsst = New ADODB.Recordset
'Set rsst = oSaccoMaster.GetRecordset(SQL)
'
'If Not rsst.EOF Then
'SQL = ""
'SQL = "set dateformat DMY INSERT INTO ag_stockbalance"
'SQL = SQL & " (p_code, productname, openningstock, changeinstock, stockbalance, transdate,companyid)"
'SQL = SQL & " VALUES     ('" & pcode & "', '" & namee & "', '" & rs.Fields("o_bal") & "', '" & txtQnty & "', '" & txtQnty.Text + rs.Fields("qout") & "', '" & Format(Get_Server_Date, "dd/mm/yyyy") & "',1)"
'oSaccoMaster.Execute (SQL)

'If Not Save_GLTRANSACTION(chcode, dtprecDate, rs.Fields("sprice"), "80102", "40201", pcode, namee, User, ErrorMessage, "agrovet goods received", 1, 1, pcode) Then
'    If ErrorMessage <> "" Then
'    MsgBox Err.description, vbInformation
'    End If
'End If

'End If

Next I

.CommitTrans

MsgBox "Records saved successfully!"
Form_Load
       On Error GoTo TransError

   Set rs = oSaccoMaster.GetRecordset("SELECT top 1 pno FROM lpo order by pno desc")
    reportname = "LPO.rpt"
    If Not IsNull(rs.Fields(0)) Then
    STRFORMULA = "{LPO.pno}=" & Trim$(rs.Fields(0)) & "  and {Requisition.Status}='Receipt'"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
    End If

Exit Sub
TransError:
        .RollbackTrans
        MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
    End With
End Sub

Private Sub Form_Load()
txtDelNo = ""
txtRef = ""
txtRemarks = ""
txtQnty = ""


cboStore.Clear
cboVendor.Clear
dtprecDate = Format(Get_Server_Date, "dd/mm/yyyy")


Set rs = oSaccoMaster.GetRecordset("SELECT  CompanyName  FROM ag_Supplier1 order by companyname")
While Not rs.EOF
cboVendor.AddItem rs.Fields(0)
rs.MoveNext
Wend
'cboVendor.Text = "<Select Vendor>"


Set rs = oSaccoMaster.GetRecordset("select description from  d_CostCent order by description")
While Not rs.EOF
cboStore.AddItem rs.Fields(0)
rs.MoveNext
Wend
'cboStore.Text = "<Select Store>"
lvwItems.ListItems.Clear
LvwselectedItems.ListItems.Clear
Set rs = oSaccoMaster.GetRecordset("spIOrdered ")

While Not rs.EOF
Set li = lvwItems.ListItems.Add(, , rs.Fields(0))
                li.SubItems(1) = rs.Fields(1) & ""
                li.SubItems(2) = rs.Fields(2) & ""
                li.SubItems(3) = rs.Fields(3) & ""
                 li.SubItems(4) = rs.Fields(5) & ""
                li.SubItems(5) = rs.Fields(4) & ""
                li.SubItems(6) = "0" & ""
                li.SubItems(7) = rs.Fields(4) & ""
                li.SubItems(8) = rs.Fields(6) & ""
                
rs.MoveNext
Wend

    InitSubClass

    Set objLabelEdit = New LabelEdit
    objLabelEdit.Init Me, lvwItems
    Set objLabelEdit2 = New LabelEdit
    objLabelEdit2.Init Me, lvwItems

End Sub

Private Sub lvwItems_DblClick()
cmdAdd_Click
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

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
cmdSave.Enabled = True
'cmdClear.Enabled = True
Else
cmdSave.Enabled = False
'cmdClear.Enabled = False
End If
End Sub
Private Sub txtDelNo_Validate(Cancel As Boolean)
    If Trim(txtQnty) = "" Then
        txtQnty = "0"
    End If
End Sub


Private Sub txtQnty_KeyPress(KeyAscii As Integer)
If (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 48) Or (KeyAscii = 49) Or (KeyAscii = 50) Or (KeyAscii = 51) Or (KeyAscii = 52) Or (KeyAscii = 53) Or (KeyAscii = 54) Or (KeyAscii = 55) Or (KeyAscii = 56) Or (KeyAscii = 57) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "Please enter a number "
End If
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
                        Price = CDbl(.ListSubItems(4))
                        qnty = CDbl(.ListSubItems(6))
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



