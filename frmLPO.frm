VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLPO 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Local Purchase Order"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   10485
   Icon            =   "frmLPO.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   10485
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txttotal 
      Height          =   375
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   7320
      Width           =   2295
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print LPO"
      Height          =   375
      Left            =   4800
      TabIndex        =   32
      Top             =   7320
      Width           =   1575
   End
   Begin VB.CommandButton cmSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   7320
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   11880
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   12648447
      TabCaption(0)   =   "Purchase Order Info"
      TabPicture(0)   =   "frmLPO.frx":0CCE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label12"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label13"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label14"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label15"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "DTPlpodate"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "DTPduedate"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtRemarks"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtLPOSerial"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cbovendors"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtPoNo"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtIName"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtRefNo"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtQnty"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtPPrice"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdNew"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Purchase Order Items"
      TabPicture(1)   =   "frmLPO.frx":0CEA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(2)=   "Lvwitems"
      Tab(1).Control(3)=   "lvwselecteditems"
      Tab(1).Control(4)=   "cmdremove"
      Tab(1).Control(5)=   "cmdadd"
      Tab(1).ControlCount=   6
      Begin VB.CommandButton cmdNew 
         Caption         =   "ADD"
         Height          =   375
         Left            =   7320
         TabIndex        =   33
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox txtPPrice 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   30
         Top             =   4080
         Width           =   1815
      End
      Begin VB.TextBox txtQnty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   28
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox txtRefNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   27
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox txtIName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   25
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txtPoNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   375
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   420
         Width           =   2295
      End
      Begin VB.ComboBox cbovendors 
         Height          =   315
         ItemData        =   "frmLPO.frx":0D06
         Left            =   4320
         List            =   "frmLPO.frx":0D08
         TabIndex        =   12
         Top             =   1380
         Width           =   2295
      End
      Begin VB.TextBox txtLPOSerial 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   9
         Top             =   2940
         Width           =   1815
      End
      Begin VB.TextBox txtRemarks 
         Height          =   2055
         Left            =   360
         TabIndex        =   8
         Top             =   4500
         Width           =   7575
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "Add"
         Height          =   375
         Left            =   -69960
         TabIndex        =   5
         Top             =   3540
         Width           =   975
      End
      Begin VB.CommandButton cmdremove 
         Caption         =   "Remove"
         Height          =   375
         Left            =   -68760
         TabIndex        =   4
         Top             =   3540
         Width           =   975
      End
      Begin MSComctlLib.ListView lvwselecteditems 
         Height          =   2475
         Left            =   -74880
         TabIndex        =   6
         Top             =   4080
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   4366
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "RefNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Lpo Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Due Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Vendor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Remarks"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView Lvwitems 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   7
         Top             =   660
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   4895
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "RefNo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Lpo Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Due Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Vendor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Remarks"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPduedate 
         Height          =   375
         Left            =   4320
         TabIndex        =   10
         Top             =   2340
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   110231553
         CurrentDate     =   40110
      End
      Begin MSComCtl2.DTPicker DTPlpodate 
         Height          =   375
         Left            =   4320
         TabIndex        =   11
         Top             =   1860
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   46268417
         CurrentDate     =   40110
      End
      Begin VB.Label Label15 
         Caption         =   "Price Per Item"
         Height          =   255
         Left            =   2640
         TabIndex        =   31
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   2640
         TabIndex        =   29
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "Ref. No."
         Height          =   255
         Left            =   2640
         TabIndex        =   26
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Header"
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
         Left            =   240
         TabIndex        =   23
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "PO No"
         Height          =   375
         Left            =   2520
         TabIndex        =   22
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Vendor"
         Height          =   375
         Left            =   2520
         TabIndex        =   21
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "LPO Date"
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         Top             =   1980
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Due Date"
         Height          =   375
         Left            =   2520
         TabIndex        =   19
         Top             =   2460
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "LPO Serial No"
         Height          =   255
         Left            =   2520
         TabIndex        =   18
         Top             =   2940
         Width           =   1095
      End
      Begin VB.Label Label9 
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
         Left            =   360
         TabIndex        =   17
         Top             =   4260
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Available Items"
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
         TabIndex        =   16
         Top             =   420
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Selected Items"
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
         TabIndex        =   15
         Top             =   3660
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Item Name :"
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Label Label16 
      Caption         =   "Total"
      Height          =   255
      Left            =   7200
      TabIndex        =   35
      Top             =   7440
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Purchase Orders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Purchase Orders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmLPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim Mylength As Integer

Private Sub cmdAdd_Click()
On Error GoTo ErrorHandler

If lvwItems.ListItems.Count <= 0 Then
MsgBox "No Items on the list to add", vbInformation, Me.Caption
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

lvwItems.ListItems.Remove (lvwItems.SelectedItem.Index)

Calculate_Total
Exit Sub

ErrorHandler:
MsgBox err.description
End Sub


Private Sub cmdnew_Click()
If Trim(txtPoNo) = "" Then
    MsgBox "Please Capture LPO NO.", vbInformation
        txtPoNo.SetFocus
    Exit Sub
End If

If Trim(txtIName) = "" Then
    MsgBox "Please selected Item Name."
        txtIName.SetFocus
    Exit Sub
End If

If Trim(txtQnty) = "" Then
    MsgBox "Please enter the quantity."
        txtQnty.SetFocus
    Exit Sub
End If

If Trim(txtRefNo) = "" Then
    MsgBox "Please enter the reference number."
        txtRefNo.SetFocus
    Exit Sub
End If

If Trim(txtPPrice) = "" Then
    MsgBox "Please enter price per item."
        txtPPrice.SetFocus
    Exit Sub
End If



If Trim(cbovendors.Text) = "" Then
    MsgBox "Please select vendor."
        cbovendors.SetFocus
    Exit Sub
End If
    

Set li = lvwItems.ListItems.Add(, , txtRefNo)
    li.SubItems(1) = txtIName
    li.SubItems(2) = CInt(txtQnty)
    li.SubItems(3) = CDbl(txtPPrice)
    li.SubItems(4) = Format(DTPlpodate, "dd/mm/yyyy")
    li.SubItems(5) = Format(DTPduedate, "dd/mm/yyyy")
    li.SubItems(6) = cbovendors.Text
    li.SubItems(7) = txtRemarks
   

    txtIName = ""
    txtQnty = ""
    txtPPrice = ""
    txtRemarks = ""
                
                Mylength = CInt(Mid(Trim$(txtRefNo), 5, 10))
            Mylength = Mylength + 1
            txtRefNo = Padding(Mylength)
            txtRefNo = "RQ-" & txtRefNo

End Sub

Private Sub cmdprint_Click()
       On Error GoTo TransError

   Set rs = oSaccoMaster.GetRecordset("SELECT top 1 pno FROM LPO order by pno desc")
    reportname = "LPO.rpt"
    If Not IsNull(rs.Fields(0)) Then
    STRFORMULA = "{lpo.pno}=" & Trim(rs.Fields(0)) & " and Uppercase({Requisition.Status})='ORDERED'"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
    End If
    Exit Sub
TransError:
        
        MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
End Sub

Private Sub cmdRemove_Click()
On Error GoTo ErrorHandler
If LvwselectedItems.ListItems.Count <= 0 Then
MsgBox "No Items on the list to Remove", vbInformation, Me.Caption
Exit Sub
End If
Set li = lvwItems.ListItems.Add(, , LvwselectedItems.SelectedItem)
                        li.SubItems(1) = LvwselectedItems.SelectedItem.ListSubItems(1) & ""
                        li.SubItems(2) = LvwselectedItems.SelectedItem.ListSubItems(2) & ""
                        li.SubItems(3) = LvwselectedItems.SelectedItem.ListSubItems(3) & ""
                        li.SubItems(4) = LvwselectedItems.SelectedItem.ListSubItems(4) & ""
                        li.SubItems(5) = LvwselectedItems.SelectedItem.ListSubItems(5) & ""
                        li.SubItems(6) = LvwselectedItems.SelectedItem.ListSubItems(6) & ""
                        li.SubItems(7) = LvwselectedItems.SelectedItem.ListSubItems(7) & ""

LvwselectedItems.ListItems.Remove (LvwselectedItems.SelectedItem.Index)  '// removes the selected item
Calculate_Total
Exit Sub
ErrorHandler:
MsgBox err.description
End Sub

Private Sub cmSave_Click()
Dim I As Long, SaveLpo As New ADODB.Connection
Dim Refno As String, LPoDate As Date, duedate As Date, itemName As String, QTY As Integer, Pprice As Double, SerialNo As String, _
Vendor As String, REMARKS As String
I = 0

If LvwselectedItems.ListItems.Count <= 0 Then
 MsgBox "No purchase order Raised", vbInformation, Me.Caption
Exit Sub
End If

On Error GoTo TransError
With SaveLpo
    
    .Open SelectedDsn
    .BeginTrans
      

For I = 1 To LvwselectedItems.ListItems.Count
 Set li = LvwselectedItems.ListItems(I)
     Refno = Trim$(li)
     LPoDate = li.ListSubItems(4)
     itemName = li.ListSubItems(1)
     QTY = CDbl(li.ListSubItems(2))
     Pprice = CDbl(li.ListSubItems(3))
     duedate = li.ListSubItems(5)
     SerialNo = Refno
     Vendor = li.ListSubItems(6)
     REMARKS = li.ListSubItems(7)
'd_sp_Requisition @RNo char(35), @TransDate varchar(12), @CostCentre varchar(150), @ServiceReq bit, @IName varchar (150), @Make varchar(150), @Qnty float, @Description varchar(300), @AuditID varchar (50),@pricing money,@Date varchar(12)   AS
sql = ""
sql = sql & "spRequisition '" & Refno & "','" & LPoDate & "','" & cbovendors & " ',0," & QTY * Pprice & ",'" & itemName & "',' '," & QTY & ",'Direct Order','" & user & "'," & Pprice & ",'" & DTPlpodate & "'," & txtPoNo & ""
oSaccoMaster.ExecuteThis (sql)

oSaccoMaster.ExecuteThis ("Update Requisition SET Status='Ordered' WHERE RNo='" & Refno & "'  and pno=" & txtPoNo & "")
    
    ' APPROVE lpo
    sql = ""
       sql = "d_insert_d_Approve '" & Refno & "','0','Order','" & user & "'"
       oSaccoMaster.ExecuteThis (sql)

Next I

'd_sp_LPO @PNo bigint, @TransDate varchar(12),@DueDate varchar(12), @Serial varchar(100),@RefNo varchar(50),@user varchar(35) as
oSaccoMaster.ExecuteThis ("spLPO " & txtPoNo & ",'" & LPoDate & "','" & duedate & "','" & SerialNo & "','" & Refno & "','" & user & "','" & REMARKS & "','" & Vendor & "'")


.CommitTrans

MsgBox "Records saved successfully!"
'Form_Load


LvwselectedItems.ListItems.Clear
Exit Sub
TransError:
        .RollbackTrans
        MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage)
    End With
End Sub
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Form_Load()
txtLPOSerial = ""
txtIName = ""
cbovendors.Text = ""
txtRemarks = ""
txtQnty = ""
txtPPrice = ""

DTPduedate = Format(Get_Server_Date + 7, "dd/mm/yyyy")
DTPlpodate = Format(Get_Server_Date, "dd/mm/yyyy")

'//LOAD THE VENDORS
sql = "SELECT  CompanyName  FROM         ag_Supplier1 order by companyname"
Set rs = oSaccoMaster.GetRecordset(sql)
                Do While Not rs.EOF
                cbovendors.AddItem rs.Fields(0)
                        rs.MoveNext
                    Loop

Set rs = oSaccoMaster.GetRecordset("d_sp_PoNo")
If Not rs.EOF Then
txtPoNo = CCur(rs.Fields(0)) + 1
Else
txtPoNo = "1"
End If
GenerateRefNo




End Sub

Sub GenerateRefNo()
'get the new requition nuo

 mysql = ""
        mysql = "select isnull(Max(Right(RTRIM(RNo),6)),1) as RNo FROM Requisition where Rno like 'RQ-%' "
        
        Set rs2 = oSaccoMaster.GetRecordset(mysql)
        
        If Not rs2.EOF Then
            Mylength = CInt(Mid(rs2!Rno, 5, 10))
            Mylength = Mylength + 1
            txtRefNo = Padding(Mylength)
            txtRefNo = "RQ-" & txtRefNo
        Else
            Mylength = 1
            txtRefNo = "RQ-" & Padding(Mylength)
            
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
                        Price = CDbl(.ListSubItems(3))
                        qnty = CDbl(.ListSubItems(2))
                        amt = Price * qnty
                        Total = Total + amt
                End With
            Next I

        Else
            Total = 0
        End If
    End With
    txttotal = Format(Total, Cfmt)
End Sub
