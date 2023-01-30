VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmProductSales 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MILK DISPATCH"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   13635
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtdr 
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
      Left            =   7320
      TabIndex        =   40
      Top             =   1200
      Width           =   1170
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
      Left            =   1440
      TabIndex        =   37
      Top             =   4680
      Width           =   1410
   End
   Begin VB.PictureBox Picture4 
      Height          =   285
      Left            =   2880
      Picture         =   "frmProductSales.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   36
      Top             =   4680
      Width           =   300
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
      Left            =   14400
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   35
      Top             =   6000
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.TextBox txtAvailable 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   405
      Left            =   6720
      TabIndex        =   20
      Text            =   "0"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtUnDispatch 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   6600
      TabIndex        =   19
      Text            =   "0"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox txtReceiptno 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   405
      Left            =   4680
      TabIndex        =   17
      Top             =   120
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customer Details"
      Height          =   1575
      Left            =   7080
      TabIndex        =   12
      Top             =   120
      Width           =   6375
      Begin VB.ComboBox cboDCode 
         Height          =   315
         ItemData        =   "frmProductSales.frx":02C2
         Left            =   240
         List            =   "frmProductSales.frx":02C4
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtNames 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label lblGl 
         Caption         =   "Gl Control Acc:"
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Top             =   1080
         Width           =   4695
      End
   End
   Begin VB.TextBox txtActualBalance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   405
      Left            =   6720
      TabIndex        =   11
      Text            =   "0"
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox txtDActuals 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   14520
      TabIndex        =   10
      Text            =   "0"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cboPricing 
      Height          =   315
      ItemData        =   "frmProductSales.frx":02C6
      Left            =   1440
      List            =   "frmProductSales.frx":02D3
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox txtPAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   405
      Left            =   3120
      TabIndex        =   8
      Text            =   "0"
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdPush 
      Caption         =   ">>"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   5880
      Width           =   735
   End
   Begin VB.TextBox txtTotalSales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Enabled         =   0   'False
      Height          =   405
      Left            =   11880
      TabIndex        =   6
      Text            =   "0"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.ComboBox cboPMode 
      Height          =   315
      ItemData        =   "frmProductSales.frx":02F1
      Left            =   11880
      List            =   "frmProductSales.frx":02FB
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "<<"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "POST!"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   11400
      TabIndex        =   3
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txtComment 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   9840
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3120
      Width           =   3615
   End
   Begin VB.TextBox txtdeposits 
      Height          =   495
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0"
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox txtbalance 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0"
      Top             =   4560
      Width           =   2775
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   0
      TabIndex        =   16
      Top             =   960
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "BranchCode"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "QSupplied"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Packets available"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CValue"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ProductUnit"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "WSPrice"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "AgPrice"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "RetPrice"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "GL Account"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPDispatchDate 
      Height          =   375
      Left            =   5520
      TabIndex        =   18
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   122290177
      CurrentDate     =   40105
   End
   Begin MSComctlLib.ListView lvwOrder 
      Height          =   2535
      Left            =   0
      TabIndex        =   21
      Top             =   6360
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "BranchCode"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "BranchCode"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Unit"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Quantity"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "UnitPrice"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "GlAccNO"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   " Milk Sales Acc"
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
      Left            =   0
      TabIndex        =   39
      Top             =   4680
      Width           =   1095
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
      Left            =   3255
      TabIndex        =   38
      Top             =   4680
      Width           =   4170
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "Sale Date :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3240
      TabIndex        =   34
      Top             =   600
      Width           =   990
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   " litres (Ltrs) for sales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3960
      TabIndex        =   33
      Top             =   3675
      Width           =   2265
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "Qty/Litres sold"
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
      Height          =   240
      Left            =   4560
      TabIndex        =   32
      Top             =   5280
      Width           =   1515
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "Invoice No. :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3240
      TabIndex        =   31
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "Actual litres (Ltrs)"
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
      Height          =   240
      Left            =   4440
      TabIndex        =   30
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label lblProductName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   29
      Top             =   3495
      Width           =   3735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "No of Crates"
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
      Height          =   240
      Left            =   14520
      TabIndex        =   28
      Top             =   6480
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Pricing Type"
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "Total Sales(Sh)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   10320
      TabIndex        =   26
      Top             =   1995
      Width           =   1395
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Payment Mode"
      Height          =   375
      Left            =   10200
      TabIndex        =   25
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Comment"
      Height          =   255
      Left            =   8640
      TabIndex        =   24
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label LBLBALANCE 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Deposits"
      Height          =   255
      Left            =   9600
      TabIndex        =   23
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label lblbal 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Balance"
      Height          =   375
      Left            =   9600
      TabIndex        =   22
      Top             =   4680
      Width           =   975
   End
End
Attribute VB_Name = "frmProductSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Price As Currency
Dim capp As Integer
Dim crate As Double
Dim sprice As Double
Dim DRaccno As String
Dim Crcess As String
Dim Drcess As String
Dim Cess As Double
Dim Craccno As String
Dim flavId As Integer
Dim productCode As String
Dim ProdUnit As String
Dim WSPrice As Double, AgPrice As Double, RtPrice As Double
Dim TotalSales As Double
Dim UnitValue As Double
Dim Kilos As Double
Dim productGl As String
Dim MSalesAcc  As String

Public Function NewReceipt() As String
    
    Dim Rno As String
    Dim thisDay As Date
    Dim rcount As Double
    Set Rst = oSaccoMaster.GetRecordset("select count(distinct OrderNo)ccount from salesorder")
    If Not Rst.EOF Then
        rcount = Rst(0) + 3
    Else
        rcount = 3
    End If
    
    'thisday = Get_Server_Date
    
    Rno = Format(CStr(rcount), "000000")
    
    NewReceipt = "INV" & "-" & CStr(Rno)
End Function
Private Sub cboDCode_Change()
    Set Rst = oSaccoMaster.GetRecordset("select p.dname,p.accdr,p.accCr,p.drcess, p.crcess, drcess, crcess,isnull(P.crate,0)crate,p.price,isnull(gl.glaccname,'')GlName " _
    & " from d_debtors p left outer join glSetup gl on p.accdr=gl.accno " _
    & " where p.dcode='" & cboDCode & "'")
    If Not Rst.EOF Then
        txtNames.Text = Rst("dname")
        DRaccno = Rst("accdr")
        lblGl = Rst("GlName")
        txtdr = DRaccno
        Cess = Rst("crate")
        Drcess = IIf(IsNull(Rst("drcess")), "", Rst("drcess"))
        Crcess = IIf(IsNull(Rst("crcess")), "", Rst("crcess"))
        txtcontra = IIf(IsNull(Rst("accCr")), "", Rst("accCr"))
        txtPAmount = IIf(IsNull(Rst("Price")), 0, Rst("price"))
    Else
        txtNames.Text = ""
        DRaccno = ""
        lblGl = ""
        txtdr = ""
        txtPAmount = 0
    End If
'    ''Depositz
Set Rst1 = oSaccoMaster.GetRecordset("select  email from d_Debtors where DCode='" & cboDCode & "'")
If Not Rst.EOF Then
'txtdeposits.Text = IIf(IsNull(Rst1!Email), 0, Rst1!Email)
'txtinvoice.Text = IIf(IsNull(Rst1!DocumentNo), 0, Rst1!DocumentNo)
Else
'txtdeposits.Text = ""
End If
'''glbalance
Set rst2 = oSaccoMaster.GetRecordset("select  isnull(CurrentBal,0) as CurrentBal from glSetup where Accno='" & DRaccno & "'")
If Not rst2.EOF Then
txtBalance.Text = IIf(IsNull(rst2!CurrentBal), 0, rst2!CurrentBal)
'txtinvoice.Text = IIf(IsNull(Rst1!DocumentNo), 0, Rst1!DocumentNo)
Else
txtBalance.Text = ""
End If
cboPricing_Change

End Sub

Private Sub cboDCode_Click()
    cboDCode_Change
End Sub

Private Sub cboPMode_Change()
txtcomment.Text = cboPMode.Text
End Sub

Private Sub cboPMode_Click()
cboPMode_Change
End Sub

Private Sub cboPricing_Change()
    Select Case cboPricing.Text
        Case "Wholesale"
            txtPAmount = WSPrice
        Case "Agent"
            txtPAmount = AgPrice
        Case "Retail"
            txtPAmount = RtPrice
    End Select
    
    Set Rst = oSaccoMaster.GetRecordset("select price from d_Debtors where dcode='" & cboDCode & "'")
    If Not Rst.EOF Then
    txtPAmount = IIf(IsNull(Rst.Fields(0)), 35, Rst.Fields(0))
    End If
    
End Sub

Private Sub cboPricing_Click()
    cboPricing_Change
End Sub

Private Sub cmdPush_Click()
    On Error GoTo Capture
    
    If txtUnDispatch = "" Then
        MsgBox "Please enter the dispatch quantity."
        txtUnDispatch.SetFocus
        Exit Sub
    ElseIf CDbl(txtUnDispatch.Text) > CDbl(txtAvailable) Then
       If MsgBox("The stated quantity (Units) is greater than the available ,Are you Sure You Want To Continue", vbCritical + vbYesNo) = vbNo Then
        Exit Sub
       End If
    ElseIf CDbl(txtDActuals.Text) > CDbl(txtActualBalance) Then
        MsgBox "The stated quantity (Units) is greater than the available ,Are you Sure You Want To Continue", vbCritical + vbYesNo
        Exit Sub
    ElseIf CDbl(txtUnDispatch) = 0 Then
        MsgBox "Cannot be dispatching 0 (Zero) litres!", vbCritical + vbOKOnly
        Exit Sub
    ElseIf CDbl(txtPAmount.Text) = 0 Then
        MsgBox "Choose the unit price for the item!", vbCritical + vbOKOnly
        Exit Sub
     ElseIf txtcontra.Text = "" Then
        MsgBox "Choose milk sale account!", vbCritical + vbOKOnly
        txtcontra.SetFocus
        Exit Sub
    End If
    
    'productGl = ""
    With lvwOrder
    
        For I = 1 To .ListItems.Count
            If productCode = .ListItems(I).Text And ProdUnit = .ListItems(I).ListSubItems(1) Then
                MsgBox "Then item is already in the list", vbCritical
                Exit Sub
            End If
        Next I
    
        Set li = lvwOrder.ListItems.Add(, , productCode)
        li.ListSubItems.Add , , ProdUnit
        li.ListSubItems.Add , , ProdUnit
        li.ListSubItems.Add , , txtUnDispatch
        li.ListSubItems.Add , , txtPAmount
        li.ListSubItems.Add , , txtcontra
    End With
    
    txtUnDispatch.Text = 0
    
    Recalculate
    
    Exit Sub
Capture:
    
End Sub

Private Sub cmdremove_Click()
    With lvwOrder
       If .ListItems.Count = 0 Then Exit Sub
       
        If MsgBox("Remove the selected Item from the list?", vbQuestion + vbYesNo, "CONFIRM ACTION") = vbNo Then
            Exit Sub
        End If
       
       .ListItems.Remove .SelectedItem.Index
        Recalculate
    End With
End Sub

Private Sub cmdsave_Click()
    On Error GoTo Capture
    Dim bcode As String
     Dim Rst As New ADODB.Recordset
    Dim DispatchTrans As ADODB.Connection
    If cboDCode = "" Then
        MsgBox "Debtors code cannot be blank; input an existing one", vbCritical
        Exit Sub
    End If
    If txtcomment = "" Then
        MsgBox "Please enter the Comment/details to reference paymode."
        txtcomment.SetFocus
        Exit Sub
    End If
    If cboPMode = "" Then
        MsgBox "Please enter the payment mode."
        cboPMode.SetFocus
        Exit Sub
    End If
    
    If txtPAmount = "" Or txtPAmount = 0 Then
        MsgBox "Please Update  the Customer Buying Price/details first.", vbInformation
        txtPAmount.SetFocus
        Exit Sub
    End If
     ''// check if the receipt no
        Set Rst = Nothing
         Set Rst = oSaccoMaster.GetRecordset("select *  from SalesOrder  where OrderNo='" & txtReceiptno & "' ")
           If Not Rst.EOF Then
           
               txtReceiptno.Text = NewReceipt
             Else
           
           End If
    If txtUnDispatch = "" Then
        MsgBox "Please enter the dispatch quantity."
        txtUnDispatch.SetFocus
        Exit Sub
    ElseIf CDbl(txtUnDispatch.Text) > CDbl(txtAvailable) Then
'       If MsgBox("The stated quantity (Units) is greater than the available ,Are you Sure You Want To Continue", vbCritical + vbYesNo) = vbNo Then
'         Exit Sub
'       End If
    ElseIf CDbl(txtDActuals.Text) > CDbl(txtActualBalance) Then
'       If MsgBox("The stated quantity (Units) is greater than the available,Are you Sure You Want To Continue", vbCritical + vbYesNo) = vbNo Then
'        Exit Sub
'      End If
    ElseIf lblGl.Caption = "" Then
        MsgBox "The Customer control account is not set, kindly consult the management", vbCritical
        Exit Sub
    End If
     ''CONTROLLING BALANCES
     Dim Y As Double
     Dim X As Double
     Dim Z As Double
     If txtBalance = "" Then
     txtBalance = 0
     End If
     Y = txtBalance
     X = txtTotalSales
     Z = Y + X
     Startdate = DateSerial(year(DTPDispatchDate), month(DTPDispatchDate), 1)
     Enddate = DateSerial(year(DTPDispatchDate), month(DTPDispatchDate) + 1, 1 - 1)
'
'      If txtdeposits.Text < Z And cboPMode = "Credit" Then
''        MsgBox "Customer Deposits Cannot Quarantee for the Above Sales", vbCritical + vbYesNo
''        Exit Sub
'''        End If
''ElseIf vbYes Then
'''        Else
''        frmAuthorize.Show vbModal
''        If UCase(User) = UCase(Authority) Or (UCase(Authority) = "ADMIN") Or Authority = "" Then
'            MsgBox "Customer Deposits Cannot Quarantee for the Above Sales", vbOKOnly + vbExclamation
'           Exit Sub
'        End If
'    End If


    If MsgBox(txtPAmount & "-" & "Is the Current Milk Buying Price for:" & "-" & txtNames, vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    If MsgBox("Commit the transaction?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    Set DispatchTrans = New ADODB.Connection
    DispatchTrans.Open oSaccoMaster.goConn, "atm", "atm"
    DispatchTrans.BeginTrans
    
        NewTransaction txtTotalSales, DTPDispatchDate, "Milk Sales"
        
        With lvwOrder
            For I = 1 To .ListItems.Count
            bcode = .ListItems(I).Text
                sql = "Insert into OrderDetails(Orderno,Dcode,item,Quantity,UnitPrice,Transactionno,Crates,Dispdate,Bcode)" _
                & " Values('" & txtReceiptno & "','" & cboDCode & "','" & .ListItems(I).ListSubItems(1) & "'," & .ListItems(I).ListSubItems(3) & "," & .ListItems(I).ListSubItems(4) & ",'" & transactionNo & "','" & txtDActuals & "','" & DTPDispatchDate & "','" & bcode & "')"
                 oSaccoMaster.ExecuteThis (sql)
'                    If Not oSaccoMaster.ExecuteThis(sql) Then
'                        GoTo TransError
'                    End If
                'For now, let the Goods Be sold!
                sql = "Insert into productflow(prodId,unit,Quantity,transdate,batchno,auditid,remarks,transactionno,stage,Bcode)" _
                & " Values('" & .ListItems(I).Text & "','" & .ListItems(I).ListSubItems(4) & "'," & .ListItems(I).ListSubItems(3) * (-1) & ",'" & DTPDispatchDate & "','" & txtReceiptno & "','" & User & "','" & txtcomment & "','" & transactionNo & "','3','" & bcode & "')"
                 oSaccoMaster.ExecuteThis (sql)
'                If Not oSaccoMaster.ExecuteThis(sql) Then
'                    GoTo TransError
'                End If
                'Gl Now
                Dim qnt As Double, sprice, CessAmt As Double
                
                qnt = .ListItems(I).ListSubItems(3)
                sprice = .ListItems(I).ListSubItems(4)
                MSalesAcc = .ListItems(I).ListSubItems(5)
                DRaccno = Trim(txtdr)
                CessAmt = qnt * Cess
'                If Not SaveGLTRANSACTION(DTPDispatchDate.value, qnt * sprice, DRaccno, MSalesAcc, txtReceiptno, cboDCode, "Milk Sales -" & .ListItems(I).ListSubItems(2) & "-" & txtNames.Text, user, transactionNo) Then
'                    GoTo Capture
'                End If
                
                 If Not Save_GLTRANSACTION(DTPDispatchDate.Value, qnt * sprice, DRaccno, MSalesAcc, txtReceiptno, _
                    cboDCode, User, ErrorMessage, "Milk Sales -" & .ListItems(I).ListSubItems(2) & "-" & txtNames.Text, 1, 1, txtReceiptno, transactionNo, "") Then
                        GoTo Capture
                        
                     End If
                If CessAmt > 0 Then
                     
                     If Not Save_GLTRANSACTION(DTPDispatchDate.Value, CessAmt, Drcess, Crcess, txtReceiptno, _
                    cboDCode, User, ErrorMessage, "Milk Sales cess-" & .ListItems(I).ListSubItems(2) & "-" & txtNames.Text, 1, 1, txtReceiptno, transactionNo, "") Then
                        GoTo Capture
                        
                     End If
                
                End If
                     
            Next I
        End With
        
        'save the summary sales order
          'Dim balance As Double
          'balance = CDbl(txtbalance.Text)
          If txtBalance = "" Then txtBalance = 0
        sql = "set dateformat dmy insert into SalesOrder (OrderNo,Dcode,orderDate,OrderAmount,Balance,Auditid,Remarks,Transactionno,bcode) " _
        & " Values ('" & txtReceiptno & "','" & cboDCode & "','" & DTPDispatchDate.Value & "'," & txtTotalSales & "," & CDbl(txtBalance.Text) & ",'" & User & "','" & txtcomment & "','" & transactionNo & "','" & bcode & "')"
        'balance = 0
         oSaccoMaster.ExecuteThis (sql)
'        If Not oSaccoMaster.ExecuteThis(sql) Then
'            GoTo Capture
'        End If
        
        'Statement
        sql = "insert into CustomerStmt (invId,TransDate,Refno,Amount,TransType,Balance,Auditid,Remarks,Transactionno,Dcode,Bcode) " _
        & " Values ('" & txtReceiptno & "','" & DTPDispatchDate.Value & "','" & txtReceiptno & "',0,'DR'," & txtTotalSales & ",'" & User & "','Product Sold and Invoiced','" & transactionNo & "','" & cboDCode & "','" & bcode & "')"
         oSaccoMaster.ExecuteThis (sql)
         
         sql = "set dateformat dmy insert into Printmilkinvoice(InvoiceNo, DCode, DispQnty, Total,Cess,Transdate, StartDate, EndDate) " _
              & " values('" & txtReceiptno & "','" & cboDCode & "'," & qnt & "," & qnt * sprice & "," & CessAmt & ",'" & DTPDispatchDate.Value & "','" & Startdate & "' ,'" & Enddate & "')"
        oSaccoMaster.ExecuteThis (sql)
    
    DispatchTrans.CommitTrans
    MsgBox "Record Saved Successfully"
    STRFORMULA = "{PrintMilkinvoice.InvoiceNo}='" & txtReceiptno & "'"
    reportname = "Printmilkinvoice.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
''     ////
'  STRFORMULA = "{Salesorder.OrderNo}='" & txtReceiptno & "'"
'    reportname = "deliverynote.rpt"
'    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
'    txtbalance.Text = 0
'    txtdeposits.Text = 0
    Form_Load
Exit Sub
TransError:
    DispatchTrans.RollbackTrans
    MsgBox IIf(ErrorMessage = "", err.Description, ErrorMessage)
    Exit Sub
Capture:
    MsgBox err.Description
End Sub

Private Sub Command1_Click()

End Sub

Private Sub dbalance_Click()
STRFORMULA = ""
    reportname = "DEBTORBALANCE.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
End Sub

Private Sub DTPDispatchDate_Change()
  Martim
End Sub

Private Sub Form_Load()
Dim Kilos As Double
    DTPDispatchDate = Format(Get_Server_Date, "dd/mm/yyyy")
    DTPDispatchDate.MaxDate = Format(Get_Server_Date, "dd/mm/yyyy")
    productCode = ""
    ProdUnit = ""
    txtUnDispatch.Text = 0
    txtReceiptno.Text = NewReceipt
    ListView1.ListItems.Clear
    lvwOrder.ListItems.Clear
    txtTotalSales.Text = 0#
    Martim
    cboDCode.Clear
    cboDCode.AddItem ""
    Set rs = oSaccoMaster.GetRecordset("SELECT dcode from d_debtors order by id asc")
    If Not rs.EOF Then
        With rs
            While Not .EOF
             cboDCode.AddItem rs.Fields("dcode")
             .MoveNext
            Wend
        End With
'        cboDCode.Text = cboDCode.List(0)
    End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    UnitValue = 0
    txtAvailable.Text = Item.ListSubItems(1)
    'txtBalance.Text = Item.ListSubItems(2)
'    lblProductName.Caption = Item.ListSubItems(1)
    UnitValue = 1
    txtActualBalance.Text = Item.ListSubItems(1) '(CDbl(txtAvailable) * CDbl(UnitValue)) / 1000
    productCode = Item.Text
    
    Set Rst = oSaccoMaster.GetRecordset("select  bname from d_branch where bcode='" & productCode & "'")
    If Not Rst.EOF Then
     lblProductName.Caption = Rst.Fields(0)
     Else
     lblProductName = ""
    End If
    
'    If productCode = "0" Then
'    lblProductName.Caption = "SALIENT KIPKAREN"
'    ElseIf productCode = "1" Then
'    lblProductName.Caption = "LEMOOK"
'    ElseIf productCode = "2" Then
'    lblProductName.Caption = "SURUNGAI"
'    ElseIf productCode = "3" Then
'    lblProductName.Caption = "SANGALO"
'    ElseIf productCode = "4" Then
'    lblProductName.Caption = "KAPKOROS"
'    ElseIf productCode = "5" Then
'    lblProductName.Caption = "KAPTEL"
'    ElseIf productCode = "6" Then
'    lblProductName.Caption = "NDAPTABWA"
'    Else
'    lblProductName.Caption = ""
'    End If
    
    
    
    
    ProdUnit = 0
    WSPrice = 20
    AgPrice = 20
    RtPrice = 20
    
   ' productGl = Item.ListSubItems(8)
'    cboPricing_Change
End Sub

Private Sub Recalculate()
    Dim qnt As Double, sprice As Double
    TotalSales = 0
    With lvwOrder
        For I = 1 To .ListItems.Count
            qnt = .ListItems(I).ListSubItems(3)
            sprice = .ListItems(I).ListSubItems(4)
            TotalSales = TotalSales + (qnt * sprice)
        Next I
        txtTotalSales.Text = TotalSales
    End With
End Sub



Private Sub lvwOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        cmdremove_Click
    End If
End Sub

Private Sub mnudeliver_Click()
    'Show_Sales_Crystal_Report "", reportname, CompanyName
 STRFORMULA = ""
    reportname = "deliverynote.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName

End Sub

Private Sub mnudeposits_Click()
STRFORMULA = ""
    reportname = "deposits.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
End Sub

Private Sub mnuPerSales_Click()
    STRFORMULA = ""
    reportname = "PeriodicSales.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
End Sub

Private Sub MNUSALERET_Click()
STRFORMULA = ""
    reportname = "salesreturnz.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, CompanyName
End Sub

Private Sub mnusalesinvoicesreport_Click()

    'STRFORMULA = "{Salesorder.OrderNo}='" & txtReceiptno & "'"
    STRFORMULA = ""
    reportname = "salesinvoice1.rpt"
     Show_Sales_Crystal_Report "", reportname, CompanyName

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
     Set rs = Nothing
     Set rs = oSaccoMaster.GetRecordset("select GlAccName from   GLSETUP where AccNo='" & txtcontra & "'")
       If Not rs.EOF Then
       lblcontra = IIf(IsNull(rs("GlAccName")), "", rs("GlAccName"))
       Else
       lblcontra = ""
       End If
End Sub

Private Sub txtcontra_Click()
txtcontra_Change
End Sub

Private Sub txtDActuals_Change()
'On Error GoTo Capture
    Dim X As Double
    Dim Y As Double
    X = 50
   If lblProductName = "PASTURIZED MILK(ROYAL)(200ML)" Then
   txtUnDispatch.Text = (CDbl(txtDActuals) * 50)
   Else
   If lblProductName = "PASTURIZED MILK(ROYAL)(500ml)" Then
 txtUnDispatch.Text = (CDbl(txtDActuals) * 20)
'Capture:
'    txtUnDispatch.Text = 0
    End If
    End If
    
End Sub

Private Sub txtDActuals_KeyPress(KeyAscii As Integer)
'    If keyIsValid(KeyAscii, 1) = False Then
'        Beep
'        KeyAscii = 0
'    End If
End Sub

Private Sub txtUnDispatch_Change()
'   On Error GoTo Capture
   If lblProductName = "PASTURIZED MILK(ROYAL)(200ML)" Then
   txtDActuals.Text = (CDbl(txtUnDispatch) / 50)
   Else
   If lblProductName = "PASTURIZED MILK(ROYAL)(500ml)" Then
   txtDActuals = (CDbl(txtUnDispatch) / 20)
   End If
    Exit Sub
'Capture:
    txtDActuals.Text = 0
    End If
End Sub

Private Sub txtUnDispatch_KeyPress(KeyAscii As Integer)
'    If keyIsValid(KeyAscii, 1) = False Then
'        Beep
'        KeyAscii = 0
'    End If
'
'    If KeyAscii = 13 Then
'        cmdPush_Click
'    End If
    
End Sub


Private Sub Martim()

Dim sales, Rejects, Cfa, Spillage, Fromstation, Tostation, Bf As Double
 sql = "set dateformat dmy SELECT  BranchCode, SUM(QSupplied) QSupplied FROM  d_Milkintake where TransDate='" & DTPDispatchDate & "' GROUP BY BranchCode"
    Set Rst = oSaccoMaster.GetRecordset(sql)
ListView1.ListItems.Clear
    While Not Rst.EOF
        Set li = ListView1.ListItems.Add(, , Rst("BranchCode"))
         sales = getsales(Rst("BranchCode"), DTPDispatchDate)
         Rejects = getRejects(Rst("BranchCode"), DTPDispatchDate)
         Cfa = getCarryF(Rst("BranchCode"), DTPDispatchDate)
         Spillage = getSpillage(Rst("BranchCode"), DTPDispatchDate)
         Tostation = getTostation(Rst("BranchCode"), DTPDispatchDate)
         Fromstation = getFromstation(Rst("BranchCode"), DTPDispatchDate)
         Bf = getBf(Rst("BranchCode"), DTPDispatchDate)
         
        Kilos = IIf(IsNull(Rst("QSupplied")), 0, Rst("QSupplied")) + Fromstation + Bf - (sales + Cfa + Rejects + Spillage + Tostation)
        'li.ListSubItems.Add , , rst("QSupplied") & "(" & rst("unit") & ")"
        li.ListSubItems.Add , , Kilos
'        li.ListSubItems.Add , , rst("CValue")
'        li.ListSubItems.Add , , rst("Unit")
'        li.ListSubItems.Add , , rst("WholesalePrice")
'        li.ListSubItems.Add , , rst("AgentPrice")
'        li.ListSubItems.Add , , rst("RetailPrice")
'        li.ListSubItems.Add , , rst("accno")
        Rst.MoveNext
    Wend
End Sub


Public Function getsales(bcode As String, ddate As Date) As Double
 Dim rssales As New ADODB.Recordset
 Set rssales = oSaccoMaster.GetRecordset("  set dateformat dmy Select isnull(sum(Quantity),0) as Quantity from OrderDetails where   Bcode='" & bcode & "' and Dispdate= '" & ddate & "'")
    If Not rssales.EOF Then
    getsales = IIf(IsNull(rssales(0)), 0, rssales(0))
    Else
    getsales = 0
    End If
   
End Function
Public Function getRejects(bcode As String, ddate As Date) As Double
 Dim rsrejects As New ADODB.Recordset
 Set rsrejects = oSaccoMaster.GetRecordset("  set dateformat dmy Select isnull(sum(Reject),0) as Rejects from milkcontrol2 where   Bcode='" & bcode & "' and Transdate= '" & ddate & "'")
    If Not rsrejects.EOF Then
     getRejects = IIf(IsNull(rsrejects(0)), 0, rsrejects(0))
    Else
     getRejects = 0
    End If
   
End Function
Public Function getCarryF(bcode As String, ddate As Date) As Double
 Dim rsCarryF As New ADODB.Recordset
 Set rsCarryF = oSaccoMaster.GetRecordset("  set dateformat dmy Select isnull(sum(cfa),0) as Cfa from milkcontrol2 where   Bcode='" & bcode & "' and Transdate= '" & ddate & "'")
    If Not rsCarryF.EOF Then
     getCarryF = IIf(IsNull(rsCarryF(0)), 0, rsCarryF(0))
    Else
     getCarryF = 0
    End If
   
End Function
Public Function getSpillage(bcode As String, ddate As Date) As Double
 Dim rsgetspill As New ADODB.Recordset
 Set rsgetspill = oSaccoMaster.GetRecordset("  set dateformat dmy Select isnull(sum(Spillage),0) as spillage from milkcontrol2 where   Bcode='" & bcode & "' and Transdate= '" & ddate & "'")
    If Not rsgetspill.EOF Then
     getSpillage = IIf(IsNull(rsgetspill(0)), 0, rsgetspill(0))
    Else
     getSpillage = 0
    End If
   
End Function
Public Function getFromstation(bcode As String, ddate As Date) As Double
 Dim rsgetfrom As New ADODB.Recordset
 Set rsgetfrom = oSaccoMaster.GetRecordset("  set dateformat dmy Select isnull(sum(fromStation),0) as Fromstation from Milktransfer where   toBranch='" & bcode & "' and Transdate= '" & ddate & "'")
    If Not rsgetfrom.EOF Then
     getFromstation = IIf(IsNull(rsgetfrom(0)), 0, rsgetfrom(0))
    Else
     getFromstation = 0
    End If
   
End Function
Public Function getTostation(bcode As String, ddate As Date) As Double
 Dim rsgetto As New ADODB.Recordset
 Set rsgetto = oSaccoMaster.GetRecordset("  set dateformat dmy Select isnull(sum(fromStation),0) as Tostation from Milktransfer where   fromBranch='" & bcode & "' and Transdate= '" & ddate & "'")
    If Not rsgetto.EOF Then
     getTostation = IIf(IsNull(rsgetto(0)), 0, rsgetto(0))
    Else
     getTostation = 0
    End If
   
End Function
Public Function getBf(bcode As String, ddate As Date) As Double
 Dim rsBf As New ADODB.Recordset
 Dim yday As Date
 yday = DateSerial(year(ddate), month(ddate), Day(ddate) - 1)
 Set rsBf = oSaccoMaster.GetRecordset("  set dateformat dmy Select isnull(sum(cfa),0) as Bf from milkcontrol2 where   Bcode='" & bcode & "' and Transdate= '" & yday & "'")
    If Not rsBf.EOF Then
     getBf = IIf(IsNull(rsBf(0)), 0, rsBf(0))
    Else
     getBf = 0
    End If
   
End Function


