VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmreturnoutwards 
   BackColor       =   &H00FF8080&
   Caption         =   "Return Outwards"
   ClientHeight    =   7980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   10995
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboVendor 
      Height          =   315
      Left            =   6960
      TabIndex        =   35
      Top             =   1680
      Width           =   3855
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
      Left            =   3240
      TabIndex        =   30
      Top             =   6720
      Width           =   1410
   End
   Begin VB.PictureBox Picture4 
      Height          =   285
      Left            =   4635
      Picture         =   "frmreturnoutwards.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   29
      Top             =   6720
      Width           =   300
   End
   Begin VB.PictureBox Picture21 
      Height          =   285
      Left            =   4725
      Picture         =   "frmreturnoutwards.frx":02C2
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   28
      Top             =   7335
      Width           =   300
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
      Left            =   3240
      TabIndex        =   27
      Top             =   7320
      Width           =   1425
   End
   Begin VB.TextBox txtrno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   15
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox txtamount 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtquantity 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtpcode 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdnextitem 
      Caption         =   "Next item"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      Height          =   360
      Left            =   4080
      Picture         =   "frmreturnoutwards.frx":0584
      ScaleHeight     =   360
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   120
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   17015
      Height          =   360
      Left            =   4080
      Picture         =   "frmreturnoutwards.frx":0706
      ScaleHeight     =   360
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   600
      Width           =   240
   End
   Begin VB.ComboBox cboproductname 
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Top             =   1080
      Width           =   4215
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
      Left            =   0
      TabIndex        =   3
      Top             =   3240
      Value           =   2  'Grayed
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox TXTTOTAL 
      Enabled         =   0   'False
      Height          =   495
      Left            =   8280
      TabIndex        =   1
      Text            =   "0"
      Top             =   5760
      Width           =   2535
   End
   Begin VB.TextBox txtcomment 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   2040
      TabIndex        =   0
      Top             =   8760
      Width           =   6255
   End
   Begin MSComCtl2.DTPicker txtransdate 
      Height          =   375
      Left            =   6480
      TabIndex        =   14
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   148963329
      CurrentDate     =   40265
   End
   Begin MSComctlLib.ListView Lvwitems 
      Height          =   3255
      Left            =   2040
      TabIndex        =   16
      Top             =   3240
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
      MouseIcon       =   "frmreturnoutwards.frx":0888
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
      Left            =   8520
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Caption         =   "Vendor"
      Height          =   255
      Left            =   6240
      TabIndex        =   36
      Top             =   1680
      Width           =   615
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
      Left            =   1440
      TabIndex        =   34
      Top             =   7365
      Width           =   1695
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
      Left            =   4935
      TabIndex        =   33
      Top             =   6720
      Width           =   4170
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
      Left            =   4920
      TabIndex        =   32
      Top             =   7320
      Width           =   4215
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
      Left            =   1440
      TabIndex        =   31
      Top             =   6720
      Width           =   1755
   End
   Begin VB.Label Label2 
      Caption         =   "Receipt No."
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Product Name"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "TransDate"
      Height          =   255
      Left            =   6480
      TabIndex        =   24
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Amount"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Quantity"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Product Code"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblbalance 
      BackColor       =   &H00FF00FF&
      Height          =   255
      Left            =   4800
      TabIndex        =   20
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Balance"
      Height          =   255
      Left            =   3960
      TabIndex        =   19
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label18 
      Caption         =   "TOTAL"
      Height          =   255
      Left            =   8280
      TabIndex        =   18
      Top             =   5400
      Width           =   1575
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
      Left            =   960
      TabIndex        =   17
      Top             =   8880
      Width           =   855
   End
End
Attribute VB_Name = "frmreturnoutwards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboproductname_Change()
        sql = "select p_code, S_No, Qout, pprice from ag_products where p_name='" & cboproductname & "'"
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
cboproductname_Change
End Sub

Private Sub cboproductname_Validate(Cancel As Boolean)
cmdnew_Click

Provider = cn
Set cn = New ADODB.Connection
Dim p As Integer
sql = ""
sql = "select p_code, S_No, Qout, pprice from ag_products where p_name='" & cboproductname & "'"
Set rs = New ADODB.Recordset
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
txtpcode = rs.Fields(0)
lblbalance = rs.Fields(2)
'txtserialno = rs.Fields(1)
txtamount = rs.Fields(3)
End If

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

Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qout from ag_products where p_code='" & txtpcode & "'"
Set rsinstock = oSaccoMaster.GetRecordset(sql)

If rsinstock.Fields(1) <= 0 Then
    MsgBox "Sorry Stock is Zero please re-stock before your proceed", vbInformation
    Exit Sub
End If

Dim piu As Double
piu = rsinstock.Fields(1) - CInt(txtquantity)

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

Do While Not j > (Lvwitems.ListItems.Count)
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

Sub savereturnoutwards()
On Error GoTo syserror
Dim C As String
C = cboVendor.Text
Dim j As Integer

If cboVendor.Text = "" Then
    MsgBox "Please select the Suppliers to Return Stock", vbInformation
        cboVendor.SetFocus
    Exit Sub
End If

If txtcontra = "" Then
    MsgBox "Please enter the Suppliers Creditors GL Account", vbInformation
        txtcontra.SetFocus
    Exit Sub
End If

If TxtOtherPAcc = "" Then
    MsgBox "Please enter the Agrovet Stock Purchase Gl Account", vbInformation
       TxtOtherPAcc.SetFocus
    Exit Sub
End If

If Lvwitems.ListItems.Count = 0 Then
MsgBox "There are no items to be Returned.", vbInformation, Me.Caption
Exit Sub
End If
j = 1

Total = 0
Do While Not j > (Lvwitems.ListItems.Count)
 Lvwitems.ListItems.Item(j).selected = True
 Total = Total + CCur(Lvwitems.SelectedItem.SubItems(4))
j = j + 1
Loop


Startdate = DateSerial(Year(txtransdate), month(txtransdate), 1)
Enddate = DateSerial(Year(txtransdate), month(txtransdate) + 1, 1 - 1)

Set rs = oSaccoMaster.GetRecordset("d_sp_IsClosed '" & Enddate & "'")
If Not rs.EOF Then
    MsgBox "The period " & Enddate & " has been closed by " & rs.Fields(0)
    Exit Sub
End If

If MsgBox("Are you sure You want to post Returned Outwards Goods " & cboVendor.Text, vbQuestion + vbYesNo, _
Me.Caption) = vbNo Then
    Exit Sub
End If
'// check if they are in stock.
For j = 1 To Lvwitems.ListItems.Count
 Lvwitems.ListItems.Item(j).selected = True

Dim rsinstock As Recordset
sql = ""
sql = "select P_CODE,qout from ag_products where p_code='" & Lvwitems.SelectedItem & "'"
Set rsinstock = oSaccoMaster.GetRecordset(sql)
If rsinstock.Fields(1) <= 0 Then
MsgBox "Sorry Stock is Zero for item " & Lvwitems.SelectedItem.SubItems(1) & " please re-stock before your proceed", vbInformation
Exit Sub
End If
Dim Remain As Double
Remain = rsinstock.Fields(1) - CInt(Lvwitems.SelectedItem.SubItems(2))
If Remain < 0 Then
MsgBox "Stock will be negative " & Remain & " please re-stock before you proceed", vbInformation
Exit Sub
End If

sql = ""
sql = sql & "SET dateformat DMY INSERT INTO ag_Receipts(R_No, P_code, T_Date, Amount, Qua, S_Bal, user_id, Cash, SNo,Remarks,Description) VALUES ("
sql = sql & txtrno & ",'" & Lvwitems.SelectedItem & "','" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4)
sql = sql & "," & Lvwitems.SelectedItem.SubItems(2) & "," & Remain & ",'" & user & "'," & Lvwitems.SelectedItem.SubItems(5) & ",'" & C & "','Return Outwards','" & txtcomment & "')"

oSaccoMaster.ExecuteThis (sql)
oSaccoMaster.ExecuteThis ("Update ag_Products SET Qout =" & CCur(Remain) & " WHERE p_code= '" & Lvwitems.SelectedItem & "'")

'\\ save to gl
    sql = ""
    sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & txtransdate & "'," & Lvwitems.SelectedItem.SubItems(4) & ",'" & txtcontra & "','" & TxtOtherPAcc & "','" & Lvwitems.SelectedItem & "','" & Lvwitems.SelectedItem.SubItems(1) & "' ,'Return Outwards','" & user & "',1,0)"
    oSaccoMaster.ExecuteThis (sql) '
Next j


If chkPrint.Value = vbChecked Then
PrintReceipt
PrintReceipt
End If

Lvwitems.ListItems.Clear
txtrno = ""
txtcomment = ""
txtpcode.Text = ""
txtquantity = 1
txtamount = ""
cboVendor.Text = ""
MsgBox "Record saved Successfully"
Exit Sub
syserror:
MsgBox err.description & " error occured."
End Sub

Private Sub cmdsave_Click()
 savereturnoutwards
End Sub

Private Sub Form_Load()
txtransdate = Format(Get_Server_Date, "dd/mm/yyyy")
Provider = "MAZIWA"
sql = "select P_NAME  from ag_products ORDER BY P_NAME ASC"
Set rs = New ADODB.Recordset
Set rs = oSaccoMaster.GetRecordset(sql)
While Not rs.EOF
cboproductname.AddItem rs.Fields(0)
rs.MoveNext
Wend

cboVendor.Clear
Set rs = oSaccoMaster.GetRecordset("SELECT  CompanyName  FROM ag_Supplier1 order by companyname")
While Not rs.EOF
cboVendor.AddItem rs.Fields(0)
rs.MoveNext
Wend

cboproductname.Enabled = True
chkPrint.Value = vbChecked
End Sub

Private Sub txtpcode_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierid from ag_products where p_code='" & txtpcode & "'"
Set rs = oSaccoMaster.GetRecordset(sql)
If Not rs.EOF Then
 
If Not IsNull(rs.Fields(1)) Then cboproductname = (rs.Fields(1))
End If
End If

End Sub



Private Sub txtquantity_Validate(Cancel As Boolean)
If Not IsNumeric(txtquantity) Then
MsgBox "Enter a value please", vbCritical
txtquantity = ""
txtquantity.SetFocus
Exit Sub
End If
End Sub
Private Sub PrintReceipt()
    On Error GoTo syserror
    Dim strReceipts As String
    Dim pay, tot, disc As Currency
    Dim Z, X As Integer
    Dim a As Integer
    Dim b As Integer
    Dim mode As String
    
    mode = "RETURN OUTWARDS"
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
    Printer.Print Tab(0); "Powered By EasyMa Amtech Technologies"
    Printer.Print Tab(2); "******************************************************"
    Printer.Print
    Printer.EndDoc

    Exit Sub
syserror:
    MsgBox err.description, vbInformation
End Sub
Private Sub txtcontra_Change()
 On Error GoTo syserror
    Dim Account As Acc_Details
    Editing = True
    Account = Get_Acc_Details(txtcontra, ErrorMessage)
    If Account.Accno <> "" Then
        lblcontra = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblcontra = ""
    End If
    Exit Sub
syserror:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub TxtOtherPAcc_Change()
On Error GoTo syserror
    Dim Account As Acc_Details
    Editing = True
    Account = Get_Acc_Details(TxtOtherPAcc, ErrorMessage)
    If Account.Accno <> "" Then
        lblOtherPaymentAcc = Account.AccName
    Else
        If ErrorMessage <> "" Then
            MsgBox ErrorMessage, vbInformation, Me.Caption
            ErrorMessage = ""
        End If
        lblOtherPaymentAcc = ""
    End If
    Exit Sub
syserror:
    MsgBox err.description, vbInformation, Me.Caption
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
