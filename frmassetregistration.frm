VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmassetregistration 
   BackColor       =   &H00C0C000&
   Caption         =   "Asset Registration"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6975
   LinkTopic       =   "Form5"
   ScaleHeight     =   4230
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   285
      Left            =   3120
      Picture         =   "frmassetregistration.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2160
      Width           =   285
   End
   Begin VB.TextBox txtserialno 
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
      Left            =   4920
      TabIndex        =   20
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtcurrentvalue 
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
      Left            =   3360
      TabIndex        =   18
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txtpurchaseamt 
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
      Left            =   1800
      TabIndex        =   16
      Top             =   3120
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   118095873
      CurrentDate     =   41757
   End
   Begin VB.CommandButton cmdSearchGL 
      Height          =   285
      Left            =   3120
      Picture         =   "frmassetregistration.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1320
      Width           =   285
   End
   Begin VB.CommandButton cmdsvae 
      Caption         =   "Save"
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "New"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtdeprate 
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
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtassetname 
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
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox txtassetcode 
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
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblbankacc 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1920
      TabIndex        =   26
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblbankaccname 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3480
      TabIndex        =   25
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label11 
      Caption         =   "Accno"
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
      Left            =   1920
      TabIndex        =   24
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Depreciation Account"
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
      Left            =   3480
      TabIndex        =   23
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "Current Value"
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
      Left            =   3480
      TabIndex        =   21
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Serial No"
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
      Left            =   4920
      TabIndex        =   19
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Purchase Amount"
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
      Left            =   1800
      TabIndex        =   17
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Purchase Date"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Account Name"
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
      Left            =   3480
      TabIndex        =   9
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Accno"
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
      Left            =   1920
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Depreciation Rate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1560
   End
   Begin VB.Label Label2 
      Caption         =   "Asset Name"
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
      Left            =   1800
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Asset Code"
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
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblaccname 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3480
      TabIndex        =   4
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label lblaccno 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1920
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmassetregistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
action = "new"
ClearControlsIn Me
lblaccno = ""
End Sub

Private Sub cmdedit_Click()
action = "edit"
End Sub

Private Sub cmdSearchGL_Click()
 frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            Dim Account As Account_Details
            Account = Get_Account_Details(SearchValue, "MAZIWA", ErrorMessage)
            If Account.AccountNo <> "" Then
                lblaccno = Account.AccountNo
                lblaccname = Account.AccountName
                
            Else
                'txtCHQNo = ""
                lblaccname = ""
            End If
        Else
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
    End If
End Sub

Private Sub cmdsvae_Click()
On Error GoTo kiki
If txtassetcode = "" Then
MsgBox "Put asset code", vbInformation
Exit Sub
End If
If txtassetname = "" Then
MsgBox "Put asset Name", vbInformation
Exit Sub
End If
If txtdeprate = "" Then
MsgBox "put depreciation rate for the asset", vbInformation
Exit Sub
End If
If txtpurchaseamt = "" Then
MsgBox "put the purchase price for the asset", vbInformation
txtpurchaseamt.SetFocus
Exit Sub
End If
If txtCURRENTVALUE = "" Then
MsgBox "put the purchase price for the asset", vbInformation
txtCURRENTVALUE.SetFocus
Exit Sub
End If

If lblaccno = "" Then
MsgBox "Put the account for the asset", vbInformation
Exit Sub
End If

If lblbankacc = "" Then
MsgBox "Put the Depreciation account for the asset", vbInformation
Exit Sub
End If

If action = "new" Then
   sql = "Select Assetcode from assets_register where assetcode='" & txtassetcode & "'"
   Set Rst1 = oSaccoMaster.GetRecordset(sql)
   If Not Rst1.EOF Then
   MsgBox "The Asset Already exist unless you want to edit ", vbInformation, Me.Caption
   Exit Sub
   End If
    sql = ""
    sql = "set dateformat dmy  INSERT INTO assets_register (assetcode, assetname, deprate, accno,ContraAccNo,pdate,PurchasePrice,CurrentValue,SerialNo)"
    sql = sql & " VALUES     ('" & txtassetcode & "', '" & txtassetname & "', " & txtdeprate & ", '" & lblaccno & "','" & lblbankacc & "','" & DTPicker1 & "'," & txtpurchaseamt & "," & txtCURRENTVALUE & ",'" & txtSERIALNO & "')"
oSaccoMaster.GetRecordset (sql)

MsgBox "Asset saved successfully", vbInformation
ElseIf action = "edit" Then
    sql = ""
    sql = "set dateformat dmy  UPDATE assets_register  SET assetname='" & txtassetname & "', deprate=" & txtdeprate & ", accno='" & lblaccno & "',ContraAccNo='" & lblbankacc & "',pdate='" & DTPicker1 & "',"
    sql = sql & "   PurchasePrice= " & txtpurchaseamt & ",CurrentValue=" & txtCURRENTVALUE & ",SerialNo='" & txtSERIALNO & "'   where assetcode='" & txtassetcode & "' "
oSaccoMaster.GetRecordset (sql)

MsgBox "Asset Updated successfully", vbInformation
ElseIf action = "" Then
MsgBox "You did not choose any task to perform", vbInformation
End If
ClearControlsIn Me
lblaccno = ""
frmassetsinquiry.Form_Load

Exit Sub
kiki:
MsgBox err.description, vbInformation
End Sub

Private Sub Command1_Click()
 frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            Dim Account As Account_Details
            Account = Get_Account_Details(SearchValue, "MAZIWA", ErrorMessage)
            If Account.AccountNo <> "" Then
                lblbankacc = Account.AccountNo
                lblbankaccname = Account.AccountName
                
            Else
                'txtCHQNo = ""
                lblbankaccname = ""
            End If
        Else
            If ErrorMessage <> "" Then
                MsgBox ErrorMessage, vbInformation, Me.Caption
                ErrorMessage = ""
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
 lblbankacc = "000-901"
End Sub

Private Sub lblaccno_Change()
Dim Glaccount As Account_Details
Glaccount = Get_Account_Details(lblaccno, "MAZIWA", ErrorMessage)
If Glaccount.AccountNo <> "" Then
lblaccname = Glaccount.AccountName
Else
lblaccno = ""
lblaccname = ""
End If

End Sub

Private Sub lblaccno_Click()
lblaccno_Change
End Sub

Private Sub lblbankacc_Change()
Dim Glaccount As Account_Details
Glaccount = Get_Account_Details(lblbankacc, "MAZIWA", ErrorMessage)
If Glaccount.AccountNo <> "" Then
lblbankaccname = Glaccount.AccountName
Else
lblbankacc = ""
lblbankaccname = ""
End If
End Sub

Private Sub lblbankacc_Click()
lblbankacc_Change
End Sub
