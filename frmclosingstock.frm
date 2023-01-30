VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmclosingstock 
   BackColor       =   &H00FF8080&
   Caption         =   "Monthly Closing Stock"
   ClientHeight    =   8685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdloadposted 
      Caption         =   "Load Closed Stock"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   8280
      Width           =   2055
   End
   Begin VB.TextBox txttotal 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   3840
      TabIndex        =   7
      Text            =   "0"
      Top             =   120
      Width           =   3015
   End
   Begin VB.CheckBox chkselectall 
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdprocess 
      Caption         =   "Post Closing Stock"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   8280
      Width           =   2055
   End
   Begin VB.CommandButton cmdexport 
      Caption         =   "Export To Excel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   4
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load Stock Bal"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   8280
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker dtpendate 
      Height          =   315
      Left            =   9000
      TabIndex        =   0
      Top             =   120
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   " dd-MM-yyyy"
      Format          =   106889219
      CurrentDate     =   39601
   End
   Begin MSComctlLib.ListView lvwproducts 
      Height          =   7425
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   13097
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pcode"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PName"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Pprice"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Value"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Sprice"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Station"
         Object.Width           =   3175
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   7920
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Total Stock(PP)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "End Month"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      TabIndex        =   1
      Top             =   120
      Width           =   1155
   End
End
Attribute VB_Name = "frmclosingstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objLabelEdit As LabelEdit
Dim objLabelEdit2 As LabelEdit
Dim objLabelEdit3 As LabelEdit

Private Sub chkselectall_Click()
If chkselectall.Value = vbChecked Then
For I = 1 To lvwproducts.ListItems.Count
    Set li = lvwproducts.ListItems(I)
    If li.Checked = False Then
     li.Checked = True
    End If
    Next I
Else
For I = 1 To lvwproducts.ListItems.Count
    Set li = lvwproducts.ListItems(I)
    If li.Checked = True Then
     li.Checked = False
    End If
    Next I
End If
calcTotal

End Sub
Sub calcTotal()
  Dim totalvalue As Double
    totalvalue = 0
    For I = 1 To lvwproducts.ListItems.Count
    Set li = lvwproducts.ListItems(I)
    If li.Checked = True Then
     totalvalue = totalvalue + (CDbl(li.SubItems(2)) * CDbl(li.SubItems(3)))
    End If
    Next I
    txttotal = Format(totalvalue, Cfmt)
End Sub

Private Sub cmdexport_Click()
  On Error GoTo SsyError
    Dim MyFso As New FileSystemObject, strData As String, MFile As TextStream, _
    FileName As String, I As Long, li As ListItem
    If lvwproducts.ListItems.Count > 0 Then
        With CommonDialog1
            .Filter = "Comma Seperated Values|*.csv"
            .FileName = "Closing Stock " & Format(dtpendate, "dd-mm-yyyy")
            .ShowSave
            If .FileName <> "" Then
                FileName = .FileName
            End If
            .FileName = ""
        End With
        Set MFile = MyFso.OpenTextFile(FileName, ForWriting, True)
        strData = "Closing Stock Income :" & Format(dtpendate, "dd-mm-yyyy")
        MFile.WriteLine strData
        strData = "PCode,PName,Qty,PPrice,Value,Sprice,Station,"
        MFile.WriteLine strData
        strData = ""
        For I = 1 To lvwproducts.ListItems.Count
            Set li = lvwproducts.ListItems(I)
            strData = li & "," & li.SubItems(1) & "," & CDbl(li.SubItems(2)) & "," & CStr(li.SubItems(3)) _
            & "," & CStr(li.SubItems(4)) & "," & CStr(li.SubItems(5)) & "," & CStr(li.SubItems(6)) & ""
            MFile.WriteLine strData
            strData = ""
        Next I
    Else
        MsgBox "There are no records to be exported", vbInformation, Me.Caption
    End If
    MsgBox "Items Successfully Imported Into CSV file", vbOKOnly
    Exit Sub
SsyError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cmdload_Click()
 lvwproducts.ListItems.Clear
 Set Rst = oSaccoMaster.GetRecordset("Select  p_code, p_name,Qout, pprice,Sprice from ag_Products order by p_name")
 With Rst
    While Not .EOF
         Set li = lvwproducts.ListItems.Add(, , !p_code)
             li.SubItems(1) = IIf(IsNull(!p_name), "", !p_name)
             li.SubItems(2) = Format(!Qout, Cfmt)
             li.SubItems(3) = IIf(IsNull(!pprice), 0, !pprice)
             li.SubItems(4) = Format(li.SubItems(3) * li.SubItems(2), Cfmt)
             li.SubItems(5) = IIf(IsNull(!sprice), 0, !sprice)
             li.SubItems(6) = "LELCHEGO"
       .MoveNext
    Wend
 End With
End Sub

Private Sub cmdloadposted_Click()
lvwproducts.ListItems.Clear
 Set Rst = oSaccoMaster.GetRecordset("Select   Pcode,PName,Qty,Pprice,ClosingStock from  Ag_ClosingStock where Month(Transdate)='" & month(dtpendate) & "' and Year(Transdate)='" & year(dtpendate) & "' order by PName")
 With Rst
    While Not .EOF
         Set li = lvwproducts.ListItems.Add(, , !Pcode)
             li.SubItems(1) = IIf(IsNull(!PName), "", !PName)
             li.SubItems(2) = Format(!Qty, Cfmt)
             li.SubItems(3) = IIf(IsNull(!pprice), 0, !pprice)
             li.SubItems(4) = Format(!ClosingStock, Cfmt)
             li.SubItems(5) = 0
             li.SubItems(6) = "LELCHEGO"
       .MoveNext
    Wend
 End With
End Sub

Private Sub cmdprocess_Click()
 Dim postingdate As Date, Qty As Double, Value As Double, Pcode As String
   Dim post As New ADODB.Connection, NAMES As String, pprice As Double
   
    If lvwproducts.ListItems.Count > 0 Then
        ProgressBar1.Max = lvwproducts.ListItems.Count
    Else
        MsgBox "Please Load Cash Income to be Received", vbInformation, Me.Caption
        Exit Sub
    End If
    If txttotal = "" Then
        MsgBox "Please Select First Closing Stock to post", vbInformation
        txttotal.SetFocus
      Exit Sub
    End If
    
    If MsgBox("Do You want to Post the selected  Closing Stock", vbQuestion + vbYesNo, _
    Me.Caption) = vbNo Then
        Exit Sub
    End If
 
    I = 0
  
  With post
    .Open "MAZIWA"
      .BeginTrans
         On Error GoTo TransError
         ProgressBar1.Visible = True
         dtpendate = DateSerial(year(dtpendate), month(dtpendate) + 1, 1 - 1)
         
         NewTransaction CDbl(txttotal), dtpendate, "Endmonth Closing Stock Posting"
      For I = 1 To lvwproducts.ListItems.Count
               Set li = lvwproducts.ListItems(I)
               ProgressBar1.Value = I
               DoEvents
            If li.Checked = True Then
               Pcode = CStr(li)
               postingdate = dtpendate
               NAMES = li.SubItems(1)
               Qty = li.SubItems(2)
               pprice = li.SubItems(3)
               Value = Qty * pprice
               
                If Qty > 0 Then
                      Set rs = oSaccoMaster.GetRecordset("select Pcode from Ag_ClosingStock Where Pcode='" & Pcode & "' and Month(Transdate)='" & month(dtpendate) & "' and Year(Transdate)='" & year(dtpendate) & "'")
                      If Not rs.EOF Then
                      oSaccoMaster.ExecuteThis ("update Ag_ClosingStock set Qty=" & Qty & ",pprice=" & pprice & ",ClosingStock=" & Value & ",AuditId='" & user & "' Where Pcode='" & Pcode & "' and Month(Transdate)='" & month(dtpendate) & "' and Year(Transdate)='" & year(dtpendate) & "'  ")
                      Else
                      sql = "set dateformat dmy INSERT INTO Ag_ClosingStock(Pcode,PName,TransDate,Qty,pprice,ClosingStock,AuditId)" _
                         & "VALUES ('" & Pcode & "','" & NAMES & "','" & postingdate & "'," & Qty & "," & pprice & "," & Value & ",'" & user & "') "
                        oSaccoMaster.ExecuteThis (sql)
                    End If
                End If
               
            End If
           li.Checked = False
    Next I
   .CommitTrans
  MsgBox "Closing Stock Posted Successfully", vbInformation, Me.Caption
  lvwproducts.ListItems.Clear
  Exit Sub
TransError:
    If err.number = 35600 Then
        Resume Next
    End If
    If ErrorMessage = "" And err.description = "" Then
       .RollbackTrans
    Else
        MsgBox IIf(ErrorMessage = "", err.description, ErrorMessage) & vbNewLine & "Action Therefore Aborted. ", vbCritical, Me.Caption
        .RollbackTrans
    End If
  End With

End Sub



Private Sub Form_Load()
 dtpendate = Format(Get_Server_Date, "dd/mm/yyyy")
 dtpendate = DateSerial(year(dtpendate), month(dtpendate) + 1, 1 - 1)
 InitSubClass
    Set objLabelEdit = New LabelEdit
    objLabelEdit.Init Me, lvwproducts
   
End Sub
