VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmstationscstock 
   BackColor       =   &H00FF8080&
   Caption         =   "Stations Closing Stock"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   9255
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtsalesvalue 
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
      Left            =   6240
      TabIndex        =   11
      Text            =   "0"
      Top             =   5280
      Width           =   2535
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load Stations"
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
      Left            =   120
      TabIndex        =   5
      Top             =   5760
      Width           =   1695
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
      Left            =   6840
      TabIndex        =   4
      Top             =   5760
      Width           =   1695
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
      Left            =   1920
      TabIndex        =   3
      Top             =   5760
      Width           =   2055
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
      TabIndex        =   2
      Top             =   120
      Width           =   1335
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
      Left            =   1800
      TabIndex        =   1
      Text            =   "0"
      Top             =   5280
      Width           =   2535
   End
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
      Left            =   4080
      TabIndex        =   0
      Top             =   5760
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker dtpendate 
      Height          =   315
      Left            =   7680
      TabIndex        =   6
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
      Format          =   148373507
      CurrentDate     =   39601
   End
   Begin MSComctlLib.ListView lvwstations 
      Height          =   4425
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   7805
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Bcode"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Station"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "StockValue(PP)"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "StockSalesValue"
         Object.Width           =   4410
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   4920
      Visible         =   0   'False
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Total Stock(SP)"
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
      Left            =   4680
      TabIndex        =   12
      Top             =   5280
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
      Left            =   6360
      TabIndex        =   10
      Top             =   120
      Width           =   1155
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
      Left            =   240
      TabIndex        =   9
      Top             =   5280
      Width           =   1575
   End
End
Attribute VB_Name = "frmstationscstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objLabelEdit As LabelEdit
Dim objLabelEdit2 As LabelEdit
Dim objLabelEdit3 As LabelEdit

Private Sub chkselectall_Click()
If chkselectall.Value = vbChecked Then
For I = 1 To lvwstations.ListItems.Count
    Set li = lvwstations.ListItems(I)
    If li.Checked = False Then
     li.Checked = True
    End If
    Next I
Else
For I = 1 To lvwstations.ListItems.Count
    Set li = lvwstations.ListItems(I)
    If li.Checked = True Then
     li.Checked = False
    End If
    Next I
End If
calcTotal

End Sub
Sub calcTotal()
  Dim totalvalue As Double, Totalsales As Double
    totalvalue = 0
    For I = 1 To lvwstations.ListItems.Count
    Set li = lvwstations.ListItems(I)
    If li.Checked = True Then
     totalvalue = totalvalue + CDbl(li.SubItems(2))
     Totalsales = Totalsales + CDbl(li.SubItems(3))
    End If
    Next I
    txttotal = Format(totalvalue, Cfmt)
    txtsalesvalue = Format(Totalsales, Cfmt)
End Sub


Private Sub cmdexport_Click()
  On Error GoTo SsyError
    Dim MyFso As New FileSystemObject, strData As String, MFile As TextStream, _
    FileName As String, I As Long, li As ListItem
    If lvwstations.ListItems.Count > 0 Then
        With CommonDialog1
            .Filter = "Comma Seperated Values|*.csv"
            .FileName = "Stations Closing Stock " & Format(dtpendate, "dd-mm-yyyy")
            .ShowSave
            If .FileName <> "" Then
                FileName = .FileName
            End If
            .FileName = ""
        End With
        Set MFile = MyFso.OpenTextFile(FileName, ForWriting, True)
        strData = "Stations Closing Stock Income :" & Format(dtpendate, "dd-mm-yyyy")
        MFile.WriteLine strData
        strData = "BCode,Station,Stock Value,Sales Value,"
        MFile.WriteLine strData
        strData = ""
        For I = 1 To lvwstations.ListItems.Count
            Set li = lvwstations.ListItems(I)
            strData = li & "," & li.SubItems(1) & "," & CDbl(li.SubItems(2)) & "," & CStr(li.SubItems(3)) _
            & ""
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
 lvwstations.ListItems.Clear
 Set Rst = oSaccoMaster.GetRecordset("Select  * from  Ag_Station order by Station")
 With Rst
    While Not .EOF
         Set li = lvwstations.ListItems.Add(, , !bcode)
             li.SubItems(1) = IIf(IsNull(!Station), "", !Station)
             li.SubItems(2) = Format(0, Cfmt)
             li.SubItems(3) = Format(0, Cfmt)
       .MoveNext
    Wend
 End With
End Sub

Private Sub cmdloadposted_Click()
  lvwstations.ListItems.Clear
 Set Rst = oSaccoMaster.GetRecordset("Select S.Bcode,S.Station,C.ClosingStockV,C.ClosingStocks from  Ag_StationCStock C inner Join Ag_Station S  on S.Bcode=C.Bcode where Month(C.Transdate)='" & month(dtpendate) & "' and Year(C.Transdate)='" & Year(dtpendate) & "' order by C.Bcode")
 With Rst
    While Not .EOF
         Set li = lvwstations.ListItems.Add(, , !bcode)
             li.SubItems(1) = IIf(IsNull(!Station), "", !Station)
             li.SubItems(2) = Format(!ClosingStockV, Cfmt)
             li.SubItems(3) = Format(!ClosingStocks, Cfmt)
       .MoveNext
    Wend
 End With
End Sub

Private Sub cmdprocess_Click()
  Dim postingdate As Date, Qty As Double, StockValue As Double, Pcode As String
   Dim post As New ADODB.Connection, NAMES As String, SalesValue As Double
   
    If lvwstations.ListItems.Count > 0 Then
        ProgressBar1.Max = lvwstations.ListItems.Count
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
         dtpendate = DateSerial(Year(dtpendate), month(dtpendate) + 1, 1 - 1)
         
         NewTransaction CDbl(txttotal), dtpendate, "Staions Closing Stock Posting"
      For I = 1 To lvwstations.ListItems.Count
               Set li = lvwstations.ListItems(I)
               ProgressBar1.Value = I
               DoEvents
            If li.Checked = True Then
               Pcode = CStr(li)
               postingdate = dtpendate
               NAMES = li.SubItems(1)
               StockValue = li.SubItems(2)
               SalesValue = li.SubItems(3)
               
                    Set rs = oSaccoMaster.GetRecordset("select Bcode from Ag_StationCStock Where Bcode='" & Pcode & "' and Month(Transdate)='" & month(dtpendate) & "' and Year(Transdate)='" & Year(dtpendate) & "'")
                    If Not rs.EOF Then
                    oSaccoMaster.ExecuteThis ("update Ag_StationCStock set ClosingStockV=" & StockValue & ",ClosingStocks=" & SalesValue & ",AuditId='" & user & "' Where Bcode='" & Pcode & "' and Month(Transdate)='" & month(dtpendate) & "' and Year(Transdate)='" & Year(dtpendate) & "'  ")
                    Else
                    sql = "set dateformat dmy INSERT INTO Ag_StationCStock(Bcode,TransDate,ClosingStockV,ClosingStocks,AuditId)" _
                       & "VALUES ('" & Pcode & "','" & postingdate & "'," & StockValue & "," & SalesValue & ",'" & user & "') "
                      oSaccoMaster.ExecuteThis (sql)
                    End If
               
            End If
           li.Checked = False
    Next I
   .CommitTrans
  MsgBox "Closing Stock Posted Successfully", vbInformation, Me.Caption
  lvwstations.ListItems.Clear
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
 dtpendate = DateSerial(Year(dtpendate), month(dtpendate) + 1, 1 - 1)
 InitSubClass
    Set objLabelEdit = New LabelEdit
    objLabelEdit.Init Me, lvwstations
End Sub
