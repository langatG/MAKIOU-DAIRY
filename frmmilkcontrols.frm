VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmmilkcontrols 
   BackColor       =   &H00FF8080&
   Caption         =   "Stations Milk Control"
   ClientHeight    =   6090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15390
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   15390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdloadposted 
      Caption         =   "Load Milk Control"
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
      Left            =   5160
      TabIndex        =   9
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load "
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
      TabIndex        =   4
      Top             =   5520
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
      Left            =   8400
      TabIndex        =   3
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton cmdprocess 
      Caption         =   "Post Milk Control"
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
      Left            =   2760
      TabIndex        =   2
      Top             =   5520
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
      TabIndex        =   1
      Top             =   240
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
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0"
      Top             =   240
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker dtpendate 
      Height          =   315
      Left            =   9000
      TabIndex        =   5
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
      Format          =   116260867
      CurrentDate     =   39601
   End
   Begin MSComctlLib.ListView lvwstations 
      Height          =   4545
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   8017
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Bcode"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Stations"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "INTAKE"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "B/F"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "DISPATCH"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "C/F"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ACTUAL"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "REJECTED"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "SPILLAGE"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "VAR"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TransDate"
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
      TabIndex        =   8
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "Total Intake"
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
      TabIndex        =   7
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmmilkcontrols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objLabelEdit As LabelEdit
Dim objLabelEdit2 As LabelEdit
Dim objLabelEdit3 As LabelEdit
Dim actual As Double, Intake As Double, Bf As Double, Dispatch As Double, Cf As Double, Rej As Double, Spil As Double, Var As Double
Dim Actuals As Double

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
    Intake = 0
    actual = 0
    Bf = 0
    Dispatch = 0
    Cf = 0
    Actuals = 0
    Rej = 0
    Spil = 0
    Var = 0
    For I = 1 To lvwstations.ListItems.Count
    Set li = lvwstations.ListItems(I)
    If li.Checked = True Then
     If I <> 8 Then
     actual = CDbl(li.SubItems(4)) + CDbl(li.SubItems(5)) - (CDbl(li.SubItems(2)) + CDbl(li.SubItems(3)))
     li.SubItems(6) = Format(CDbl(actual), Cfmt)
     li.SubItems(9) = Format(CDbl(actual) + CDbl(li.SubItems(7)) + CDbl(li.SubItems(8)), Cfmt)
     Intake = Format(Intake + CDbl(li.SubItems(2)), Cfmt)
     Bf = Format(Bf + CDbl(li.SubItems(3)), Cfmt)
     Dispatch = Format(Dispatch + CDbl(li.SubItems(4)), Cfmt)
     Cf = Format(Cf + CDbl(li.SubItems(5)), Cfmt)
     Rej = Format(Rej + CDbl(li.SubItems(7)), Cfmt)
     Spil = Format(Spil + CDbl(li.SubItems(8)), Cfmt)
     Var = Format(Var + CDbl(li.SubItems(9)), Cfmt)
     Actuals = Format(Actuals + CDbl(actual), Cfmt)
     End If
     If I = 8 Then
     li.SubItems(2) = Intake
     li.SubItems(3) = Bf
     li.SubItems(4) = Dispatch
     li.SubItems(5) = Cf
     li.SubItems(6) = Actuals
     li.SubItems(7) = Rej
     li.SubItems(8) = Spil
     li.SubItems(9) = Var
     End If
    End If
    Next I
    txttotal = Format(Intake, Cfmt)
End Sub

Private Sub cmdexport_Click()
  On Error GoTo SsyError
    Dim MyFso As New FileSystemObject, strData As String, MFile As TextStream, _
    FileName As String, I As Long, li As ListItem
    If lvwstations.ListItems.Count > 0 Then
        With CommonDialog1
            .Filter = "Comma Seperated Values|*.csv"
            .FileName = "Stations Milk Control " & Format(dtpendate, "dd-mm-yyyy")
            .ShowSave
            If .FileName <> "" Then
                FileName = .FileName
            End If
            .FileName = ""
        End With
        Set MFile = MyFso.OpenTextFile(FileName, ForWriting, True)
        strData = "Stations Milk Control :" & Format(dtpendate, "dd-mm-yyyy")
        MFile.WriteLine strData
        strData = "STATIONS,INTAKE,B/F,DISPATCH,C/F,ACTUALS,REJECTED,SPILLAGE,VAR,"
        MFile.WriteLine strData
        strData = ""
        For I = 1 To lvwstations.ListItems.Count
            Set li = lvwstations.ListItems(I)
            strData = CStr(li.SubItems(1)) & "," & CDbl(li.SubItems(2)) & "," & CDbl(li.SubItems(3)) _
            & "," & CDbl(li.SubItems(4)) & "," & CDbl(li.SubItems(5)) & "," & CDbl(li.SubItems(6)) _
            & " , " & CDbl(li.SubItems(7)) & "," & CDbl(li.SubItems(8)) & "," & CDbl(li.SubItems(9)) & ""
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
 Set Rst = oSaccoMaster.GetRecordset("Select Bcode,Bname  from D_branch where bcode<>'All' order by Bcode")
 With Rst
    While Not .EOF
         Set li = lvwstations.ListItems.Add(, , IIf(IsNull(!bcode), "", !bcode))
             li.SubItems(1) = IIf(IsNull(!Bname), "", !Bname)
             li.SubItems(2) = 0
             li.SubItems(3) = 0
             li.SubItems(4) = 0
             li.SubItems(5) = 0
             li.SubItems(6) = 0
             li.SubItems(7) = 0
             li.SubItems(8) = 0
             li.SubItems(9) = 0
       .MoveNext
    Wend
 End With
 Set li = lvwstations.ListItems.Add(, , "TOTALS")
             li.SubItems(1) = 0
             li.SubItems(2) = 0
             li.SubItems(3) = 0
             li.SubItems(4) = 0
             li.SubItems(5) = 0
             li.SubItems(6) = 0
             li.SubItems(7) = 0
             li.SubItems(8) = 0
             li.SubItems(9) = 0

End Sub

Private Sub cmdloadposted_Click()
lvwstations.ListItems.Clear
 Set Rst = oSaccoMaster.GetRecordset("set dateformat dmy Select  M.*,B.BName from  milkcontrol2 M Inner Join D_Branch B on B.Bcode=M.Bcode where Transdate='" & dtpendate & "'  order by M.Bcode")
 With Rst
    While Not .EOF
         Set li = lvwstations.ListItems.Add(, , !bcode)
             li.SubItems(1) = IIf(IsNull(!Bname), "", !Bname)
             li.SubItems(2) = Format(!Intake, Cfmt)
            li.SubItems(3) = Format(!Bf, Cfmt)
            li.SubItems(4) = Format(!SQuantity, Cfmt)
            li.SubItems(5) = Format(!Cfa, Cfmt)
            li.SubItems(6) = 0
            li.SubItems(7) = Format(!Reject, Cfmt)
            li.SubItems(8) = Format(!Spillage, Cfmt)
            li.SubItems(9) = 0
       .MoveNext
    Wend
 End With
 Set li = lvwstations.ListItems.Add(, , "TOTALS")
             li.SubItems(1) = 0
             li.SubItems(2) = 0
             li.SubItems(3) = 0
             li.SubItems(4) = 0
             li.SubItems(5) = 0
             li.SubItems(6) = 0
             li.SubItems(7) = 0
             li.SubItems(8) = 0
             li.SubItems(9) = 0
End Sub

Private Sub cmdprocess_Click()
 Dim postingdate As Date, Qty As Double, Value As Double, Pcode As String
   Dim post As New ADODB.Connection, NAMES As String, pprice As Double
   
    If lvwstations.ListItems.Count > 0 Then
        ProgressBar1.Max = lvwstations.ListItems.Count
    Else
        MsgBox "Please Load Cash Income to be Received", vbInformation, Me.Caption
        Exit Sub
    End If
    If txttotal = "" Then
        MsgBox "Please Select First Stations Milk Controls  post", vbInformation
        chkselectall.SetFocus
      Exit Sub
    End If
    
    If MsgBox("Do You want to Post the selected  Stations Milk Controls", vbQuestion + vbYesNo, _
    Me.Caption) = vbNo Then
        Exit Sub
    End If
 
    I = 0
  
  With post
    .Open "MAZIWA", "atm", "atm"
      .BeginTrans
         On Error GoTo TransError
         ProgressBar1.Visible = True
         dtpendate = Format(dtpendate, "dd/mm/yyyy") 'DateSerial(Year(dtpendate), month(dtpendate) + 1, 1 - 1)
         
         NewTransaction CDbl(txttotal), dtpendate, "Stations Milk Controls Posting"
      For I = 1 To lvwstations.ListItems.Count
               Set li = lvwstations.ListItems(I)
               ProgressBar1.Value = I
               DoEvents
            If li.Checked = True Then
                If I <> 8 Then
                   bcode = Trim(CStr(li))
                   postingdate = Format(dtpendate, "dd/mm/yyyy")
                   NAMES = li.SubItems(1)
                    Intake = Format(CDbl(li.SubItems(2)), Cfmt)
                    Bf = Format(CDbl(li.SubItems(3)), Cfmt)
                    Dispatch = Format(CDbl(li.SubItems(4)), Cfmt)
                    Cf = Format(CDbl(li.SubItems(5)), Cfmt)
                    Rej = Format(CDbl(li.SubItems(7)), Cfmt)
                    Spil = Format(CDbl(li.SubItems(8)), Cfmt)
                    Var = Format(CDbl(li.SubItems(9)), Cfmt)
                    Actuals = Format(CDbl(li.SubItems(6)), Cfmt)
                   
                    If Intake > 0 Then
                        Set Rst = oSaccoMaster.GetRecordset(" set dateformat dmy select * from milkcontrol2 where transdate='" & postingdate & "'and bcode='" & bcode & "' ")
                         If Rst.EOF Then
                           oSaccoMaster.ExecuteThis " set dateformat dmy insert into milkcontrol2 (Intake,SQuantity,Reject,transdate,auditid,cfa,Spillage,FromStation,Tostation,Bf,bcode)values ('" & Intake & "','" & Dispatch & "','" & Rej & "','" & postingdate & "','" & User & "','" & Cf & "','" & Spil & "',0,0,'" & Bf & "','" & bcode & "') "
                         Else
                           oSaccoMaster.ExecuteThis " set dateformat dmy update  milkcontrol2  set intake='" & Intake & "',SQuantity='" & Dispatch & "',Reject= '" & Rej & "',spillage= '" & Spil & "',Bcode= '" & bcode & "',auditid='" & User & "',cfa='" & Cf & "',FromStation=0 ,Tostation=0,bf='" & Bf & "'   where transdate='" & postingdate & "'  AND bcode='" & bcode & "' "
                           
                         End If
                    End If
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
 'dtpendate = DateSerial(Year(dtpendate), month(dtpendate) + 1, 1 - 1)
 InitSubClass
    Set objLabelEdit = New LabelEdit
    objLabelEdit.Init Me, lvwstations
   
End Sub

Private Sub lvwstations_ItemClick(ByVal Item As MSComctlLib.ListItem)
  calcTotal
End Sub
