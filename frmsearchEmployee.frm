VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsearchEmployee 
   Caption         =   "Browse for Employee"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmsearchEmployee.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6855
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lvwMembers 
      Height          =   2775
      Left            =   0
      TabIndex        =   10
      Top             =   1560
      Width           =   6735
      _ExtentX        =   11880
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "empNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "surname"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "othernames"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtOrganisation 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1545
      TabIndex        =   9
      Top             =   360
      Width           =   5190
   End
   Begin VB.ComboBox cboOrganisation 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmsearchEmployee.frx":0442
      Left            =   60
      List            =   "frmsearchEmployee.frx":0444
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   360
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5445
      TabIndex        =   6
      Top             =   4380
      Width           =   1275
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   4440
      Width           =   1275
   End
   Begin VB.TextBox txtValue 
      Height          =   315
      Left            =   1890
      TabIndex        =   3
      Top             =   1035
      Width           =   2625
   End
   Begin VB.ComboBox cboSearchField 
      Height          =   315
      ItemData        =   "frmsearchEmployee.frx":0446
      Left            =   90
      List            =   "frmsearchEmployee.frx":0456
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1035
      Width           =   1695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Search Value"
      Height          =   195
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Organisation Code"
      Height          =   195
      Left            =   90
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Search Value"
      Height          =   195
      Left            =   1905
      TabIndex        =   2
      Top             =   795
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Search Field"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   795
      Width           =   870
   End
End
Attribute VB_Name = "frmsearchEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboOrganisation_Change()
    On Error GoTo sysError
    Dim rsOrg As New Recordset
    Set rsOrg = oSaccoMaster.Get_Payroll_Recordset("Select * From D_MCompany WHERE mcode='" & cboOrganisation & "'")
    With rsOrg
        If .State = adStateOpen Then
            If Not .EOF Then
                txtOrganisation = !name
            Else
                txtOrganisation = "All"
            End If
        Else
            txtOrganisation = "All"
        End If
    End With
    txtValue_Change
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub cboOrganisation_Click()
    cboOrganisation_Change
End Sub

Private Sub CboSearchField_Change()
    txtValue_Change
End Sub

Private Sub cboSearchField_Click()
    txtValue_Change
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If lvwMembers.ListItems.Count > 0 Then
        'Continue = True
        SearchValue = lvwMembers.SelectedItem
    Else
        'Continue = False
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo sysError
    Dim rsOrg As New Recordset
    cboOrganisation.Clear
    cboOrganisation.AddItem "All"
    Set rsOrg = oSaccoMaster.Get_Payroll_Recordset("Select * From d_MCOMPANY mcode ORDER BY MCODE")
    With rsOrg
        If .State = adStateOpen Then
            While Not .EOF
                cboOrganisation.AddItem !MCODE
                .MoveNext
            Wend
        End If
    End With
    'CboSearchField.Text = CboSearchField.List(2)
    
    'Continue = True
    cboOrganisation.ListIndex = 0
    cboSearchField.ListIndex = 2
    'Continue = False
    On Error Resume Next
    'txtValue.SetFocus
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub

Private Sub lvwMembers_DblClick()
    If lvwMembers.ListItems.Count > 0 Then
        SearchValue = lvwMembers.SelectedItem
        'Continue = True
    Else
        'Continue = False
    End If
    Unload Me
End Sub

Private Sub txtValue_Change()
    On Error GoTo sysError
    Dim RsMembers As New Recordset
   
    If Trim$(txtValue) <> "" Then
        Select Case cboOrganisation.Text
            Case "All"
            Select Case cboSearchField.ListIndex
                Case 0 'empno
                'Set rsMembers = oSaccoMaster.Get_Payroll_Recordset("Select * From EMPLOYEES where " _
                & "empno Like '" & txtValue & "%' order by empno")
                Set RsMembers = oSaccoMaster.Get_Payroll_Recordset("Select * From employees order by mcode,empcode")
                Case 1 'StaffNo
                Set RsMembers = oSaccoMaster.Get_Payroll_Recordset("Select * From employees where " _
                & "empno Like '%" & txtValue _
                & "%' order by SurName")
                Case 2 'Names
                Set RsMembers = oSaccoMaster.Get_Payroll_Recordset("Select * From employees where " _
                & " SurName+Othernames  Like '%" & txtValue _
                & "%' order by mcode,empNO")
                Case 3 'IdNo
                Set RsMembers = oSaccoMaster.Get_Payroll_Recordset("Select * From employees where " _
                & " IdNo  Like '%" & txtValue _
                & "%' order by mcode,empNO")
            End Select
            Case Else
            Select Case cboSearchField.ListIndex
                Case 0 'empno
                Set RsMembers = oSaccoMaster.Get_Payroll_Recordset("Select * From employees where " _
                & "MCode='" & cboOrganisation & "' and empno Like '" & txtValue _
                & "%' order by mcode,empNO")
                Case 1 'StaffNo
                Set RsMembers = oSaccoMaster.Get_Payroll_Recordset("Select * From Employees where " _
                & "mcode='" & cboOrganisation & "' and StaffNo Like '%" & txtValue _
                & "%' order by mcode,empNO")
                Case 2 'Names
                Set RsMembers = oSaccoMaster.Get_Payroll_Recordset("Select * From Employees where " _
                & " mcode='" & cboOrganisation & "' and SurName+Othernames  Like '%" & txtValue _
                & "%' order by mcode,empNO")
                Case 3 'IdNo
                Set RsMembers = oSaccoMaster.Get_Payroll_Recordset("Select * From EMPLOYEES where " _
                & " IDNo  Like '%" & txtValue _
                & "%' order by mcode,empNO")
            End Select
        End Select
    Else
        Select Case cboOrganisation.Text
            Case "All"
            Select Case cboSearchField.ListIndex
                Case 0
                Set RsMembers = oSaccoMaster.Get_Payroll_Recordset("Select * From EMPLOYEES Order By empno")
                Case 1
                Set RsMembers = oSaccoMaster.Get_Payroll_Recordset("Select * From EMPLOYEES order by mcode,empNO")
                Case 2
                Set RsMembers = oSaccoMaster.Get_Payroll_Recordset("Select * From EMPLOYEES order by mcode,empNO")
            End Select
            Case Else
            Select Case cboSearchField.ListIndex
                Case 0
                Set RsMembers = oSaccoMaster.Get_Payroll_Recordset("Select * From EMPLOYEES where mcode" _
                & "='" & cboOrganisation & "' Order By empno")
                Case 1
                Set RsMembers = oSaccoMaster.Get_Payroll_Recordset("Select * From EMPLOYEES where mcode" _
                & "='" & cboOrganisation & "' order by mcode,empcode")
                Case 2
                Set RsMembers = oSaccoMaster.Get_Payroll_Recordset("Select * From EMPLOYEES where mcode" _
                & "='" & cboOrganisation & "' order by mcode,empcode")
            End Select
        End Select
    End If
    lvwMembers.ListItems.Clear
    With RsMembers
        If .State = adStateOpen Then
        While Not .EOF
            Set li = lvwMembers.ListItems.Add(, , !EMPNO)
            li.SubItems(1) = IIf(IsNull(!surname), "", !surname)
            li.SubItems(2) = IIf(IsNull(!othernames), "", !othernames)
            .MoveNext
        Wend
        End If
    End With
 
    Exit Sub
sysError:
    MsgBox err.description, vbInformation, Me.Caption
End Sub
