VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsearchusergroups 
   Caption         =   "SEARCH GROUPS"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton cmdCancel 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         Picture         =   "frmsearchusergroups.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cancel"
         Top             =   4080
         Width           =   495
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "SELECT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   8
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   3840
         TabIndex        =   4
         Top             =   120
         Width           =   2895
         Begin VB.TextBox txtTo 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox txtFrom 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton cmdFind 
            Enabled         =   0   'False
            Height          =   375
            Left            =   2400
            Picture         =   "frmsearchusergroups.frx":0102
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.ComboBox cboCrieria 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         ItemData        =   "frmsearchusergroups.frx":0204
         Left            =   2280
         List            =   "frmsearchusergroups.frx":021D
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cboField 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         ItemData        =   "frmsearchusergroups.frx":0241
         Left            =   120
         List            =   "frmsearchusergroups.frx":0243
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmdRef 
         Height          =   495
         Left            =   3840
         Picture         =   "frmsearchusergroups.frx":0245
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Refresh"
         Top             =   4080
         Width           =   495
      End
      Begin MSComctlLib.ListView lstSearch 
         Height          =   2535
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4471
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   16711680
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label3 
         Caption         =   "Criteria"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblRecords 
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Records Found"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Search Field"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmsearchusergroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()
 Sel = ""
    frmSearchVendor.Visible = False
End Sub

Private Sub cmdFind_Click()
Dim Find As Long
Dim li As ListItem
Dim Field As String
Set cn = New connection
Set Rst = New Recordset
Field = cboField.Text
'Find = Button.Index
lstSearch.ListItems.Clear
  

If Not cboField.Text = "" Then
    If Not cboCrieria.Text = "" Then
        If Not cboCrieria.Text = "Between" And Not cboCrieria.Text = "Like" Then
        If cboField.Text = "GroupID" Then
        sql = "SELECT GroupName,Groupid FROM  UserGroupss where " & cboField.Text & "" & cboCrieria.Text & "" & txtFrom.Text & ""
 
        Else
            sql = "SELECT GroupName,Groupid FROM  UserGroups where " & cboField.Text & "" & cboCrieria.Text & "'" & txtFrom.Text & "'"
            End If
           oSaccoMaster.GetRecordSet (sql)
            
            With rs
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
 Set li = frmSearchVendor.lstSearch.ListItems.Add(, , !groupId)
                li.SubItems(1) = !GroupName & ""
                
                
                
                
                        .MoveNext
                    Loop
                End If
            End With
            
            Set rs = Nothing
            
        ElseIf cboCrieria.Text = "Like" Then
         'If cboField.Text = "R_No" Then
            sql = "SELECT Groupname,Groupid FROM  UserAccounts order by Groupid"
            'Else
            'end if
           oSaccoMaster.GetRecordSet (sql)
            
            With rs
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
                     If cboField.Text = "GroupID" Then
                        .Find "" & cboField.Text & " " & cboCrieria.Text & " " & txtFrom.Text & "%", , adSearchForward
                        Else
                         .Find "" & cboField.Text & " " & cboCrieria.Text & " '" & txtFrom.Text & "%'", , adSearchForward

                        End If
                        If Not .EOF Then
 Set li = frmsearchusergroups.lstSearch.ListItems.Add(, , !groupId)
                li.SubItems(1) = !GroupName & ""
              
                
                            
                            .MoveNext
                        End If
                        
                    Loop
                End If
            End With
            
            Set rs = Nothing
            
            
        Else
            If cboField.Text = "GroupName" Then
                sql = "SELECT username,userloginid FROM  UserAccounts where " & cboField.Text & " " & cboCrieria & " " & txtFrom.Text & " And " & txtTo.Text & ""
               oSaccoMaster.GetRecordSet (sql)
            Else
                sql = "SELECT username,userloginid FROM  UserAccounts where " & cboField.Text & " " & cboCrieria & " '" & txtFrom.Text & "' And '" & txtTo.Text & "'"
               oSaccoMaster.GetRecordSet (sql)
            End If
            
            With rs
                If .RecordCount > 0 Then
                    .MoveFirst
                    Do While Not .EOF
 Set li = frmsearchusergroups.lstSearch.ListItems.Add(, , !groupId)
                li.SubItems(1) = !userName & ""
                
                
                        .MoveNext
                    Loop
                End If
            End With
            
            Set rs = Nothing
            
        End If
        
    Else
        MsgBox "Select the search criteria.", vbExclamation
    End If
Else
    MsgBox "Select the search field.", vbExclamation
End If

End Sub

Private Sub cmdRef_Click()
Call SRefresh
End Sub

Private Sub cmdSelect_Click()
Sel = ""
    If lstSearch.ListItems.Count > 0 Then
        Sel = lstSearch.SelectedItem
        Me.Visible = False
       ' Me.Unload Me
    Else
        MsgBox "No record selected.", vbExclamation
    End If
Unload Me
End Sub

Private Sub Form_Load()

    With frmsearchusergroups.lstSearch
        .ListItems.Clear
                
        .Columnheaders.Clear
        .Columnheaders.Add , , "Group ID", 1500
        .Columnheaders.Add , , "Group Name", 1500
       
        .View = lvwReport
        .Gridlines = True
    End With
    
    With frmsearchusergroups.cboField
        .AddItem "Group ID"
        .AddItem "Group Name"
    End With
    
'    CConnect.cnnConnect
    Set rs = oSaccoMaster.GetRecordSet("SELECT GroupID,GroupName FROM  UserGroups")
    'oSaccoMaster.GetRecordSet (sql)
    
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            While Not .EOF
        
 Set li = frmsearchusergroups.lstSearch.ListItems.Add(, , !groupId)
                li.SubItems(1) = !GroupName
                'li.SubItems(2) = !ContactPerson & ""
              
            .MoveNext
            Wend
            
            
        End If
        .Close
    End With
    
    Set rs = Nothing
'    Set cnnPayroll = Nothing
        
    cboCrieria.Text = cboCrieria.List(0)
    cboField.Text = cboField.List(0)
    'txtFrom.SetFocus
    
    Me.Top = (Screen.Height - Height) / 2
    Me.Left = (Screen.Width - Width) / 1.4
    
End Sub

Private Sub txtFrom_Change()
If txtFrom.Text = "" Then
        cmdFind.Enabled = False
    Else
        cmdFind.Enabled = True
    End If
End Sub

Private Sub txtTo_Change()
   If Len(Trim(txtTo.Text)) > 20 Then
        Beep
        MsgBox "Can't enter more than 20 characters", vbExclamation
        KeyAscii = 8
    End If
  Select Case KeyAscii
    'Case Asc("vbBack")
    Case Asc("A") To Asc("Z")
    Case Asc("a") To Asc("z")
    Case Asc("0") To Asc("9")
    Case Asc("/")
    Case Asc("-")
    Case Asc("(")
    Case Asc(")")
    Case Asc(" ")
    Case Asc(".")
    Case Is = 8
    
    Case Else
    Beep
    KeyAscii = 0
  End Select
End Sub
Public Sub SRefresh()
lstSearch.ListItems.Clear

'    CConnect.cnnConnect
    sql = "SELECT GroupName,Groupid FROM  UserGroups"
    oSaccoMaster.GetRecordSet (sql)
    
       
    With rs
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
 Set li = frmsearchusergroups.lstSearch.ListItems.Add(, , !groupId)
                li.SubItems(1) = !GroupName & ""
         
                .MoveNext
                
            Loop
            
        End If
        .Close
    End With
    
    Set rs = Nothing
'    Set cnnPayroll = Nothing
    
    txtFrom.Text = ""
    txtTo.Text = ""
    cboCrieria.Text = "="
    cboField.Text = "GroupID"
End Sub
