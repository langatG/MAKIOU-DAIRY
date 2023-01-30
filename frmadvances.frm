VERSION 5.00
Begin VB.Form frmadvances 
   Caption         =   "Advances Report"
   ClientHeight    =   1275
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   5325
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Dairy Trans Advance"
      Height          =   615
      Left            =   1320
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Fsa Sup Advance"
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fsa Trans Advance"
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dairy Sup Advance"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmadvances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    On Error Resume Next
     reportname = "DairyAdvance.rpt"
     Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    STRFORMULA = ""
    reportname = "TransportersAdvance.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    STRFORMULA = ""
    reportname = "SUPPLIERSMILKADVANCE.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, title
End Sub

Private Sub Command4_Click()
    On Error Resume Next
     STRFORMULA = ""
     reportname = "DairyAdvance.rpt"
     Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub
