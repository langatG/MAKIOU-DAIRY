VERSION 5.00
Begin VB.Form frmcreditors 
   Caption         =   "Form1"
   ClientHeight    =   4170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      DataField       =   "CompanyName"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   1260
      MaxLength       =   20
      TabIndex        =   6
      Top             =   540
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      DataField       =   "Address"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   1260
      MaxLength       =   40
      TabIndex        =   5
      Top             =   1980
      Width           =   6255
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      DataField       =   "Phone"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   1260
      MaxLength       =   16
      TabIndex        =   4
      Top             =   1500
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      DataField       =   "SupplierID"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   1260
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      DataField       =   "ContactPerson"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   1260
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1020
      Width           =   2415
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      DataField       =   "Email"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   5100
      MaxLength       =   20
      TabIndex        =   1
      Top             =   540
      Width           =   2415
   End
   Begin VB.PictureBox Picture5 
      Height          =   255
      Left            =   3660
      Picture         =   "frmcreditors.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   180
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Company Name"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   540
      Width           =   1395
   End
   Begin VB.Label Label4 
      Caption         =   "Address"
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   1980
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Phone Number"
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   1500
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "Supplier Code"
      Height          =   375
      Index           =   5
      Left            =   0
      TabIndex        =   9
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Contact Person"
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   8
      Top             =   1020
      Width           =   1395
   End
   Begin VB.Label Label2 
      Caption         =   "E-mail Address"
      Height          =   375
      Index           =   6
      Left            =   3900
      TabIndex        =   7
      Top             =   540
      Width           =   1395
   End
End
Attribute VB_Name = "frmcreditors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
