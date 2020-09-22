VERSION 5.00
Begin VB.Form frmPrompt 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DataSource Type"
   ClientHeight    =   1440
   ClientLeft      =   5175
   ClientTop       =   9345
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optDataSource 
      Caption         =   "DAO (Access Database *.mdb)"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Value           =   -1  'True
      Width           =   2895
   End
   Begin VB.OptionButton optDataSource 
      Caption         =   "ADO (ODBC)"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.CheckBox chkDontAsk 
      Caption         =   "Don't Ask Me Again"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblField 
      AutoSize        =   -1  'True
      Caption         =   "Indicate the DataSource Type"
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
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2580
   End
End
Attribute VB_Name = "frmPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pCancel As Boolean

Private Sub cmdCancel_Click()
    pCancel = True
    Me.Visible = False
End Sub

Private Sub cmdOK_Click()
    pCancel = False
    Me.Visible = False
End Sub
