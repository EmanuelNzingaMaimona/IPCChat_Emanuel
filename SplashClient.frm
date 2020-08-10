VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmSplashClient 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleMode       =   0  'User
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5640
      Top             =   1560
   End
   Begin MSComctlLib.ProgressBar pgb 
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   2640
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "IPCChat"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   735
      Left            =   2160
      TabIndex        =   7
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "IPCChat"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   975
      Left            =   2400
      TabIndex        =   6
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "   Microluz"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   735
      Left            =   15
      TabIndex        =   5
      Top             =   120
      Width           =   6705
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   5415
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "beta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Carregando"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2018 Todos direitos reservados "
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3615
      Left            =   6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmSplashClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
End
End If

End Sub

Private Sub Timer1_Timer()
If pgb.Value = 100 Then
Unload Me
frmLoginClient.Show 1
Timer1.Enabled = False
Else

pgb.Value = pgb.Value + 2
End If
Label3.Caption = "Carregando " & pgb.Value & "%"
End Sub

