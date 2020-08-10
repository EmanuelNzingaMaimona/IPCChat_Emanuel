VERSION 5.00
Begin VB.Form frmlauch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inicializador IpcChat"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Launch.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox msg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4425
      ScaleWidth      =   7185
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.Image Image4 
         Height          =   375
         Left            =   240
         Picture         =   "Launch.frx":0442
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   480
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sair"
         BeginProperty Font 
            Name            =   "Script MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Top             =   3800
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808000&
         BorderWidth     =   12
         FillColor       =   &H0000C000&
         Height          =   4335
         Left            =   0
         Top             =   0
         Width           =   6855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "                                                                                  Iniciar como Cliente"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   -120
         TabIndex        =   3
         Top             =   2640
         Width           =   6855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "                                                                                 Iniciar como Servidor"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   -120
         TabIndex        =   2
         Top             =   1560
         Width           =   6855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Como deseja iniciar a aplicação?"
         BeginProperty Font 
            Name            =   "Script MT Bold"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   1680
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmlauch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Image4_Click()
Beep
frmvideo.Show 1
End Sub

Private Sub Label1_Click()
Unload Me
frmSplashServer.Show 1
End Sub

Private Sub Label3_Click()
Unload Me
frmSplashClient.Show 1
End Sub

Private Sub Label4_Click()
Beep
End
End Sub
