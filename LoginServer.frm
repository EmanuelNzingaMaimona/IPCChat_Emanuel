VERSION 5.00
Begin VB.Form frmLoginServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login-Servidor"
   ClientHeight    =   3510
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   6135
   Icon            =   "LoginServer.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ma 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   -3000
      ScaleHeight     =   2145
      ScaleWidth      =   4785
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Shape Shape3 
         BorderColor     =   &H00808000&
         BorderWidth     =   12
         FillColor       =   &H0000C000&
         Height          =   2175
         Left            =   0
         Top             =   0
         Width           =   4815
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Você não inseriu nenhum nome, deseja continuar?"
         BeginProperty Font 
            Name            =   "Script MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   720
         TabIndex        =   15
         Top             =   600
         Width           =   3975
      End
      Begin VB.Image Image3 
         Height          =   345
         Left            =   240
         Picture         =   "LoginServer.frx":0442
         Stretch         =   -1  'True
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Aviso"
         BeginProperty Font 
            Name            =   "Script MT Bold"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Sim"
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
         Left            =   1200
         TabIndex        =   13
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Não"
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
         Left            =   2640
         TabIndex        =   12
         Top             =   1560
         Width           =   1335
      End
   End
   Begin VB.PictureBox mls 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   -4920
      ScaleHeight     =   1905
      ScaleWidth      =   5025
      TabIndex        =   6
      Top             =   -1080
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Não"
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
         Left            =   2640
         TabIndex        =   10
         Top             =   1410
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Sim"
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
         Left            =   1200
         TabIndex        =   9
         Top             =   1410
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fechando o Programa"
         BeginProperty Font 
            Name            =   "Script MT Bold"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   120
         Width           =   3255
      End
      Begin VB.Image Image2 
         Height          =   345
         Left            =   240
         Picture         =   "LoginServer.frx":0884
         Stretch         =   -1  'True
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Você tem certeza que deseja sair?"
         BeginProperty Font 
            Name            =   "Script MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   720
         TabIndex        =   7
         Top             =   720
         Width           =   3975
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808000&
         BorderWidth     =   12
         FillColor       =   &H0000C000&
         Height          =   1935
         Left            =   0
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.PictureBox msg 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3585
      ScaleWidth      =   6345
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "PCEMA"
         Top             =   2280
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Script MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         MaxLength       =   13
         TabIndex        =   5
         Text            =   "Emanuel"
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Image Image4 
         Height          =   360
         Left            =   240
         Picture         =   "LoginServer.frx":0CC6
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nome de utilazador remoto"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   17
         Top             =   1920
         Width           =   4575
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
         ForeColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   4440
         TabIndex        =   4
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Script MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
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
         Height          =   615
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   3255
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   120
         Picture         =   "LoginServer.frx":1590
         Stretch         =   -1  'True
         Top             =   2880
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Introduza o nome de usuário"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   960
         Width           =   4575
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808000&
         BorderWidth     =   12
         FillColor       =   &H0000C000&
         Height          =   3495
         Left            =   0
         Top             =   0
         Width           =   6135
      End
   End
End
Attribute VB_Name = "frmLoginServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim posanterior As Integer




Private Sub Form_Load()
Text2.Text = frmchatserver.wsk.LocalHostName
mls.Top = 840
mls.Left = 650
mls.Visible = False
mls.Width = 4815

ma.Top = 650
ma.Left = 700
ma.Visible = False
ma.Width = 4815
Text1.Text = ""
End Sub



Private Sub Image4_Click()
Beep
Unload Me
frmlauch.Show 1
End Sub

Private Sub Label10_Click()
ma.Visible = False
End Sub

Private Sub Label11_Click()
lognome = frmchatserver.wsk.LocalHostName
Unload Me
frmchatserver.Show
End Sub

Private Sub Label3_Click()
If Text1.Text = "" Then
ma.Visible = True
Else
lognome = Text1.Text
Unload Me
frmchatserver.Show
End If
End Sub

Private Sub Label4_Click()
Beep
mls.Visible = True
End Sub

Private Sub Label8_Click()
End
End Sub

Private Sub Label9_Click()
mls.Visible = False
End Sub


