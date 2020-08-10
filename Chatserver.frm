VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmchatserver 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IPCChat-Servidor"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   11220
   Icon            =   "Chatserver.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   Picture         =   "Chatserver.frx":0442
   ScaleHeight     =   4665
   ScaleWidth      =   11220
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox mh 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   -4320
      ScaleHeight     =   1905
      ScaleWidth      =   5445
      TabIndex        =   42
      Top             =   -1440
      Visible         =   0   'False
      Width           =   5475
      Begin VB.Shape Shape10 
         BorderColor     =   &H00808000&
         BorderWidth     =   12
         FillColor       =   &H0000C000&
         Height          =   1935
         Left            =   0
         Top             =   0
         Width           =   5175
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Comando inválido! Histórico vazio."
         BeginProperty Font 
            Name            =   "Script MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   1200
         TabIndex        =   45
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Informação de erro"
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
         Height          =   495
         Left            =   840
         TabIndex        =   44
         Top             =   120
         Width           =   3615
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
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
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   3600
         TabIndex        =   43
         Top             =   1440
         Width           =   1110
      End
      Begin VB.Image Image7 
         Height          =   345
         Left            =   600
         Picture         =   "Chatserver.frx":B395
         Stretch         =   -1  'True
         Top             =   720
         Width           =   435
      End
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   8400
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox mst 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   -5280
      ScaleHeight     =   1905
      ScaleWidth      =   5445
      TabIndex        =   38
      Top             =   -120
      Visible         =   0   'False
      Width           =   5475
      Begin VB.Image Image6 
         Height          =   345
         Left            =   720
         Picture         =   "Chatserver.frx":B7D7
         Stretch         =   -1  'True
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
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
         ForeColor       =   &H80000005&
         Height          =   345
         Left            =   3600
         TabIndex        =   41
         Top             =   1440
         Width           =   1110
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Informação de  estado "
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
         Height          =   495
         Left            =   840
         TabIndex        =   40
         Top             =   120
         Width           =   3615
      End
      Begin VB.Label lblst 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Você está desconectado"
         BeginProperty Font 
            Name            =   "Script MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   360
         TabIndex        =   39
         Top             =   720
         Width           =   4695
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H00808000&
         BorderWidth     =   12
         FillColor       =   &H0000C000&
         Height          =   1935
         Left            =   0
         Top             =   0
         Width           =   5175
      End
   End
   Begin VB.PictureBox mmu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   -6000
      ScaleHeight     =   2265
      ScaleWidth      =   6345
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Label Label24 
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
         Left            =   3240
         TabIndex        =   37
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label23 
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
         Left            =   1560
         TabIndex        =   36
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mudanndo o usuário"
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
         Left            =   1440
         TabIndex        =   35
         Top             =   120
         Width           =   3255
      End
      Begin VB.Image Image5 
         Height          =   465
         Left            =   720
         Picture         =   "Chatserver.frx":BC19
         Stretch         =   -1  'True
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Todos os dados não salvos serão apagados! Deseja continuar?"
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
         Height          =   855
         Left            =   1440
         TabIndex        =   34
         Top             =   720
         Width           =   4815
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00808000&
         BorderWidth     =   12
         FillColor       =   &H0000C000&
         Height          =   2295
         Left            =   0
         Top             =   0
         Width           =   6135
      End
   End
   Begin VB.PictureBox Pb 
      BackColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   4695
      Left            =   0
      Picture         =   "Chatserver.frx":C05B
      ScaleHeight     =   4635
      ScaleWidth      =   11235
      TabIndex        =   29
      Top             =   6000
      Width           =   11295
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   120
         Top             =   240
      End
      Begin VB.Timer Timer3 
         Interval        =   1000
         Left            =   240
         Top             =   1440
      End
      Begin VB.Timer Timer4 
         Interval        =   30
         Left            =   480
         Top             =   2520
      End
      Begin VB.Timer Timer5 
         Interval        =   30
         Left            =   1200
         Top             =   2520
      End
      Begin MSComctlLib.ProgressBar pg 
         Height          =   375
         Left            =   360
         TabIndex        =   30
         Top             =   1480
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label ipc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ipcchat"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1590
         Left            =   2640
         TabIndex        =   32
         Top             =   1200
         Width           =   5805
      End
      Begin VB.Label logo 
         BackStyle       =   0  'Transparent
         Caption         =   "Emanuel Nzinga Maimona/Copyright 2018"
         BeginProperty Font 
            Name            =   "Vivaldi"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   31
         Top             =   4250
         Width           =   5535
      End
   End
   Begin MSComctlLib.ProgressBar pst 
      Height          =   255
      Left            =   3480
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   3
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   120
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   480
      Left            =   1560
      TabIndex        =   26
      Top             =   3840
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   847
      _Version        =   393217
      BackColor       =   -2147483648
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"Chatserver.frx":16FAE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgb 
      Left            =   7560
      Top             =   -240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   880
      ImageHeight     =   663
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Chatserver.frx":1702B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Chatserver.frx":42D05
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Chatserver.frx":4DC68
            Key             =   ""
            Object.Tag             =   "img3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Chatserver.frx":65092
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Chatserver.frx":7FFD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Chatserver.frx":F5BCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Chatserver.frx":131895
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Chatserver.frx":14349B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Chatserver.frx":158CC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Chatserver.frx":16F7F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Chatserver.frx":1BA6B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Chatserver.frx":1DAC02
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Chatserver.frx":1E9BA8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wsk 
      Left            =   8280
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   888
   End
   Begin VB.PictureBox mg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   4440
      ScaleHeight     =   2265
      ScaleWidth      =   6345
      TabIndex        =   23
      Top             =   720
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox Thelp 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Script MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Text            =   "Chatserver.frx":1FB898
         Top             =   600
         Width           =   5415
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Ver vídeo"
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
         Left            =   3120
         TabIndex        =   47
         Top             =   1750
         Width           =   1335
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00808000&
         BorderWidth     =   12
         FillColor       =   &H0000C000&
         Height          =   2295
         Left            =   0
         Top             =   0
         Width           =   6135
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Guia de usuário"
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
         Left            =   1440
         TabIndex        =   25
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Voltar"
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
         Left            =   4560
         TabIndex        =   24
         Top             =   1750
         Width           =   1340
      End
   End
   Begin VB.PictureBox mc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   -6240
      ScaleHeight     =   2265
      ScaleWidth      =   6345
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "IPC 2018"
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
         Height          =   375
         Left            =   600
         TabIndex        =   22
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto: 947221912"
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
         Height          =   375
         Left            =   600
         TabIndex        =   21
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Voltar"
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
         Left            =   4560
         TabIndex        =   20
         Top             =   1750
         Width           =   1340
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Créditos"
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
         Left            =   1440
         TabIndex        =   19
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Desenvolvedor: Emanuel Nzinga Maimona"
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
         Height          =   375
         Left            =   600
         TabIndex        =   18
         Top             =   720
         Width           =   5175
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00808000&
         BorderWidth     =   12
         FillColor       =   &H0000C000&
         Height          =   2295
         Left            =   0
         Top             =   0
         Width           =   6135
      End
   End
   Begin VB.PictureBox mp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   -5760
      ScaleHeight     =   2265
      ScaleWidth      =   6345
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright IPC 2018"
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
         Height          =   375
         Left            =   600
         TabIndex        =   16
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Versão: 1.0"
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
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Voltar"
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
         Left            =   4560
         TabIndex        =   14
         Top             =   1750
         Width           =   1340
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sobre o Programa"
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
         Height          =   495
         Left            =   1440
         TabIndex        =   13
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nome: IpcChat"
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
         Height          =   375
         Left            =   600
         TabIndex        =   12
         Top             =   720
         Width           =   4575
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808000&
         BorderWidth     =   12
         FillColor       =   &H0000C000&
         Height          =   2295
         Left            =   0
         Top             =   0
         Width           =   6135
      End
   End
   Begin VB.PictureBox msg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   -6520
      ScaleHeight     =   2265
      ScaleWidth      =   6345
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808000&
         BorderWidth     =   12
         FillColor       =   &H0000C000&
         Height          =   2295
         Left            =   0
         Top             =   0
         Width           =   6135
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Obrigado pelo uso do Nosso aplicativo Você tem certeza que deseja sair?"
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
         Height          =   855
         Left            =   1080
         TabIndex        =   10
         Top             =   720
         Width           =   4575
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   480
         Picture         =   "Chatserver.frx":1FBB7E
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label2 
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
         Height          =   495
         Left            =   1440
         TabIndex        =   9
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label Label3 
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
         Left            =   1680
         TabIndex        =   8
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label4 
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
         Left            =   3360
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
      End
   End
   Begin VB.TextBox tre 
      Appearance      =   0  'Flat
      Height          =   3495
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   240
      Width           =   8055
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   2400
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      Height          =   135
      Left            =   8840
      Shape           =   2  'Oval
      Top             =   800
      Width           =   135
   End
   Begin VB.Label lblchat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "198.168.100.111"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   9000
      TabIndex        =   27
      Top             =   720
      Width           =   1815
   End
   Begin VB.Image Image3 
      Height          =   600
      Left            =   960
      Picture         =   "Chatserver.frx":1FBFC0
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   240
      Picture         =   "Chatserver.frx":1FC402
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Actualizar"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   8760
      TabIndex        =   4
      Top             =   3555
      Width           =   2175
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Desligar"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8760
      TabIndex        =   3
      Top             =   3870
      Width           =   2175
   End
   Begin VB.Shape Shape4 
      Height          =   735
      Left            =   8760
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Chat"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8760
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   " Enviar"
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
      Left            =   7200
      TabIndex        =   1
      Top             =   3900
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   3975
      Left            =   8760
      Top             =   240
      Width           =   2175
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   0
      Picture         =   "Chatserver.frx":1FC844
      Stretch         =   -1  'True
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label sta 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Estado de comunicação: desconetado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -120
      TabIndex        =   0
      Top             =   4440
      Width           =   11415
   End
   Begin VB.Menu mn 
      Caption         =   "Menu"
      Begin VB.Menu bk 
         Caption         =   "Bloquear"
         Shortcut        =   ^B
      End
      Begin VB.Menu p 
         Caption         =   "Personalizar"
         Begin VB.Menu ft 
            Caption         =   "Fonte"
         End
         Begin VB.Menu pf 
            Caption         =   "Cor de fundo"
         End
         Begin VB.Menu cl 
            Caption         =   "Cor de letras"
         End
      End
      Begin VB.Menu mu 
         Caption         =   "Mudar de usuário"
      End
      Begin VB.Menu ht 
         Caption         =   "Histórico"
         Begin VB.Menu gh 
            Caption         =   "Guardar Histórico"
            Shortcut        =   ^G
         End
         Begin VB.Menu lh 
            Caption         =   "Limpar Histórico"
            Shortcut        =   ^L
         End
      End
   End
   Begin VB.Menu help 
      Caption         =   "Ajuda"
      Begin VB.Menu sb 
         Caption         =   "Sobre"
         Begin VB.Menu ap 
            Caption         =   "Aplicativo"
         End
         Begin VB.Menu is 
            Caption         =   "Informação do Sistema"
         End
      End
      Begin VB.Menu cd 
         Caption         =   "Créditos"
      End
      Begin VB.Menu gu 
         Caption         =   "Guia de Utilização"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu exit 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "frmchatserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim txt As TextBox
Dim st As Integer

Private Sub ap_Click()
If Pb.Visible = True Then

Else
Beep
mp.Top = 960
mp.Left = 2400
mp.Width = 6155
mp.Visible = True
End If

End Sub

Private Sub bk_Click()
If Pb.Visible = True Then

Else
Timer2.Enabled = True
Timer3.Enabled = True
Pb.Top = 0
Pb.Visible = True
Pb.SetFocus
End If
End Sub

Private Sub cd_Click()
If Pb.Visible = True Then

Else
Beep
mc.Top = 960
mc.Left = 2400
mc.Width = 6155
mc.Visible = True
End If
End Sub

Private Sub cl_Click()
If Pb.Visible = True Then

Else
cmd.ShowColor
tre.ForeColor = cmd.Color
End If
End Sub











Private Sub ft_Click()
If Pb.Visible = True Then

Else
cmd.ShowFont
tre.Font = cmd.FontName
tre.FontSize = cmd.FontSize
End If
End Sub

Private Sub gh_Click()
If tre.Text = "" Then
mh.Top = 1200
mh.Left = 2900
mh.Width = 5175
mh.Visible = True
Else
cmd.DialogTitle = "Salvando histórico"
cmd.Filter = "Arquivo Texto(*.TXT)|*.TXT"
cmd.FileName = "Histórico"
cmd.ShowSave
Open cmd.FileName For Output As #1
Print #1, "                                    ''IPCChat v1.0''" & vbCrLf & tre.Text
Close #1
End If
End Sub

Private Sub Image3_Click()
On Error GoTo n
Beep
Shell (App.Path & "\tabtip.exe")
n:
MsgBox "Ocorreu um erro." & vbCrLf & "O teclado táctil não é compatível com o tipo de sistema!!!", vbCritical + vbOKOnly, "Erro"
End Sub

Private Sub exit_Click()
If Pb.Visible = True Then

Else
Beep
msg.Top = 960
msg.Left = 2400
msg.Width = 6155
msg.Visible = True
End If

End Sub

Private Sub Form_Load()
Pb.Visible = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
wsk.Close
st = 0
Timer6.Enabled = True
tre.Text = ""
lblchat.Caption = ""
mst.Visible = False
mst.Top = 1200
mst.Left = 2900
mst.Width = 5175

End Sub

Private Sub gu_Click()
If Pb.Visible = True Then

Else
Beep
mg.Top = 960
mg.Left = 2400
mg.Width = 6155
mg.Visible = True
End If
End Sub

Private Sub is_Click()
If Pb.Visible = True Then

Else
Shell ("msinfo32.exe")
End If
End Sub

Private Sub Label11_Click()
mp.Visible = False
End Sub

Private Sub Label16_Click()
mc.Visible = False
End Sub

Private Sub Label21_Click()
mg.Visible = False
End Sub

Private Sub Label23_Click()
wsk.Close
Unload Me
frmLoginServer.Show 1
End Sub

Private Sub Label24_Click()
mmu.Visible = False
End Sub

Private Sub Label25_Click()
mh.Visible = False
End Sub

Private Sub Label27_Click()
mst.Visible = False
End Sub

Private Sub Label3_Click()
End
End Sub

Private Sub Label30_Click()
Beep
frmvideo.Show 1
End Sub

Private Sub Label4_Click()
msg.Visible = False

End Sub

Private Sub Label6_Click()
If st = 0 Or wsk.State = 0 Or wsk.State = 8 Then
Beep
mst.Visible = True
Exit Sub
ElseIf st = 3 Then
Beep
mst.Visible = True
Exit Sub
Else

wsk.SendData lognome & ">> " & Text1.Text & " //recebida às " & Left(Time, 5)

If tre.Text = "" Then
tre.Text = lognome & ">> " & Text1.Text & " //enviada às " & Left(Time, 5) & tre.Text
Else
tre.Text = lognome & ">> " & Text1.Text & " //enviada às " & Left(Time, 5) & vbCrLf & tre.Text
End If
Text1.Text = ""
End If
End Sub

Private Sub Label8_Click()
Beep
wsk.Close
st = 0
lblchat.Caption = ""
End Sub

Private Sub Label9_Click()
If st = 0 Then
Beep
st = 3
wsk.Close
wsk.Listen
Else
Beep
mst.Visible = True
End If
End Sub

Private Sub lh_Click()
If Pb.Visible = True Then

Else
tre.Text = ""
End If
End Sub



Private Sub mu_Click()
If Pb.Visible = True Then

Else
Beep
mmu.Top = 960
mmu.Left = 2400
mmu.Width = 6155
mmu.Visible = True

End If
End Sub

Private Sub Pb_DblClick()
Pb.Visible = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
logo.Left = 0
End Sub

Private Sub Pb_KeyPress(KeyAscii As Integer)
Pb.Visible = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
logo.Left = 0
End Sub

Private Sub Pb_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Pb.Visible = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
logo.Left = 0
End Sub

Private Sub pf_Click()
If Pb.Visible = True Then

Else
cmd.ShowColor
tre.BackColor = cmd.Color
End If
End Sub

Private Sub Text1_GotFocus()
Text1.BackColor = &H80000000
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Beep
Label6_Click
'keycode = 8

End If
End Sub

Private Sub Text1_LostFocus()
Text1.BackColor = vbWhite
End Sub

Private Sub Timer1_Timer()
If st = 1 Then
sta.ForeColor = vbGreen
sta.Caption = "Estado de comunicação: conetado"
lblst.Caption = "Você está conectado!"
ElseIf st = 0 Then
sta.ForeColor = vbRed
sta.Caption = "Estado de comunicação: desconetado"
lblst.Caption = "Você está desconectado!"
ElseIf st = 3 Then
sta.ForeColor = vbYellow
sta.Caption = "Estado de comunicação: procurando..."
lblst.Caption = "Aguardando conexão"
End If


If tre.BackColor = vbBlack Then tre.ForeColor = vbWhite

If wsk.State = 7 Then
st = 1
lblchat.Caption = Chatnome
ElseIf wsk.State = 8 Then
st = 0
Timer6.Enabled = True
lblchat.Caption = ""
End If

End Sub

Private Sub Timer2_Timer()
Select Case pg.Value
Case 10
ipc.ForeColor = vbYellow
Case 20
ipc.ForeColor = vbWhite
Case 30
ipc.ForeColor = vbRed
Case 40
ipc.ForeColor = vbBlue
Case 50
ipc.ForeColor = vbCyan
Case 60
ipc.ForeColor = vbGreen
Case 70
ipc.ForeColor = vborange
Case 80
ipc.ForeColor = vbMagenta
Case 90
ipc.ForeColor = &HFF8080
Case 100
ipc.ForeColor = vbBlack
End Select


If logo.Left = 0 Then
Timer4.Enabled = True
Timer5.Enabled = False
ElseIf logo.Left = 5640 Then
Timer5.Enabled = True
Timer4.Enabled = False
End If


End Sub

Private Sub Timer3_Timer()
If pg.Value = 100 Then
pg.Value = 0
Else
pg.Value = pg.Value + 10
End If
End Sub


Private Sub Timer4_Timer()
logo.Left = logo.Left + 20
End Sub

Private Sub Timer5_Timer()
logo.Left = logo.Left - 20
End Sub

Private Sub Timer6_Timer()

If pst.Value = 3 Then

wsk.Close
wsk.Listen
st = 3
pst.Value = 0
Timer6.Enabled = False
Else
pst.Value = pst.Value + 1
End If

End Sub

Private Sub Timer7_Timer()

End Sub

Private Sub wsk_ConnectionRequest(ByVal requestID As Long)
If wsk.State <> closed Then wsk.Close

wsk.Accept requestID
st = 1
lblchat.Caption = Chatnome
wsk.SendData lognome
End Sub

Private Sub wsk_DataArrival(ByVal bytesTotal As Long)
Dim dados As String
wsk.GetData dados
If Len(dados) <= 22 Then
Chatnome = dados
Else
If tre.Text = "" Then
tre.Text = dados + tre.Text
Else
tre.Text = dados + vbCrLf & tre.Text
End If
End If
End Sub

