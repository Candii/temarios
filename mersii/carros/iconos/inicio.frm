VERSION 5.00
Begin VB.Form inicio 
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame inicio 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton salir 
         Caption         =   "SALIR"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3600
         TabIndex        =   5
         Top             =   4680
         Width           =   2655
      End
      Begin VB.CheckBox mostrar 
         Caption         =   "MOSTRA LA CONTRASEÑA"
         Height          =   495
         Left            =   3480
         TabIndex        =   4
         Top             =   4080
         Width           =   2655
      End
      Begin VB.CommandButton iniciar 
         Caption         =   "ENTRAR"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   3
         Top             =   4680
         Width           =   2655
      End
      Begin VB.TextBox txtpass 
         Alignment       =   2  'Center
         DataField       =   "Password"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3600
         PasswordChar    =   "*"
         TabIndex        =   2
         Text            =   "123"
         Top             =   3480
         Width           =   2535
      End
      Begin VB.TextBox txtuser 
         Alignment       =   2  'Center
         DataField       =   "Usuarios"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   1
         Text            =   "Axel"
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "LAMBORGÜINI"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         TabIndex        =   8
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   7
         Top             =   3240
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "USUARIO"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   6
         Top             =   2400
         Width           =   2655
      End
   End
End
Attribute VB_Name = "inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public con As Integer
Private Sub mostrar_Click()
con = con + 1
If (con / 2) = Int((con / 2)) Then
txtpass.PasswordChar = "*"
Else
txtpass.PasswordChar = ""
End If
End Sub

Private Sub iniciar_Click()
If txtuser.Text = "Axel" And txtpass.Text = "123" Then
registros.Show
Me.Hide
'txtuser.Text = ""
'txtpass.Text = ""
Else
MsgBox "Usuario o Contraseña incorrecto", , "Error"
End If
End Sub

Private Sub salir_Click()
End
End Sub



