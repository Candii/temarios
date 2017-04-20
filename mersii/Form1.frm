VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   2760
      TabIndex        =   19
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   9240
      TabIndex        =   17
      Top             =   8760
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   9240
      TabIndex        =   16
      Top             =   8160
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   9240
      TabIndex        =   15
      Top             =   7560
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PAGAR"
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "FINALIZAR"
      Height          =   495
      Left            =   4680
      TabIndex        =   10
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton BORRAR 
      Caption         =   "BORRAR"
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AÑADIR"
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   7560
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid lista 
      Height          =   3495
      Left            =   360
      TabIndex        =   7
      Top             =   3720
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   15
      Cols            =   6
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   2760
      Width           =   6375
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   1920
      Width           =   6375
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   18
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   14
      Top             =   8760
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "I.V.A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   13
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "SUBTOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   12
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "FACTURA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   6
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DETALLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BORRAR_Click()
TOT = TOT - X
Text4 = TOT
FILA = FILA - 1
lista.Col = 1
lista.Row = FILA
lista.Text = ""
lista.Col = 2
lista.Row = FILA
lista.Text = ""
lista.Col = 3
lista.Row = FILA
lista.Text = ""
lista.Col = 4
lista.Row = FILA
lista.Text = ""
lista.Col = 5
lista.Row = FILA
lista.Text = ""
End Sub

Private Sub Command1_Click()
lista.Col = 1
lista.Row = FILA
lista.Text = Text7.Text

lista.Col = 2
lista.Row = FILA
lista.Text = Text1.Text

lista.Col = 3
lista.Row = FILA
lista.Text = Text2.Text

lista.Col = 4
lista.Row = FILA
lista.Text = Text3.Text

X = Val(Text1) * Val(Text3)
lista.Col = 5
lista.Row = FILA
lista.Text = X
TOT = TOT + X
Text4 = TOT
BORRAR_TEXTO
FILA = FILA + 1
End Sub

Private Sub Command3_Click()
End
End Sub


Private Sub Command4_Click()
Text5 = TOT * 0.12
Text6 = TOT + Text5
End Sub

Private Sub Form_Load()
lista.ColWidth(0) = 10

lista.Col = 1
lista.Row = 0
lista.Text = "CODIGO"

lista.Col = 2
lista.Row = 0
lista.Text = "CANTIDAD"

lista.ColWidth(3) = 5000
lista.ColAlignment(3) = 5
lista.Col = 3
lista.Row = 0
lista.Text = "D E T A L L E"

lista.Col = 4
lista.Row = 0
lista.Text = "PRECIO"

lista.Col = 5
lista.Row = 0
lista.Text = "SUBTOTAL"

FILA = 1
End Sub

