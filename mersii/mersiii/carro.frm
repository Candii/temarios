VERSION 5.00
Begin VB.Form carro 
   ClientHeight    =   6360
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame comandos 
      Caption         =   "COMANDOS"
      Height          =   1935
      Left            =   6000
      TabIndex        =   14
      Top             =   0
      Width           =   3015
      Begin VB.CommandButton eliminar 
         Caption         =   "ELIMINAR"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton agregar 
         Caption         =   "AGREGAR"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame principal 
      Caption         =   "PRINCIPAL"
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   1680
         TabIndex        =   6
         Top             =   5040
         Width           =   3135
         Begin VB.CommandButton left 
            DownPicture     =   "carro.frx":0000
            DragIcon        =   "carro.frx":08CA
            Height          =   615
            Left            =   0
            Picture         =   "carro.frx":1194
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton rigth 
            DownPicture     =   "carro.frx":1A5E
            DragIcon        =   "carro.frx":2328
            Height          =   615
            Left            =   1560
            Picture         =   "carro.frx":2BF2
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.TextBox modelo 
         DataField       =   "MODELO"
         DataSource      =   "data1"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   5
         Top             =   3840
         Width           =   2775
      End
      Begin VB.Data data1 
         Caption         =   "BASE DE DATOS"
         Connect         =   "Access"
         DatabaseName    =   "E:\carros\empresa de carros.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CARROS"
         Top             =   6120
         Width           =   1380
      End
      Begin VB.TextBox linea 
         DataField       =   "LINEA"
         DataSource      =   "data1"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   4
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox marca 
         DataField       =   "MARCA"
         DataSource      =   "data1"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   3
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox id 
         DataField       =   "ID_PLACA"
         DataSource      =   "data1"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox color 
         DataField       =   "COLOR"
         DataSource      =   "data1"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   1
         Top             =   3000
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "LINEA"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "MARCA"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "COLOR"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "MODELO"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   9
         Top             =   3720
         Width           =   2655
      End
   End
   Begin VB.Menu menuprincipal 
      Caption         =   "MENU PRINCIPAL"
   End
   Begin VB.Menu volver 
      Caption         =   "VOLVER"
   End
   Begin VB.Menu salir 
      Caption         =   "SALIR"
   End
End
Attribute VB_Name = "carro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub agregar_Click()
If data1.Recordset.EOF = False And data1.Recordset.BOF = False Then
    data1.Recordset.AddNew
    data1.Recordset("ID_PLACA") = id.Text
    data1.Recordset("MARCA") = marca.Text
    data1.Recordset("LINEA") = linea.Text
    data1.Recordset("COLOR") = color.Text
    data1.Recordset("MODELO") = modelo.Text
    data1.Recordset.Update
    End If
End Sub



Private Sub eliminar_Click()
If data1.Recordset.EOF = False And data1.Recordset.BOF = False Then
    data1.Recordset.Delete
    data1.Recordset.Requery
    End If
End Sub
Private Sub left_Click()
 data1.Recordset.MovePrevious
If data1.Recordset.BOF = True Then
    data1.Recordset.MoveLast
 End If

End Sub

Private Sub mprincipal_Click()
inicio.Show
Me.Hide
End Sub


Private Sub menuprincipal_Click()
inicio.Show
Me.Hide
End Sub

Private Sub rigth_Click()
data1.Recordset.MoveNext
If data1.Recordset.EOF = True Then
data1.Recordset.MoveFirst
End If
End Sub


Private Sub salir_Click()
End
End Sub

Private Sub volver_Click()
registros.Show
Me.Hide
End Sub

