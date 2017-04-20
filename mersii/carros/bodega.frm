VERSION 5.00
Begin VB.Form bodega 
   ClientHeight    =   8970
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   Picture         =   "bodega.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame comandos 
      Caption         =   "COMANDOS"
      Height          =   1935
      Left            =   6000
      TabIndex        =   14
      Top             =   0
      Width           =   3015
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
         TabIndex        =   16
         Top             =   240
         Width           =   2415
      End
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
         TabIndex        =   15
         Top             =   960
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
      Begin VB.TextBox repuestos 
         DataField       =   "REPUESTOS"
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
         TabIndex        =   8
         Top             =   3000
         Width           =   2775
      End
      Begin VB.TextBox vehiculo 
         DataField       =   "VEHICULO"
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
         TabIndex        =   7
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox accesorios 
         DataField       =   "ACCESORIOS"
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
         TabIndex        =   6
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox herramientas 
         DataField       =   "HERRAMIENTAS"
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
         Top             =   2160
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
         RecordSource    =   "BODEGA"
         Top             =   6120
         Width           =   1380
      End
      Begin VB.TextBox pintura 
         DataField       =   "PINTURA"
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
         Top             =   3840
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   1680
         TabIndex        =   1
         Top             =   5040
         Width           =   3135
         Begin VB.CommandButton rigth 
            DownPicture     =   "bodega.frx":0342
            DragIcon        =   "bodega.frx":0C0C
            Height          =   615
            Left            =   1560
            Picture         =   "bodega.frx":14D6
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton left 
            DownPicture     =   "bodega.frx":1DA0
            DragIcon        =   "bodega.frx":266A
            Height          =   615
            Left            =   0
            Picture         =   "bodega.frx":2F34
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Label Label1 
         Caption         =   "PINTURA"
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
         TabIndex        =   13
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "repuestos"
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
         TabIndex        =   12
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "ACCESORIOS"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   14.25
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
      Begin VB.Label Label5 
         Caption         =   "VEHICULO"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "HERRAMIENTAS"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   2160
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
Attribute VB_Name = "bodega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub agregar_Click()
If data1.Recordset.EOF = False And data1.Recordset.BOF = False Then
    data1.Recordset.AddNew
    data1.Recordset("VEHICULO") = vehiculo.Text
    data1.Recordset("ACCESORIOS") = accesorios.Text
    data1.Recordset("HERRAMIENTAS") = herramientas.Text
    data1.Recordset("REPUESTOS") = repuestos.Text
    data1.Recordset("PINTURA") = pintura.Text
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


