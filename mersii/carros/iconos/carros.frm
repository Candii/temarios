VERSION 5.00
Begin VB.Form registros 
   ClientHeight    =   4545
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10905
   LinkTopic       =   "Form2"
   ScaleHeight     =   4545
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu salir 
      Caption         =   "SALIR"
   End
   Begin VB.Menu incio 
      Caption         =   "INCIO"
   End
   Begin VB.Menu personal 
      Caption         =   "PERSONAL"
   End
   Begin VB.Menu clientes 
      Caption         =   "CLIENTES"
   End
   Begin VB.Menu bodegadrepuestos 
      Caption         =   "BODEGA DE REPUESTOS"
   End
   Begin VB.Menu tallerdservicios 
      Caption         =   "TALLER DE SERVICIOS"
   End
   Begin VB.Menu carros 
      Caption         =   "CARROS"
   End
End
Attribute VB_Name = "registros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bodegadrepuestos_Click()
bodega.Show
Me.Hide
End Sub

Private Sub carros_Click()
carro.Show
Me.Hide
End Sub

Private Sub clientes_Click()
cliente.Show
Me.Hide
End Sub

Private Sub incio_Click()
inicio.Show
Me.Hide
End Sub

Private Sub personal_Click()
perso.Show
Me.Hide
End Sub

Private Sub salir_Click()
End
End Sub

Private Sub tallerdservicios_Click()
taller.Show
Me.Hide
End Sub
