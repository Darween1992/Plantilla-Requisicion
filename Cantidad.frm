VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cantidad 
   Caption         =   "Cantidad "
   ClientHeight    =   2000
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   3100
   OleObjectBlob   =   "Cantidad.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Cantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub registro()

Dim cantidad As String

cantidad = (TextBox1)
 
Sheets("Requisicion").Range("E8").Value = cantidad


Call registro_completo

Unload Me


End Sub


Private Sub CommandButton1_Click()
Call registro
End Sub


