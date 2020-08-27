VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ELIMINAR REGISTRO"
   ClientHeight    =   5880
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7350
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ELIMINAR()

Dim buscando As String

buscando = ListBox1.List(ListBox1.ListIndex, 0)


Sheets("Requisicion").Select

Call desbloquear_hoja


For Each CELDA In Range("B13:B200")


If CELDA.Value = buscando Then

CELDA.Select

Selection.EntireRow.Select

Selection.Delete
Range("A1").Activate

End If
Next CELDA

Call bloquear_hoja
End Sub


Private Sub CommandButton1_Click()
Call ELIMINAR
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub


Private Sub UserForm_Click()

End Sub
