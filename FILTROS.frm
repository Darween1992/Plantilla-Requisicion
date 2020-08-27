VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FILTROS 
   Caption         =   "FILTRAR ITEAMS"
   ClientHeight    =   6840
   ClientLeft      =   130
   ClientTop       =   450
   ClientWidth     =   6910
   OleObjectBlob   =   "FILTROS.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FILTROS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub filtrarformu()

Dim filtro As String

 Application.DisplayAlerts = False
Application.ScreenUpdating = False


Sheets("Articulos").Visible = xlSheetVisible
Sheets("R_filtro").Visible = xlSheetVisible

filtro = ListBox1.List(ListBox1.ListIndex, 0)



 Sheets("R_filtro").Select
    ActiveSheet.Range("$A$2:$K$2975").AutoFilter Field:=4, Criteria1:=filtro
    
       
 'Cells.Select
 
 Range("A1").CurrentRegion.Select
 
 
Sheets("Articulos").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Sheets("R_filtro").Select
    Selection.Copy
    Sheets("Articulos").Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("R_filtro").Select
    Application.CutCopyMode = False
    Sheets("Requisicion").Select
    Range("J10").Select

Application.DisplayAlerts = True
Application.ScreenUpdating = True




Sheets("Articulos").Visible = xlSheetVeryHidden
Sheets("R_filtro").Visible = xlSheetVeryHidden


End Sub

Private Sub CommandButton1_Click()

Call filtrarformu


End Sub



Private Sub CommandButton2_Click()

Unload Me
Call restablecer_filtro
End Sub

Sub restablecer_filtro()

 Application.DisplayAlerts = False
Application.ScreenUpdating = False



Sheets("Articulos").Visible = xlSheetVisible
Sheets("R_filtro").Visible = xlSheetVisible

 Sheets("R_filtro").Select
    ActiveSheet.Range("$A$1:$I$2975").AutoFilter Field:=4
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Articulos").Select
    Range("A1").Select
    ActiveSheet.PasteSpecial Format:=3, Link:=1, DisplayAsIcon:=False, _
        IconFileName:=False
    Sheets("R_filtro").Select
    Range("A1").Select
    Sheets("Requisicion").Select
    Range("I10").Select

Sheets("Articulos").Visible = xlSheetVeryHidden
Sheets("R_filtro").Visible = xlSheetVeryHidden


End Sub

Private Sub CommandButton3_Click()
Call restablecer_filtro
End Sub

Sub registrar_desde_filtro()

Dim nuevo_registro As String
nuevo_registro = ListBox2.List(ListBox2.ListIndex, 0)

Sheets("Requisicion").Range("B8").Value = nuevo_registro

Load cantidad
cantidad.Show



End Sub


Private Sub CommandButton4_Click()
Call registrar_desde_filtro

End Sub



Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Terminate()
Call restablecer_filtro
End Sub
