Attribute VB_Name = "sub_funciones"
Option Explicit



    
  
Sub buscarV()


Dim contenido As String
Dim codigo As Variant
Dim saldo_iteam As Variant
Dim por_llegar As Variant

contenido = Sheets("Requisicion").Range("B8").Value

codigo = Sheets("Requisicion").Range("F8").Value


saldo_iteam = Application.VLookup(codigo, Sheets("BBDD1").Range("A3:D1000"), 4, False)

por_llegar = Application.VLookup(contenido, Sheets("Ultimo pedido").Range("A2:J100"), 4, False)


If IsError(saldo_iteam) Then

saldo_iteam = 0
End If

If IsError(por_llegar) Then

por_llegar = 0

End If

 'MsgBox "Usted Tiene Un Saldo De " & saldo_iteam & "    " & contenido & " En Bodega ", vbOKOnly + vbExclamation, "Saldo Almacen"
 
' MsgBox " En la Requisicion Anterior pidio " & por_llegar, vbExclamation, " Ultimo Pedido "
 



End Sub


    


Sub bloquear_hoja()

Sheets("Requisicion").Protect "123"

End Sub



Sub desbloquear_hoja()

Sheets("Requisicion").Unprotect "123"


End Sub





Sub borrar_canv()



Sheets("Requisicion").Select
 
Range("E12").End(xlDown).Offset(1, 0).Select

Selection.Offset(0, 1).ClearContents

Selection.Offset(0, 2).ClearContents

Selection.Offset(0, 3).ClearContents

Selection.Offset(0, 4).ClearContents

Selection.Offset(0, 5).ClearContents




End Sub

Sub copiaRegistros2()

Dim tipo As String

tipo = Sheets("Requisicion").Range("I8").Value


Sheets("Requisicion").Select


Range("B8:J8").Copy

Range("B11").End(xlDown).Offset(1, 0).Activate


Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    

If tipo = "SERVICIO" Then

    Range("B11").End(xlDown).Offset(0, 9).Value = Application.InputBox("Por Favor Justifique Su Pedido")
    
Else
  
    Range("B11").End(xlDown).Offset(0, 9).Value = Application.InputBox("Desea Ingresar Observaciones")
   
End If

    
    'Call ultimo_registro
    
    Range("B8").Select
    Selection.ClearContents
    Range("E8").Select
    Selection.ClearContents
    Range("B8").Select


' Range("B13:K1000").Select
 '   ActiveWorkbook.Worksheets("Requisicion").Sort.SortFields.Clear
  '  ActiveWorkbook.Worksheets("Requisicion").Sort.SortFields.Add2 Key:=Range( _
   '     "B13:B446"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    '    xlSortNormal
'    With ActiveWorkbook.Worksheets("Requisicion").Sort
 '       .SetRange Range("B13:K446")
  '      .Header = xlGuess
   '     .MatchCase = False
    '    .Orientation = xlTopToBottom
     '   .SortMethod = xlPinYin
      '  .Apply
 '
 ' ISERTAR FUNCION PARA ORDENAR QUEDA PENDIENTE POR ORDENAR
 
 
 
' End With
    Range("I10").Select
 
 Call borrar_canv
 Call bloquear_hoja
 
End Sub

'Sub copy_ultimo_pedido()

'Call desbloquear_hoja

'Sheets("Ultimo pedido").Visible = True


'Sheets("Ultimo pedido").Select
'Range("A2:J1000").ClearContents


'Sheets("Requisicion").Select
'Range("B13:K1000").Copy

'Sheets("Ultimo pedido").Select
'Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
 '       :=False, Transpose:=False
  '  Application.CutCopyMode = False


'Sheets("Requisicion").Select

'Sheets("Ultimo pedido").Visible = False
'Call bloquear_hoja

'End Sub
Sub ultimo_registro()




Sheets("Requisicion").Select
Range("B8").Copy

Sheets("Granjas").Visible = xlSheetVisible
Sheets("Granjas").Select
Range("I5").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
Sheets("Granjas").Visible = xlSheetVeryHidden
Sheets("Requisicion").Select
End Sub

Sub activaCalculos()

 Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationSemiautomatic
End Sub

Sub desactivaCalculos()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


End Sub



Sub BORRAR_ULTIMO_REGISTRO()


Load UserForm1
UserForm1.Show

End Sub

