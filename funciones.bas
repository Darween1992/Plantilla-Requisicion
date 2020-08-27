Attribute VB_Name = "funciones"
Option Explicit

Sub registrar_articulo()
 
 'comprovacion de iteam
 
 Sheets("Requisicion").Select
 
Call desbloquear_hoja
 
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


Dim CELDA As Range
Dim contenido As String
Dim comprovacion As Byte

contenido = Range("B8").Value


Dim I As Byte

Dim FINAL
FINAL = Sheets("Requisicion").Range("K7").Value

Range("B11").Select

Dim COMPROVAR As Byte
For I = 1 To FINAL + 1




    If ActiveCell.Value = contenido Then
    comprovacion = 1


    MsgBox (" Ya Esta Registrado El Iteam " & contenido & " ")
    
    Range("B8").ClearContents
    Range("E8").ClearContents
    
    Exit Sub
    
    Else
    
    COMPROVAR = 1
    End If
      

ActiveCell.Offset(1, 0).Select
 
Next I
 
 
 If COMPROVAR = 1 Then
 
Call buscarV
      
       Call copiaRegistros2
       
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationSemiautomatic
      
End If
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationSemiautomatic



End Sub

Sub limpiar_plantilla()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


Dim respuesta As Byte
Dim titulo As String

titulo = "Mensaje de Aviso"
respuesta = MsgBox("Se Eliminaran todos los Registros?" & Chr(13) & vbNewLine & " ¿Desea Continuar?", vbQuestion + vbYesNo, titulo)

If respuesta = vbYes Then


Sheets("Requisicion").Select

Call desbloquear_hoja
Range("B13:L300").ClearContents

'Range("B13:K300").Select
 '
  '  Selection.ClearFormats
   ' Range("E11").Select

MsgBox ("Registros Eliminados Con Exito")



Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationSemiautomatic


Call bloquear_hoja

Else

MsgBox ("No se Borro Ningun Registro")

End If
End Sub



Sub generar_requisicion()




Dim nombre_documento As String

Dim mes As String

Dim centro_trabajo As String


mes = Sheets("Granjas").Range("H1").Value

centro_trabajo = Sheets("Requisicion").Range("C5").Value

nombre_documento = "Resquisicion " & centro_trabajo & " " & mes





    Columns("A:S").Select
    Selection.Copy
    Workbooks.Add
    Columns("A:A").Select
    ActiveSheet.Paste
    Range("A3").Select
    
    Sheets(1).Protect ("123")
    
   
    
    
    Application.CutCopyMode = False
    ActiveWorkbook.SaveAs Filename:=Workbooks("FormatoRequisicion.xlsm").Path & "\" & nombre_documento, FileFormat _
        :=xlOpenXMLWorkbook, CreateBackup:=False
    

    
    ActiveWindow.Close
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 1
    Range("A2").Select
    
    
    
    MsgBox nombre_documento & " Se guardo con exito en la carpeta actual "
    
End Sub

'Sub BORRAR_ULTIMO_REGISTRO()

'Dim celda As Range
'Dim buscando As String
'buscando = Sheets("Granjas").Range("I5").Value

'Sheets("Requisicion").Select

'Call desbloquear_hoja


'For Each celda In Range("B13:B100")


'If celda.Value = buscando Then

'celda.Select

'Range(Selection, Selection.Offset(0, 10)).ClearContents


'End If
'Next celda

 'ActiveWindow.SmallScroll Down:=-3
'    Range("B13:K1000").Select
 '   ActiveWorkbook.Worksheets("Requisicion").Sort.SortFields.Clear
  '  ActiveWorkbook.Worksheets("Requisicion").Sort.SortFields.Add2 Key:=Range( _
   '     "B13:B1000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'        xlSortNormal
 '   With ActiveWorkbook.Worksheets("Requisicion").Sort
  '      .SetRange Range("B13:K1000")
   '     .Header = xlGuess
    '    .MatchCase = False
     '   .Orientation = xlTopToBottom
      '  .SortMethod = xlPinYin
 '       .Apply
  '  End With
  '  Range("E8").Select

'MsgBox "El Iteam " & buscando & "Se elimino con Exito ", vbInformation, " Ultimo Registro Borrado "

'Call bloquear_hoja

'ThisWorkbook.RefreshAll
'End Sub

Sub filtrarIteams()

Load FILTROS
   FILTROS.Show


End Sub

Sub restablecer_filtros()



End Sub
