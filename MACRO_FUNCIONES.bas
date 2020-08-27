Attribute VB_Name = "MACRO_FUNCIONES"
Option Explicit

Sub registro_completo()

Dim CENTRO As String
CENTRO = Sheets("Requisicion").Range("C5").Value

If CENTRO = "" Then

MsgBox "POR FAVOR SELECCIONE UN CENTRO DE TRABAJO "

Exit Sub
End If



Dim cantidad As Variant

Dim COMPROVAR As Double

Dim comprovar_web As Double


Dim saldoMantenimiento As Double
Dim saldoweb As Double

COMPROVAR = Sheets("Granjas").Range("B18").Value
comprovar_web = Sheets("Granjas").Range("E18").Value


saldoMantenimiento = Sheets("Granjas").Range("B9").Value


saldoweb = Sheets("Granjas").Range("E9").Value

If comprovar_web > saldoweb Then

MsgBox "Supera su presupuesto WEB de  " & saldoweb, vbExclamation, " Presupuesto Web Superado "

Call activaCalculos

Exit Sub

Call activaCalculos

End If

Call activaCalculos

If COMPROVAR > saldoMantenimiento Then

MsgBox ("Supera su presupuesto de mantenimiento de " & saldoMantenimiento)

Call activaCalculos

Exit Sub
Call activaCalculos
Else

Call activaCalculos

cantidad = Sheets("Requisicion").Range("E8").Value

'On Error GoTo ManejadoErrores

Select Case cantidad

    Case Is = ""
    
    Call activaCalculos


    MsgBox "Falta Registrar Cantidad"
    
    Call activaCalculos
    ThisWorkbook.RefreshAll
    Exit Sub

Call activaCalculos

    Case Else
    
Call activaCalculos

    Call registrar_articulo

Call activaCalculos


End Select

Call activaCalculos

Exit Sub

ManejadoErrores:

MsgBox " Ha ocurrido un error"
 Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationSemiautomatic


End If




End Sub

Sub genera_completo()

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Sheets("Requisicion").Select

Dim CENTRO As Variant

CENTRO = Sheets("Requisicion").Range("C5").Value

On Error GoTo ManejadoErrores

If CENTRO = "" Then

MsgBox "Falta Registrar El Centro De Trabajo"

Call activaCalculos
Exit Sub


Else

Call generar_requisicion
 
'Call copy_ultimo_pedido

End If

Call activaCalculos

Exit Sub

ManejadoErrores:

MsgBox " Ha ocurrido un error"

 Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationSemiautomatic

End Sub

