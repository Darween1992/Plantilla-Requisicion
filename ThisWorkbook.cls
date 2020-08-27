VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True



Private Sub Workbook_Open()

'Application.DisplayAlerts = True
'Application.ScreenUpdating = True
'Application.Calculation = xlCalculationSemiautomatic

Dim HOY As Date
HOY = Now

MsgBox ("Saludos, ¡NO OLVIDE SELECIONAR SU CENTRO DE TRABAJO ! ")
MsgBox ("Hoy es " & HOY & " NO ELVIDE PASAR LA REQUISICION ANTES DE LOS 10 PRIMEROS DIAS DEL MES")
'Sheets("BBDD1").Visible = True
'Sheets("Requisicion").Visible = False
'Sheets("BBDD1").Select

Call bloquear_hoja

End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)


'MsgBox ("Gracias Por Guardar Cambios")
'Sheets("BBDD1").Visible = False

'Sheets("Requisicion").Visible = True



'Sheets("Requisicion").Select

End Sub

