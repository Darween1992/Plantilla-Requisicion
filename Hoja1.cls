VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
    

Private Sub ComboArticulos_Change()
 Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationSemiautomatic
End Sub

Private Sub ComboArticulos_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
 Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationSemiautomatic
End Sub

