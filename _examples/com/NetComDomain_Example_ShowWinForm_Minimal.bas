Attribute VB_Name = "NetComDomain_Example_ShowWinForm_Minimal"
'---------------------------------------------------------------------------------------
' Package: _examples.com.NetComDomain_Example_ShowWinForm_Minimal
'---------------------------------------------------------------------------------------
'
' Example for using the NetComDomain class.
' A simple .NET WinForm is created and shown.
'
' Author:
'     Sten Schmidt
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_examples/com/NetComDomain_Example_ShowWinForm_Minimal.bas</file>
'  <use>com/NetComDomain.cls</use>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Public Sub NetComDomain_Example_ShowWinForm_Minimal()

    'Pfad ggf. Anpassen
    Dim DllPath As String
        DllPath = "C:\Windows\Microsoft.NET\Framework\v2.0.50727\System.Windows.Forms.dll"

    Dim mylib As NetComDomain
    Set mylib = New NetComDomain

    'Instanz der Form-Klasse über LateBinding, da keine TLB als Verweis hinzugefügt wurde
    Dim form As Object
    Set form = mylib.CreateObject("Form", "System.Windows.Forms", DllPath)
    
    form.Text = "Das ist ein .NET Winform"
    form.StartPosition = 1
    form.ShowInTaskbar = True
    form.ShowIcon = False
    
    form.ShowDialog
        
    Set form = Nothing

End Sub
