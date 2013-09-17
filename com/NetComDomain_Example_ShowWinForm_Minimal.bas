Attribute VB_Name = "NetComDomain_Example_ShowWinForm_Minimal"
'---------------------------------------------------------------------------------------
' Class Module: NetComDomain_Example_ShowWinForm_Minimal
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Beispiel zur Verwendung der NetComDomain Klasse.
'               Es wird ein einfaches .NET Winform erzeugt und angezeigt.
' \ingroup COM
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>COM/NetComDomain_Example_ShowWinForm_Minimal.bas</file>
'  <use>COM/NetComDomain.cls</use>
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
