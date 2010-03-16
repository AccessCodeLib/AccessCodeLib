VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphvizDOTconnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: GraphvizDOTconnection (Josef Pötzl)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Graphviz-Verbindung zw. zwei Knoten
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/graphviz/GraphvizDOTconnection.cls</file>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

Private Const conDefaultLabelFontsize As Long = 8
Private Const conLabelFontname As String = "Veranda"

Public idFrom As String
Public idTo As String
Public headlabel As String
Public taillabel As String
Public Label As String
Public optionally As Boolean

Public Function DOTstring() As String
   Dim strTemp As String
   If Len(headlabel) > 0 Then
      strTemp = strTemp & "headlabel=""" & headlabel & """"
   End If
   If Len(taillabel) > 0 Then
      If Len(strTemp) > 0 Then
         strTemp = strTemp & ","
      End If
      strTemp = strTemp & "taillabel=""" & taillabel & """"
   End If
   If Len(Label) > 0 Then
      If Len(strTemp) > 0 Then
         strTemp = strTemp & ","
      End If
      strTemp = strTemp & "label=""" & Label & """"
   End If
   
   If optionally Then
      If Len(strTemp) > 0 Then
         strTemp = strTemp & ","
      End If
      strTemp = strTemp & "style=dashed"
   End If
   
   
   If Len(strTemp) > 0 Then
      strTemp = "[fontsize=" & conDefaultLabelFontsize & ",labelangle=60,labelfontsize=" & conDefaultLabelFontsize & ",labelfontname=""" & conLabelFontname & """," & strTemp & "]"
   End If
   DOTstring = idFrom & "->" & idTo & strTemp & ";"
End Function
