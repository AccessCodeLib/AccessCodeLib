VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphvizDOTnode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: GraphvizDOTnode (Josef Pötzl)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Graphviz-Knoten
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/graphviz/GraphvizDOTnode.cls</file>
'  <use>api/graphviz/GraphvizDOTgraph.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

'Private Const p_conDonStringMinBreakLen As Long = 0 '16

Public ID As String
Public shape As String
Public style As String

Public Fillcolor As String
Public Fontname As String
Public Fontcolor As String
Public Fontsize As Long
Public MiscOptions As String
Public SubGraph As GraphvizDOTgraph
Public HtmlExport As Boolean
Public URL As String

Private p_Label As String
Private m_color As String
Private p_LabelDOTstring As String

Public Property Let LayoutMode(LayoutID As Long)

   Dim strColor As String
   Dim strFillColor As String
   Dim strFontName As String
   Dim strFontColor As String
   Dim rst As DAO.Recordset
   Set rst = CodeDb.OpenRecordset("select DotShape, DotStyle, DotColor, DotFillcolor, DotFontName, DotFontcolor from ADF_Graphviz_NodeLayout Where idNodeLayout=" & LayoutID)
   With rst
   If Not .EOF Then
      shape = Nz(.Fields(0), vbNullString)
      style = Nz(.Fields(1), vbNullString)
      strColor = Nz(.Fields(2), vbNullString)
      strFillColor = Nz(.Fields(3), vbNullString)
      strFontName = Nz(.Fields(4), vbNullString)
      strFontColor = Nz(.Fields(5), vbNullString)
   End If
   .Close
   End With
   Set rst = Nothing
   
   If Len(strFillColor) > 0 Then
      Fillcolor = strFillColor
   End If
   If Len(strColor) > 0 Then
      Color = strColor
   End If
   If Len(strFontName) > 0 Then
      Fontname = strFontName
   End If
   If Len(strFontColor) > 0 Then
      Fontcolor = strFontColor
   End If

End Property

Public Property Get Label() As String
   Label = p_Label
End Property

Public Property Let Label(sLabel As String)
   p_Label = sLabel
   p_LabelDOTstring = OptimizeDotLabel(sLabel)
   p_LabelDOTstring = Replace(p_LabelDOTstring, vbNewLine, "  \n")
   p_LabelDOTstring = Replace(p_LabelDOTstring, """", "\""")
End Property

Public Property Get Color() As String
   Color = m_color
End Property

Public Property Let Color(sColor As String)
   m_color = sColor
   If Len(Fontcolor) = 0 Then
      Fontcolor = sColor
   End If
End Property


Public Function DOTstring(Optional bNoSubGraph As Boolean = False) As String

   Dim strDOT As String
   Dim strTemp As String
   'Dim strTempLabelName As String
   
   If Len(shape) > 0 Then
      strTemp = strTemp & "shape=" & shape & ","
   End If
   If Len(p_Label) > 0 Then
      strTemp = strTemp & "label=""" & p_LabelDOTstring & "  ""," 'Leerzeichen für bessere Lesbarkeit
   End If
   If Len(Color) > 0 Then
      strTemp = strTemp & "color=""" & Color & ""","
   End If
   If Len(Fillcolor) > 0 Then
      strTemp = strTemp & "fillcolor=""" & Fillcolor & ""","
   End If
   If Len(style) > 0 Then
      strTemp = strTemp & "style=" & style & ","
   End If
   If Len(Fontname) > 0 Then
      strTemp = strTemp & "fontname=""" & Fontname & ""","
   End If
   If Fontsize > 0 Then
      strTemp = strTemp & "fontsize=" & Fontsize & ","
   End If
   If Len(Fontcolor) > 0 Then
      strTemp = strTemp & "fontcolor=""" & Fontcolor & ""","
   End If
   
   If Len(URL) > 0 Then
      strTemp = strTemp & "URL=""" & URL & ""","
   ElseIf HtmlExport Then
      strTemp = strTemp & "URL=""" & ID & EXPORTFILEEXTENSION & ""","
   End If

   If Len(MiscOptions) > 0 Then
      strTemp = strTemp & MiscOptions & ","
   End If
   
   If Len(strTemp) > 0 Then
      strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
   End If
   
   If SubGraph Is Nothing Then
      strDOT = ID
      If Len(strTemp) > 0 Then
         strDOT = strDOT & " [" & strTemp & "]"
      End If
      strDOT = strDOT & ";" & vbNewLine
   ElseIf bNoSubGraph Then
      strDOT = SubGraph.DOTstring(bNoSubGraph) & vbNewLine
   Else
      strDOT = _
         "subgraph cluster_" & ID & _
         "{" & _
         ID & "[style=invis,shape=point,color=white];"
      If Len(strTemp) > 0 Then
         strDOT = strDOT & Replace(strTemp, ",", ";") & ";"
      End If
      If Len(Fontcolor) > 0 Then
         strDOT = strDOT & "fontcolor=""" & Fontcolor & """;"
      ElseIf Len(Color) > 0 Then
         strDOT = strDOT & "fontcolor=""" & Color & """;"
      End If
      
      'nodes einstellen
      strTemp = vbNullString
      If Len(shape) > 0 Then
         strTemp = strTemp & "shape=" & shape & ","
      End If
      If Len(p_Label) > 0 Then
         strTemp = strTemp & "label=""" & p_LabelDOTstring & ""","
      End If
      If Len(Color) > 0 Then
         strTemp = strTemp & "color=""" & Color & ""","
      End If
      'If Len(Fontname) > 0 Then
      '   strTemp = strTemp & "fontname=""" & Fontname & ""","
      'End If
      If Len(strTemp) > 0 Then
         strDOT = strDOT & "node [" & strTemp & "];" & vbNewLine
      End If
      
      strDOT = strDOT & SubGraph.DOTstring & vbNewLine & _
            "};"
   End If
   strDOT = strDOT & vbNewLine
   
   DOTstring = strDOT

End Function
