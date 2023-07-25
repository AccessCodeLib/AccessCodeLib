VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphvizDOTgraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: api.graphviz.GraphvizDOTgraph
'---------------------------------------------------------------------------------------
'
' Class for creating the graphviz graph definition
'
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/graphviz/GraphvizDOTgraph.cls</file>
'  <use>api/graphviz/GraphvizDOTnode.cls</use>
'  <use>api/graphviz/GraphvizDOTconnection.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

Private p_colNodes As Collection
Private p_colConnections As Collection

Private Sub Class_Initialize()
   Set p_colNodes = New Collection
   Set p_colConnections = New Collection
End Sub

Private Sub Class_Terminate()
   Set p_colNodes = Nothing
   Set p_colConnections = Nothing
End Sub

Public Property Get DOTstring(Optional NoSubGraph As Boolean = False) As String
   DOTstring = getDOTstring(NoSubGraph)
End Property

Public Property Get Node(vIndex As Variant) As GraphvizDOTnode
   Set Node = p_colNodes(vIndex)
End Property

Public Property Get Nodes() As Collection
   If p_colNodes Is Nothing Then
      Set p_colNodes = New Collection
   End If
   Set Nodes = p_colNodes
End Property

Public Sub AddNode(ByVal sID As String, _
               Optional ByVal sLabel As String = vbNullString, _
               Optional ByVal sShape As String = vbNullString, _
               Optional ByVal sColor As String = vbNullString, _
               Optional ByVal sStyle As String = vbNullString, _
               Optional ByVal sMiscOptions As String = vbNullString, _
               Optional ByVal oSubGraph As GraphvizDOTgraph = Nothing, _
               Optional ByVal bHtmlExport As Boolean = False, _
               Optional ByVal sHtmlFileName As String = vbNullString, _
               Optional ByVal sHtmlFilePath As String = vbNullString, _
               Optional ByVal sHtmlExternalLink As String = vbNullString)
               
   Dim n As GraphvizDOTnode
   Set n = New GraphvizDOTnode
   With n
      .ID = sID
      .Label = sLabel
'      If Len(sLabel) > 0 Then
'         .Label = sLabel
'      Else
'         .Label = sID
'      End If
      If Len(sShape) > 0 Then
         .shape = sShape
      Else
         .shape = "box"
      End If
      If Len(sColor) > 0 Then
         .Color = sColor
      End If
      If Len(sStyle) > 0 Then
         .style = sStyle
      End If
      
      If Len(sHtmlExternalLink) > 0 Then
         .URL = sHtmlExternalLink
      End If
      
      .HtmlExport = bHtmlExport
      If bHtmlExport Then
         If Len(sHtmlFileName) > 0 Then
            .URL = sHtmlFileName
         End If
      End If
      
      If Len(sMiscOptions) > 0 Then
         .MiscOptions = sMiscOptions
      End If
      Set .SubGraph = oSubGraph
   End With
   p_colNodes.Add n, n.ID
   
End Sub

Public Sub AddConnection(sFrom As String, sTo As String, _
                   Optional sLabel As String = vbNullString, _
                   Optional sHeadLabel As String = vbNullString, _
                   Optional sTailLabel As String = vbNullString, _
                   Optional bOptionalObj As Boolean = False)

   Dim c As GraphvizDOTconnection
   Set c = New GraphvizDOTconnection
   With c
      .idFrom = sFrom
      .idTo = sTo
      .Label = sLabel
      .headlabel = sHeadLabel
      .taillabel = sTailLabel
      .optionally = bOptionalObj
   End With
   p_colConnections.Add c
   
End Sub

Private Function getDOTstring(Optional NoSubGraph As Boolean = False) As String

   Dim strDOT As String
   Dim strTemp As String
   Dim i As Long
   Dim lngMaxCount As Long
   
   lngMaxCount = p_colNodes.Count
   For i = 1 To lngMaxCount
      strDOT = strDOT & p_colNodes(i).DOTstring(NoSubGraph)
   Next i
   
   lngMaxCount = p_colConnections.Count
   For i = 1 To lngMaxCount
      strDOT = strDOT & p_colConnections(i).DOTstring
   Next i
   
   getDOTstring = strDOT

End Function
