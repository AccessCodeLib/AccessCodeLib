VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphvizStructViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: _codelib.addins.StructViewer.GraphvizStructViewer
'---------------------------------------------------------------------------------------
'
' GraphvizStructViewer
'
' Remarks:
'
'  @todo Setting of Graphviz parameters is still missing (there are still too many constants)
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/StructViewer/GraphvizStructViewer.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/shared/CodeModuleReader.cls</use>
'  <use>_codelib/addins/shared/CodeModuleReader.cls</use>
'  <use>api/graphviz/GraphvizDOT.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const conDefaultMaxLevel As Long = 3

Private Const conDefaultGraphName As String = "DOT"

Private Const conDefaultFontname As String = "Verdana"

Private Const conObjGrpColor0 As String = "#0000FF" 'blau
Private Const conObjGrpColor1 As String = "#993399" 'lila
Private Const conObjGrpColor2 As String = "#996600" 'braun
Private Const conObjGrpColor3 As String = "#666600" 'grün

Private Const conObjGrpExtObjColor As String = "#006600" 'grün
Private Const conObjGrpExtObjColorOnlyInputOrOutput As String = "#FF0000" '"red"
Private Const conObjGrp_FirstGroupColor As String = "#000000"

Private Const conObjGrpStyleObj As String = ""
Private Const conObjGrpStyleObjGrp As String = "solid" '"dashed" '"bold"
Private Const conObjGrpFontnameFu As String = "Arial" '"Arial bold"
Private Const conDefaultObjGrpFontsizeFu As Long = 8 '10
Private Const conDefaultObjGrpFontsizeObjGrpDiff As Long = 2 'diese Zahlt wird zu Fu-Size addiert

Private Const conDOT_GraphProp As String = _
                  "graph [bgcolor=""#FFFFFF"", margin=0.05,overlap=false, splines=true, concentrade=true, fontname=""" & conDefaultFontname & """, ratio=compress];" & _
                  "node[color=""#000000"",fontname=""" & conDefaultFontname & """];" & _
                  "edge[color=""#000000"", arrowsize=0.8, labelangle=60, labelfontsize=" & conDefaultObjGrpFontsizeFu & ",labelfontname=""" & conDefaultFontname & """];"
               

Private m_DOT As GraphvizDOT

Private m_CodeModuleStructReader As CodeModuleStructReader

Private m_BaseCodeModuleName As String
Private m_RefreshStruct As Boolean

Private Sub Class_Initialize()
   Set m_CodeModuleStructReader = New CodeModuleStructReader
End Sub

Public Property Get BaseCodeModuleName() As String

   BaseCodeModuleName = m_CodeModuleStructReader.BaseCodeModuleName
   
End Property

Public Property Let BaseCodeModuleName(ByVal CmName As String)

   m_CodeModuleStructReader.BaseCodeModuleName = CmName
   
End Property

Public Sub ReadStruct()

   m_CodeModuleStructReader.ReadStruct
   
End Sub

Private Function ModulReaderExistsInCollection(col As Collection, cmr As CodeModuleReader) As Boolean
   Dim TempModuleReader As CodeModuleReader
   For Each TempModuleReader In col
      If TempModuleReader.Name = cmr.Name Then
         ModulReaderExistsInCollection = True
         Exit For
      End If
   Next
End Function


Public Sub CreateDOT( _
      Optional ByVal GraphMode As DOT_GraphMode = DOT_GraphMode.GraphMode_DOT, _
      Optional ByVal RankDir As DOT_RankDir = DOT_RankDir.RankDir_LR, _
      Optional ByVal GraphName As String = vbNullString, _
      Optional ByVal RndSeed As Long = 0, _
      Optional ByVal markFirstItem As Boolean = False)
   
   If m_CodeModuleStructReader.CodeModuleReaderCollection.Count = 0 Or m_RefreshStruct Then
      ReadStruct
   End If
   
   Set m_DOT = New GraphvizDOT
   
   m_DOT.Init GraphMode, RankDir, GraphName, , conDOT_GraphProp
   m_DOT.RndSeed = RndSeed
   
   addDOTobjects markFirstItem
   m_DOT.ToBinaryGraph DOT2PNG, True
   
End Sub

Private Sub addDOTobjects(ByVal markFirstItem As Boolean)
   
   Dim collectionItem As CodeModuleReader
   
   Dim strNodeID As String
   Dim strColor As String
   Dim strStyle As String
   
   If markFirstItem Then
      strColor = "blue"
   End If
   For Each collectionItem In m_CodeModuleStructReader.CodeModuleReaderCollection

      strNodeID = collectionItem.Name
      If collectionItem.Name = "modErrorHandler" Then
         strColor = "gray"
      End If
      m_DOT.AddNode strNodeID, collectionItem.Name, , strColor, strStyle, , , GetLink(collectionItem)
      m_DOT.Node(strNodeID).LayoutMode = collectionItem.CodeModule.Parent.Type
      
      If Len(strColor) > 0 Then strColor = vbNullString
      
      addDOTconnections collectionItem

   Next
   
End Sub

Private Function GetLink(cmr As CodeModuleReader) As String
   
   Dim strName As String
   Dim i As Long
   
   strName = Replace(cmr.Name, "_", "__")

   i = 2
   Do While i <= Len(strName)
      If StrComp(Mid(strName, i, 1), UCase(Mid(strName, i, 1)), vbBinaryCompare) = 0 _
         And StrComp(Mid(strName, i, 1), LCase(Mid(strName, i, 1)), vbBinaryCompare) <> 0 Then
         strName = Left(strName, i - 1) & "_" & LCase(Mid(strName, i, 1)) & Mid(strName, i + 1)
         i = i + 1
      End If
      i = i + 1
   Loop
   
   Select Case cmr.CodeModule.Parent.Type
      Case vbext_ct_StdModule
         strName = "namespace" & strName
      Case vbext_ct_ClassModule
         strName = "class_" & strName
      Case Else
         If Left(strName, 6) = "Form__" Then
            strName = Mid(strName, 7) & "_8frm"
         ElseIf Left(strName, 8) = "Report__" Then
            strName = Mid(strName, 9) & "_8rep"
         End If
         ' lassen
   End Select
   
   GetLink = LCase(strName) & ".html"
   
End Function

Private Sub addDOTconnections(ByRef cmr As CodeModuleReader)

   Dim TempModuleReader As CodeModuleReader

   For Each TempModuleReader In cmr.RequiredModules
      If TempModuleReader.Name <> "modErrorHandler" Then  ' Errorhandler ignorieren
         'm_DOT.AddConnection cmr.Name, tempModuleReader.Name
         m_DOT.AddConnection TempModuleReader.Name, cmr.Name
      End If
   Next

End Sub


'Private Function getDotShape(cli As CodeLibInfo) As String
'
'   Select Case cli.Type
'      Case CodeLibElementType.clet_ClassModule
'         getDotShape = "octagon"
'      Case CodeLibElementType.clet_StdModule
'         getDotShape = "box"
'      Case CodeLibElementType.clet_Form
'         getDotShape = "octagon"
'      Case CodeLibElementType.clet_Report
'         getDotShape = "invhouse"
'   End Select
'
'End Function

'Private Function getDotColor(cli As CodeLibInfo) As String
'
'   Select Case cli.Type
'      Case CodeLibElementType.clet_Form
'         getDotColor = "#FFFFCC"
'      Case CodeLibElementType.clet_Report
'         getDotColor = "#FFFFCC"
'      Case Else
'         getDotColor = vbNullString
'   End Select
'
'End Function

'Private Function getDotStyle(cli As CodeLibInfo) As String
'
'   Select Case cli.Type
'      Case CodeLibElementType.clet_Form
'         getDotStyle = "filled"
'      Case CodeLibElementType.clet_Report
'         getDotStyle = "filled"
'      Case Else
'         getDotStyle = vbNullString
'   End Select
'
'End Function

#If EARLYBINDING Then
Public Property Set CurrentVbProject(ref As VBIDE.VBProject)
#Else
Public Property Set CurrentVbProject(ref As Object)
#End If

   Set m_CodeModuleStructReader.CurrentVbProject = ref

End Property

#If EARLYBINDING Then
Public Property Get CurrentVbProject() As VBIDE.VBProject
#Else
Public Property Get CurrentVbProject() As Object
#End If

   Set CurrentVbProject = m_CodeModuleStructReader.CurrentVbProject

End Property
