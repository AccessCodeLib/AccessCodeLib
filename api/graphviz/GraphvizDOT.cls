VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphvizDOT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: GraphvizDOT (Josef Pötzl)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' GraphvizDOT
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/graphviz/GraphvizDOT.cls</file>
'  <use>api/graphviz/GraphvizDOTgraph.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Text
Option Explicit

'EXE bzw. OCX verwenden
'Wenn EXE, dann müssen exe-dateien im App-Verzeichnung unter /bin liegen, falls Graphviz nicht installiert wurde
Private Const m_conUseEXE As Boolean = True

Public Enum DOT_GraphMode
   GraphMode_DOT = 0&
   GraphMode_TWOPI = 1&
   GraphMode_NEATO = 2&
End Enum

Public Enum DOT_RankDir
   RankDir_LR = 0&
   RankDir_TB = 1&
   RankDir_RL = 2&
   RankDir_BT = 3&
End Enum

Public Enum DOT_OutputGraphType
   DOT2GIF = 11&
   DOT2JPEG = 12&
   DOT2PNG = 13&
   DOT2EMF = 27&
End Enum

Private Const conDefaultGraphName As String = "DOT"

Private p_objDOT As Object
Private p_strDOT As String

Private p_DOTgraph As GraphvizDOTgraph

Private p_strGraphName As String
Private p_lngGraphMode As Long
Private p_RankDir As DOT_RankDir
Private p_strGraphLabel As String

Private p_strGraphProp As String
Private p_strFirstLevelNodeProp As String

Private p_InvisibleSubGraph As Boolean

Private p_RndSeed As Double

Private Sub Class_Initialize()

On Error GoTo HandleErr

   Set p_DOTgraph = New GraphvizDOTgraph
   p_strGraphName = conDefaultGraphName
   p_InvisibleSubGraph = False

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "Class_Initialize", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub Class_Terminate()
On Error Resume Next
   Set p_DOTgraph = Nothing
   Set p_objDOT = Nothing
End Sub

Public Property Get UseExe() As Boolean

On Error GoTo HandleErr

   UseExe = m_conUseEXE

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "UseExe", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Private Property Get CheckedObjWingraph() As Object

   Dim strLibFile As String
   Dim strObjName As String
   
On Error GoTo HandleErr

   strLibFile = CurrentProject.Path & "\lib\WinGraphviz.dll"
   
   If p_objDOT Is Nothing Then
      Select Case p_lngGraphMode
         Case 0
            'strObjName = "Wingraphviz.DOT" ' ...
            strObjName = "DOT"
         Case 1
            strObjName = "TWOPI"
         Case 2
            strObjName = "NEATO"
      End Select
      'Set p_objDOT = CreateObject(strObjName)
      'Set p_objDOT = modDirectCOM.GetInstance(strLibFile, strObjName)
   End If
   Set CheckedObjWingraph = p_objDOT

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "CheckedObjWingraph", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Property

Public Property Get DOTstring() As String

On Error GoTo HandleErr

   If Len(p_strDOT) = 0 Then
      CreateDOT
   End If
   DOTstring = p_strDOT

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "DOTstring", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Public Sub ClearDOT()
On Error Resume Next
   p_strDOT = vbNullString 'DOT-String zurücksetzen
   Set p_DOTgraph = Nothing
End Sub

Public Sub init(Optional GraphMode As DOT_GraphMode = 0, _
            Optional RankDir As DOT_RankDir = 0, _
            Optional GraphName As String = vbNullString, _
            Optional FirstNodeProp As String = vbNullString, _
            Optional GraphProp As String = vbNullString)

On Error GoTo HandleErr

   p_strDOT = vbNullString
   p_lngGraphMode = GraphMode
   p_strGraphName = GraphName
   p_RankDir = RankDir
   
   p_strFirstLevelNodeProp = FirstNodeProp
   p_strGraphProp = GraphProp

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "init", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Public Property Let RndSeed(sNewValue As Double)
On Error Resume Next
   p_RndSeed = sNewValue
End Property

Public Property Get RndSeed() As Double
On Error Resume Next
   RndSeed = p_RndSeed
End Property

Public Property Get DOTgraph() As GraphvizDOTgraph

On Error GoTo HandleErr

   If p_DOTgraph Is Nothing Then
      Set p_DOTgraph = New GraphvizDOTgraph
   End If
   Set DOTgraph = p_DOTgraph

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "DOTgraph", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Public Property Let GraphMode(lNewValue As DOT_GraphMode)

On Error GoTo HandleErr

   If p_lngGraphMode <> lNewValue Then
      p_lngGraphMode = lNewValue
      ClearDOT
      Set p_objDOT = Nothing
   End If

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "GraphMode", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Public Property Get GraphMode() As DOT_GraphMode
On Error Resume Next
   GraphMode = p_lngGraphMode
End Property

Public Property Let GraphName(sNewValue As String)

On Error GoTo HandleErr

   If p_strGraphName <> sNewValue Then
      p_strGraphName = sNewValue
      Me.ClearDOT
   End If

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "GraphName", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Public Property Get GraphName() As String

On Error GoTo HandleErr

   If Len(p_strGraphName) = 0 Then
      GraphName = conDefaultGraphName
   Else
      GraphName = p_strGraphName
   End If

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "GraphName", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Public Property Let GraphLabel(sNewValue As String)

On Error GoTo HandleErr

   If p_strGraphLabel <> sNewValue Then
      p_strGraphLabel = sNewValue
      Me.ClearDOT
   End If

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "GraphLabel", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Public Property Get GraphLabel() As String
On Error Resume Next
   GraphLabel = p_strGraphLabel
End Property

Public Property Let RankDir(sNewValue As DOT_RankDir)

On Error GoTo HandleErr

   If p_RankDir <> sNewValue Then
      p_RankDir = sNewValue
      Me.ClearDOT
      Set p_objDOT = Nothing
   End If

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "RankDir", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Public Property Get RankDir() As DOT_RankDir
On Error Resume Next
   RankDir = p_RankDir
End Property

Public Property Let InvisibleSubGraph(sNewValue As Boolean)

On Error GoTo HandleErr

   If p_InvisibleSubGraph <> sNewValue Then
      p_InvisibleSubGraph = sNewValue
      Me.ClearDOT
      Set p_objDOT = Nothing
   End If

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "InvisibleSubGraph", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Public Property Get InvisibleSubGraph() As Boolean
On Error Resume Next
   InvisibleSubGraph = p_InvisibleSubGraph
End Property

Public Property Get Node(vIndex As Variant) As GraphvizDOTnode

On Error GoTo HandleErr

   Set Node = p_DOTgraph.Node(vIndex)

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "Node", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Public Property Get Nodes() As Collection

On Error GoTo HandleErr

   Set Nodes = p_DOTgraph.Nodes

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "Nodes", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property


Public Sub AddNode(sID As String, _
               Optional ByVal sLabel As String = vbNullString, _
               Optional ByVal sShape As String = vbNullString, _
               Optional ByVal sColor As String = vbNullString, _
               Optional ByVal sStyle As String = vbNullString, _
               Optional ByVal sMiscOptions As String = vbNullString, _
               Optional ByRef oSubGraph As GraphvizDOTgraph = Nothing, _
               Optional ByVal sLink As String = vbNullString)
               
On Error GoTo HandleErr

   DOTgraph.AddNode sID, sLabel, sShape, sColor, sStyle, sMiscOptions, oSubGraph, , , , sLink

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "AddNode", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Public Sub AddConnection(sFrom As String, sTo As String)

On Error GoTo HandleErr

   p_DOTgraph.AddConnection sFrom, sTo

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "AddConnection", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Public Function ToBinaryGraph(ByVal OutputType As DOT_OutputGraphType, _
                     Optional ByVal show As Boolean = False) As String 'Rückgabe = Dateiname

On Error GoTo HandleErr

   If m_conUseEXE Then
      ToBinaryGraph = ToBinaryGraphExe(OutputType, show)
   Else
      ToBinaryGraph = ToBinaryGraphOcx(OutputType, show)
   End If

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ToBinaryGraph", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function

Private Function ToBinaryGraphOcx(OutputType As DOT_OutputGraphType, _
                       Optional show As Boolean = False) As String 'Rückgabe = Dateiname

   Dim strFile As String
   Dim bolDotValid As Boolean
   Dim Img As Object
   
   Dim strTempDOT As String
   Dim IfileNo As Long
   
On Error GoTo HandleErr
   
'   If ValidateDOT = False Then
'      MsgBox "DOT-String ist fehlerhaft."
'      ToBinaryGraphOcx = vbNullString
'      Exit Function
'   End If
   
   'DOT ermitteln
   strTempDOT = CheckedObjWingraph.ToDot(DOTstring)

   'DOT-File zwischenspeichern
'   strFile = TempPath & Me.GraphName & ".dot"
'   IfileNo = FreeFile
'   Open strFile For Output As #IfileNo
'      Print #IfileNo, strTempDOT
'   Close #IfileNo
   
   'Image-Map zwischenspeichern
   strFile = TempPath & Me.GraphName & ".cmap"
   IfileNo = FreeFile
   Open strFile For Output As #IfileNo
      If p_lngGraphMode = DOT_GraphMode.GraphMode_DOT Then
         Print #IfileNo, Replace(CheckedObjWingraph.ToCMAP(strTempDOT), "\n", " ")
      Else
         Print #IfileNo, Replace(CheckedObjWingraph.ToCMAP(DOTstring), "\n", " ")
      End If
   Close #IfileNo
   
   If p_lngGraphMode = DOT_GraphMode.GraphMode_DOT Then
      Set Img = CheckedObjWingraph.ToBinaryGraph(strTempDOT, OutputType)
   Else
      Set Img = CheckedObjWingraph.ToBinaryGraph(DOTstring, OutputType)
   End If
   strFile = TempPath & Me.GraphName & "." & Img.Type
   Img.Save strFile
   Set Img = Nothing
   
   If show Then
      ShellExecuteOpenFile strFile, "open"
      
   End If

   ToBinaryGraphOcx = strFile
   'ClearDOT

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ToBinaryGraphOcx", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

Private Function ToBinaryGraphExe(OutputType As DOT_OutputGraphType, _
                       Optional show As Boolean = False) As String 'Rückgabe = Dateiname

   Dim strFile As String
   Dim strDotFile As String
   Dim bolDotValid As Boolean
   Dim Img As Object
   
   Dim strTempDOT As String
   Dim IfileNo As Long
   
   Dim strExe As String
   Dim strFileExtension As String
   
On Error GoTo HandleErr

'   If ValidateDOT = False Then
'      MsgBox "DOT-String ist fehlerhaft."
'      ToBinaryGraphEXE = vbNullString
'      Exit Function
'   End If
   
   Select Case OutputType
      Case DOT_OutputGraphType.DOT2EMF
         strFileExtension = ".emf"
      Case DOT_OutputGraphType.DOT2GIF
         strFileExtension = ".gif"
      Case DOT_OutputGraphType.DOT2JPEG
         strFileExtension = ".jpg"
      Case DOT_OutputGraphType.DOT2PNG
         strFileExtension = ".png"
      Case Else
         strFileExtension = ".png"
   End Select
   
   'DOT-File zwischenspeichern
   strDotFile = TempPath & Me.GraphName & ".dot"
   IfileNo = FreeFile
   Open strDotFile For Output As #IfileNo
      Print #IfileNo, DOTstring
   Close #IfileNo
   
   'DOT ermitteln
   'If p_lngGraphMode = DOT_GraphMode.DOT Then
      'strExe = "-Gcharset=latin1 -Tdot -o " & strDotFile & " " & strDotFile
      strExe = "-Gcharset=latin1 -Tdot -o """ & strDotFile & """ """ & strDotFile & """"
      runDotExe strExe
   'End If
   
   'Image-Map zwischenspeichern
   strFile = TempPath & Me.GraphName & ".cmap"
   'strExe = "-Gcharset=latin1 -Tcmap -o " & strFile & " " & strDotFile
   strExe = "-Tcmap -o " & strFile & " " & strDotFile
   
   runDotExe strExe
   
   strFile = TempPath & Me.GraphName & strFileExtension
   strExe = "-Gcharset=latin1 -T" & Mid(strFileExtension, 2) & " -o """ & strFile & """ """ & strDotFile & """"
   'strExe = "-T" & Mid(strFileExtension, 2) & " -o " & strFile & " " & strDotFile
   runDotExe strExe

   If show Then
      ShellExecuteOpenFile strFile, "open"
   End If

   ToBinaryGraphExe = strFile
   'ClearDOT

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ToBinaryGraphExe", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

Private Function runDotExe(sDotCmd As String)

   Static strDotPath As String
   Dim bolReadPath As Boolean
   Dim strFile As String
   Dim strDotExe As String

On Error GoTo HandleErr

   If Len(strDotPath) = 0 Then
      bolReadPath = True
   ElseIf Len(Dir$(strDotPath & "dot.exe")) = 0 Then
      bolReadPath = True
   End If
   
   If bolReadPath Then

      If modWinAPI_Registry.RegValueGet(HKEY_LOCAL_MACHINE, "SOFTWARE\AT&T Research Labs\Graphviz", "InstallPath", strDotPath) = 0 Then
         strDotPath = Trim$(strDotPath)
         If Right$(strDotPath, 1) <> "\" Then
            strDotPath = strDotPath & "\"
         End If
         strDotPath = strDotPath & "bin\"
         If Len(Dir$(strDotPath & "dot.exe")) = 0 Then
            strDotPath = CodeProject.Path & "\bin\graphviz\bin\"
         End If
      Else
         strDotPath = CodeProject.Path & "\bin\graphviz\bin\"
      End If
      
      If Len(Dir(strDotPath & "dot.exe")) = 0 Then
         strDotPath = CodeProject.Path & "\bin\graphviz\"
         If Len(Dir$(strDotPath & "dot.exe")) = 0 Then
            Err.Raise vbObjectError, "RunDotExe", "Graphviz-Execute-Dateien wurde nicht gefunden"
         End If
      End If
   
   End If
   
   Select Case p_lngGraphMode
      Case DOT_GraphMode.GraphMode_NEATO
         strDotExe = strDotPath & "neato.exe"
      Case DOT_GraphMode.GraphMode_TWOPI
         strDotExe = strDotPath & "twopi.exe"
      Case Else 'DOT_GraphMode.DOT
         strDotExe = strDotPath & "dot.exe"
   End Select
   
   modWinAPI.LaunchAppSynchronous strDotExe, sDotCmd, vbHide


ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "runDotExe", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function


Public Function ToTextGraph(lType As Long, _
                     Optional show As Boolean = False) As String 'Rückgabe = Dateiname

   Dim strFile As String, strFileExt As String
   Dim bolDotValid As Boolean
   Dim strGraphText As String

On Error GoTo HandleErr

   If ValidateDOT = False Then
      MsgBox "DOT-String ist fehlerhaft."
      ToTextGraph = vbNullString
      Exit Function
   End If
   
   strGraphText = CheckedObjWingraph.ToTextGraph(DOTstring, lType)
   Select Case lType
      Case 0  'GRAPHVIZ_ATTRIBUTED_DOT = 0
         strFileExt = "dot"
      Case 1  'GRAPHVIZ_PS = 1
         strFileExt = "eps"
      Case 19 'GRAPHVIZ_VRML = 19
         strFileExt = "vrml"
      Case 23 'GRAPHVIZ_SVG = 23
         strFileExt = "svg"
      Case Else 'GRAPHVIZ_PLAIN = 6
         strFileExt = "txt"
   End Select
   
   strFile = TempPath & Me.GraphName & "." & strFileExt
   Open strFile For Output As #1
      Print #1, strGraphText
   Close #1

   If show Then
      modWinAPI.ShellExecuteOpenFile strFile, "open"
   End If

   ToTextGraph = strFile


ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ToTextGraph", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

Public Function ToPS(Optional ByVal show As Boolean = False) As String 'Rückgabe = Dateiname

   Dim strFile As String, strFileExt As String
   Dim bolDotValid As Boolean
   Dim strGraphText As String

On Error GoTo HandleErr

   If ValidateDOT = False Then
      MsgBox "DOT-String ist fehlerhaft."
      ToPS = vbNullString
      Exit Function
   End If
   
   strGraphText = CheckedObjWingraph.ToPS(DOTstring)
   strFileExt = "eps"
   strFile = TempPath & Me.GraphName & "." & strFileExt
   Open strFile For Output As #1
      Print #1, strGraphText
   Close #1

   If show Then
      modWinAPI.ShellExecuteOpenFile strFile, "open"
   End If

   ToPS = strFile

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ToPS", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function


Public Function ToPNG(Optional ByVal show As Boolean = False) As String
'Generate BinaryImage(PNG)

On Error GoTo HandleErr

   ToPNG = Me.ToBinaryGraph(DOT2PNG, show)

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ToPNG", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

Public Function ToJPEG(Optional ByVal show As Boolean = False) As String
'Generate BinaryImage(PNG)

On Error GoTo HandleErr

   ToJPEG = Me.ToBinaryGraph(DOT2JPEG, show)

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ToJPEG", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

Public Function ToGIF(Optional ByVal show As Boolean = False) As String
'Generate BinaryImage(PNG)

On Error GoTo HandleErr

   ToGIF = Me.ToBinaryGraph(DOT2GIF, show)

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ToGIF", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

Public Function ToSvg() As String  'Text-Output

   Dim strFile As String, strFileExt As String
   Dim bolDotValid As Boolean
   Dim strGraphText As String

On Error GoTo HandleErr

   If ValidateDOT = False Then
      MsgBox "DOT-String ist fehlerhaft."
      ToSvg = vbNullString
      Exit Function
   End If
   
   strGraphText = CheckedObjWingraph.ToSvg(DOTstring)
   strFileExt = "svg"
   strFile = TempPath & Me.GraphName & "." & strFileExt
   Open strFile For Output As #1
      Print #1, strGraphText
   Close #1
'
'   If show Then
'      WinAPI.ShellExecuteOpenFile strFile, "open"
'   End If

   ToSvg = strFile

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ToSvg", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function

Private Sub CreateDOT()

   Dim strDOT As String
   Dim strRankDir As String

'   If p_RankDir = 0 Then
'      strRankDir = "LR"
'   Else
'      strRankDir = "TB"
'   End If
On Error GoTo HandleErr

   Select Case p_RankDir
      Case DOT_RankDir.RankDir_LR
         strRankDir = "LR"
      Case DOT_RankDir.RankDir_TB
         strRankDir = "TB"
      Case DOT_RankDir.RankDir_RL
         strRankDir = "RL"
      Case DOT_RankDir.RankDir_BT
         strRankDir = "BT"
   End Select

   
   strDOT = "digraph G {" & vbNewLine '& _
          "graph [rankdir=" & strRankDir & "];" & vbNewLine
   'nodesep=0.25;rank=same;
   'nslimit=1;
   ' => p_strGraphProp aus clsFuDOT

   If p_lngGraphMode = DOT_GraphMode.GraphMode_DOT Then
      strDOT = strDOT & "graph [rankdir=" & strRankDir & "];" & vbNewLine
   End If
   
   If p_RndSeed > 0 Then
      strDOT = strDOT & "graph [start=" & Str(p_RndSeed) & "];" & vbNewLine
   End If

   If Len(p_strGraphProp) > 0 Then
      strDOT = strDOT & p_strGraphProp
   End If
   If Len(p_strGraphLabel) > 0 Then
      strDOT = strDOT & "label=""" & p_strGraphLabel & """;"
   End If
   If Len(p_strFirstLevelNodeProp) > 0 Then
      strDOT = strDOT & "node [" & p_strFirstLevelNodeProp & "];" & vbNewLine
   End If
   strDOT = strDOT & p_DOTgraph.DOTstring(p_InvisibleSubGraph) & vbNewLine & "}" & vbNewLine
   p_strDOT = strDOT

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "CreateDOT", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Function ValidateDOT() As Boolean

On Error GoTo HandleErr

   If m_conUseEXE Then
      ValidateDOT = True
   Else
      ValidateDOT = CheckedObjWingraph.Validate(DOTstring)
   End If

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ValidateDOT", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function
