VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeModuleStructReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: CodeModuleStructReader
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Struktur/Abhängigkeiten der Module bestimmen
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/shared/CodeModuleStructReader.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/shared/CodeModuleReader.cls</use>
'  <use>file/modFiles.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_CodeModuleReaderCollection As Collection

Private m_BaseCodeModuleName As String
Private m_RefreshStruct As Boolean

#If EARLYBINDING Then
Private m_CurrentVbProject As VBIDE.VBProject
#Else
Private m_CurrentVbProject As Object
#End If

Private Sub Class_Initialize()
   Set m_CodeModuleReaderCollection = New Collection
End Sub

'---------------------------------------------------------------------------------------
' Property: CodeModuleReaderCollection
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Collection von CodeModuleReader
' </summary>
' <returns>Collection</returns>
' <remarks>
' je Code-Modul gibt es einen CodeModuleReader
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get CodeModuleReaderCollection() As Collection
   Set CodeModuleReaderCollection = m_CodeModuleReaderCollection
End Property

'---------------------------------------------------------------------------------------
' Property: BaseCodeModuleName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ausgangsmodul für Struktur/Abhängigkeitsermittlung
' </summary>
' <param name="Param"></param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get BaseCodeModuleName() As String
   BaseCodeModuleName = m_BaseCodeModuleName
End Property

'---------------------------------------------------------------------------------------
' Property: BaseCodeModuleName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ausgangsmodul für Struktur/Abhängigkeitsermittlung
' </summary>
' <param name="Param"></param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let BaseCodeModuleName(ByVal CmName As String)
   If CmName <> m_BaseCodeModuleName Then
      m_RefreshStruct = True
      m_BaseCodeModuleName = CmName
   End If
End Property

'---------------------------------------------------------------------------------------
' Sub: ReadStruct (Josef Pötzl, 2010-03-23)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Struktur/Abhängigkeiten ermitteln
' </summary>
' <returns></returns>
' <remarks>
' Ergebnis ist über CodeModuleReaderCollection auswertbar
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub ReadStruct()

#If EARLYBINDING Then
   Dim vbp As VBIDE.VBProject
   Dim vbc As VBIDE.VBComponent
#Else
   Dim vbp As Object
   Dim vbc As Object
#End If
   Dim tempModuleReader As CodeModuleReader
   Dim baseModuleReader As CodeModuleReader
   
   Dim strBaseCodeModuleName As String
   Dim lngCount As Long
   
   strBaseCodeModuleName = BaseCodeModuleName
   
   Set vbp = CurrentVbProject

   'alle Module sammeln
   Set m_CodeModuleReaderCollection = New Collection
   For Each vbc In vbp.VBComponents
      Set tempModuleReader = New CodeModuleReader
      Set tempModuleReader.CodeModule = vbc.CodeModule
      m_CodeModuleReaderCollection.Add tempModuleReader, tempModuleReader.Name
      If tempModuleReader.Name = strBaseCodeModuleName Then Set baseModuleReader = tempModuleReader
   Next
   
   'Abhängigkeiten ermitteln
   If Len(strBaseCodeModuleName) = 0 Then 'alle Abhängigkeiten
      For Each tempModuleReader In m_CodeModuleReaderCollection
         Call tempModuleReader.CheckDependency(, m_CodeModuleReaderCollection)
      Next
   Else
   
      If baseModuleReader Is Nothing Then
         Err.Raise vbObjectError, "ReadStruct", "Das Code-Modul '" & strBaseCodeModuleName & "' wurde nicht gefunden!"
      End If
   
      'm_CodeModuleReaderCollection auf Basis des Ausgangsmoduls bereinigen
      Dim newCol As Collection
      Dim checkReader As CodeModuleReader
      
      Set newCol = New Collection
      Call baseModuleReader.CheckDependency(, m_CodeModuleReaderCollection)
      newCol.Add baseModuleReader, baseModuleReader.Name
      
      lngCount = 1
      Do While lngCount <= newCol.Count
         Set tempModuleReader = newCol(lngCount)
         lngCount = lngCount + 1
         For Each checkReader In tempModuleReader.RequiredModules
            If Not modulReaderExistsInCollection(newCol, checkReader) Then
               Call checkReader.CheckDependency(, m_CodeModuleReaderCollection)
               newCol.Add checkReader, checkReader.Name
            End If
         Next
      Loop
   
      Set m_CodeModuleReaderCollection = Nothing
      Set m_CodeModuleReaderCollection = newCol
   
   End If

   m_RefreshStruct = False
   
End Sub

Private Function modulReaderExistsInCollection(col As Collection, cmr As CodeModuleReader) As Boolean
   Dim tempModuleReader As CodeModuleReader
   For Each tempModuleReader In col
      If tempModuleReader.Name = cmr.Name Then
         modulReaderExistsInCollection = True
         Exit For
      End If
   Next
End Function


Private Function getLink(cmr As CodeModuleReader) As String
   
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
   
   getLink = LCase(strName) & ".html"
   
End Function


'---------------------------------------------------------------------------------------
' Property: CurrentVbProject (Josef Pötzl, 2010-03-23)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Zu verwendendes VB-Projekt
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
#If EARLYBINDING Then
Public Property Set CurrentVbProject(ref As VBIDE.VBProject)
#Else
Public Property Set CurrentVbProject(ref As Object)
#End If

   Set m_CurrentVbProject = ref

End Property

'---------------------------------------------------------------------------------------
' Property: CurrentVbProject (Josef Pötzl, 2010-03-23)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Zu verwendendes VB-Projekt
' </summary>
' <returns>VBIDE.VBProject</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
#If EARLYBINDING Then
Public Property Get CurrentVbProject() As VBIDE.VBProject
#Else
Public Property Get CurrentVbProject() As Object
#End If

#If EARLYBINDING Then
   Dim proj As VBProject
#Else
   Dim proj As Object
#End If
   Dim strCurrentDbName As String

On Error GoTo HandleErr

   If m_CurrentVbProject Is Nothing Then
      Set m_CurrentVbProject = VBE.ActiveVBProject
      'Prüfen, ob das richtige VbProject gewählt wurde (muss das von CurrentDb sein)
      strCurrentDbName = UNCPath(CurrentDb.Name)
      If m_CurrentVbProject.FileName <> strCurrentDbName Then
         Set m_CurrentVbProject = Nothing
         For Each proj In VBE.VBProjects
            If proj.FileName = strCurrentDbName Then
               Set m_CurrentVbProject = proj
               Exit For
            End If
         Next
      End If
   End If
   
   Set CurrentVbProject = m_CurrentVbProject

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "CurrentVbProject", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Property
