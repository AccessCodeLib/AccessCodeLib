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
' Struktur/Abh�ngigkeiten der Module bestimmen
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
' Ausgangsmodul f�r Struktur/Abh�ngigkeitsermittlung
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
' Ausgangsmodul f�r Struktur/Abh�ngigkeitsermittlung
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
      Set m_CodeModuleReaderCollection = New Collection
   End If
End Property

'---------------------------------------------------------------------------------------
' Sub: ReadStruct (Josef P�tzl, 2010-03-23)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Struktur/Abh�ngigkeiten ermitteln
' </summary>
' <returns></returns>
' <remarks>
' Ergebnis ist �ber CodeModuleReaderCollection auswertbar
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
   Dim TempModuleReader As CodeModuleReader
   Dim BaseModuleReader As CodeModuleReader
   
   Dim strBaseCodeModuleName As String
   Dim lngCount As Long
   
   strBaseCodeModuleName = BaseCodeModuleName
   
   Set vbp = CurrentVbProject

   'alle Module sammeln
   Set m_CodeModuleReaderCollection = New Collection
   For Each vbc In vbp.VBComponents
      Set TempModuleReader = New CodeModuleReader
      Set TempModuleReader.CodeModule = vbc.CodeModule
      m_CodeModuleReaderCollection.Add TempModuleReader, TempModuleReader.Name
      If TempModuleReader.Name = strBaseCodeModuleName Then Set BaseModuleReader = TempModuleReader
   Next
   
   'Abh�ngigkeiten ermitteln
   If Len(strBaseCodeModuleName) = 0 Then 'alle Abh�ngigkeiten
      For Each TempModuleReader In m_CodeModuleReaderCollection
         Call TempModuleReader.CheckDependency(, m_CodeModuleReaderCollection)
      Next
   Else
   
      If BaseModuleReader Is Nothing Then
         Err.Raise vbObjectError, "ReadStruct", "Das Code-Modul '" & strBaseCodeModuleName & "' wurde nicht gefunden!"
      End If
   
      'm_CodeModuleReaderCollection auf Basis des Ausgangsmoduls bereinigen
      Dim newCol As Collection
      Dim checkReader As CodeModuleReader
      
      Set newCol = New Collection
      Call BaseModuleReader.CheckDependency(, m_CodeModuleReaderCollection)
      newCol.Add BaseModuleReader, BaseModuleReader.Name
      
      lngCount = 1
      Do While lngCount <= newCol.Count
         Set TempModuleReader = newCol(lngCount)
         lngCount = lngCount + 1
         For Each checkReader In TempModuleReader.RequiredModules
            If Not ModulReaderExistsInCollection(newCol, checkReader) Then
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

Private Function ModulReaderExistsInCollection(col As Collection, cmr As CodeModuleReader) As Boolean
   Dim TempModuleReader As CodeModuleReader
   For Each TempModuleReader In col
      If TempModuleReader.Name = cmr.Name Then
         ModulReaderExistsInCollection = True
         Exit For
      End If
   Next
End Function

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


'---------------------------------------------------------------------------------------
' Property: CurrentVbProject (Josef P�tzl, 2010-03-23)
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
' Property: CurrentVbProject (Josef P�tzl, 2010-03-23)
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
   Dim Proj As VBProject
#Else
   Dim Proj As Object
#End If
   Dim strCurrentDbName As String

   If m_CurrentVbProject Is Nothing Then
      Set m_CurrentVbProject = VBE.ActiveVBProject
      'Pr�fen, ob das richtige VbProject gew�hlt wurde (muss das von CurrentDb sein)
      strCurrentDbName = UncPath(CurrentDb.Name)
      If m_CurrentVbProject.FileName <> strCurrentDbName Then
         Set m_CurrentVbProject = Nothing
         For Each Proj In VBE.VBProjects
            If Proj.FileName = strCurrentDbName Then
               Set m_CurrentVbProject = Proj
               Exit For
            End If
         Next
      End If
   End If
   
   Set CurrentVbProject = m_CurrentVbProject

End Property
