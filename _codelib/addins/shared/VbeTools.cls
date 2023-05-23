VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbeTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Attribute VB_PredeclaredId = True
'---------------------------------------------------------------------------------------
' Class Module: VbeTools
'---------------------------------------------------------------------------------------
'/**
' <summary>
' VBE Programming Support
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/shared/VbeTools.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>file/FileTools.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

#Const EARLYBINDING = 0
#Const DEBUGMODE = 0

Private m_AccessApplication As Access.Application

Friend Property Get AccessApplication() As Access.Application
   If m_AccessApplication Is Nothing Then
      Set AccessApplication = Access.Application
   Else
      Set AccessApplication = m_AccessApplication
   End If
End Property

Friend Property Set AccessApplication(ByVal NewRef As Access.Application)
   Set m_AccessApplication = NewRef
End Property

#If EARLYBINDING Then
Friend Property Get CurrentVbProject() As VBIDE.VBProject
#Else
Friend Property Get CurrentVbProject() As Object
#End If

   Set CurrentVbProject = FindVbProject(AccessApplication.CurrentDb.Name)

End Property

#If EARLYBINDING Then
Friend Property Get FindVbProject(ByVal FilePath As String) As VBIDE.VBProject
#Else
Friend Property Get FindVbProject(ByVal FilePath As String) As Object
#End If

#If EARLYBINDING Then
   Dim Proj As VBProject
   Dim Temp As VBProject
#Else
   Dim Proj As Object
   Dim Temp As Object
#End If

   Set Proj = AccessApplication.VBE.ActiveVBProject
   'Check if the correct VbProject has been selected (must be the one from database file)
   FilePath = UncPath(FilePath)
   If Proj.FileName <> FilePath Then
      Set Proj = Nothing
      For Each Temp In VBE.VBProjects
         If Temp.FileName = FilePath Then
            Set Proj = Temp
            Exit For
         End If
      Next
   End If
   
   Set FindVbProject = Proj

End Property

Public Sub ActivateCurrentProject()
   ActivateVbProject CurrentVbProject
End Sub

#If EARLYBINDING = 1 Then
Public Sub ActivateVbProject(ByVal VbProjectToActivate As VBIDE.VBProject)
#Else
Public Sub ActivateVbProject(ByVal VbProjectToActivate As Object)
#End If

   Dim Activated As Boolean
#If EARLYBINDING Then
   Dim vbc As VBComponent
#Else
   Dim vbc As Object
#End If

   If VbProjectToActivate.VBComponents.Count > 0 Then
      'Only use code modules at first, otherwise forms/reports will be opened in draft form
      For Each vbc In VbProjectToActivate.VBComponents
         If vbc.Type <> vbext_ct_Document Then
            vbc.Activate
            Activated = True
            Exit For
         End If
      Next
      'if there are only vbext_ct_Document, use the first one
      If Not Activated Then
         VbProjectToActivate.VBComponents(1).Activate
      End If
   End If

End Sub

#If EARLYBINDING = 1 Then
Public Function CodeModuleExists(ByVal CodeModulName As String, Optional VbProjectToScan As VBIDE.VBProject = Nothing) As Boolean
#Else
Public Function CodeModuleExists(ByVal CodeModulName As String, Optional VbProjectToScan As Object = Nothing) As Boolean
#End If

#If EARLYBINDING = 1 Then
   Dim VbcCol As VBIDE.VBComponents
   Dim vbc As VBIDE.VBComponent
#Else
   Dim VbcCol As Object
   Dim vbc As Object
#End If

   If VbProjectToScan Is Nothing Then
      Set VbProjectToScan = CurrentVbProject
   End If
   
   Set VbcCol = VbProjectToScan.VBComponents
   For Each vbc In VbcCol
      If vbc.Name = CodeModulName Then
         CodeModuleExists = True
         Exit For
      End If
   Next
   
End Function
