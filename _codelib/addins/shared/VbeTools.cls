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
' VBE-Programmierunterstützung
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/shared/VbeTools.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

#Const EARLYBINDING = 0
#Const DEBUGMODE = 0

#If EARLYBINDING Then
Friend Property Get CurrentVbProject() As VBIDE.VBProject
#Else
Friend Property Get CurrentVbProject() As Object
#End If

#If EARLYBINDING Then
   Dim Proj As VBProject
   Dim Temp As VBProject
#Else
   Dim Proj As Object
   Dim Temp As Object
#End If
   Dim strCurrentDbName As String

   Set Proj = VBE.ActiveVBProject
   'Prüfen, ob das richtige VbProject gewählt wurde (muss das von CurrentDb sein)
   strCurrentDbName = UncPath(CurrentDb.Name)
   If Proj.FileName <> strCurrentDbName Then
      Set Proj = Nothing
      For Each Temp In VBE.VBProjects
         If Temp.FileName = strCurrentDbName Then
            Set Proj = Temp
            Exit For
         End If
      Next
   End If
   
   Set CurrentVbProject = Proj

End Property

Public Sub ActivateCurrentProject()

   Dim Activated As Boolean
#If EARLYBINDING Then
   Dim vbc As VBComponent
#Else
   Dim vbc As Object
#End If

   If CurrentVbProject.VBComponents.Count > 0 Then
      'zuerst nur Code-Module verwenden, da sonst Formulare/Berichte im Entwurf geöffnet werden
      For Each vbc In CurrentVbProject.VBComponents
         If vbc.Type <> vbext_ct_Document Then
            vbc.Activate
            Activated = True
            Exit For
         End If
      Next
      'falls nur vbext_ct_Document vorhanden sind, das erste verwenden
      If Not Activated Then
         CurrentVbProject.VBComponents(1).Activate
      End If
   End If

End Sub

Public Function CodeModuleExists(ByVal CodeModulName As String) As Boolean

#If EARLYBINDING = 1 Then
   Dim VbcCol As VBIDE.VBComponents
   Dim vbc As VBIDE.VBComponent
#Else
   Dim VbcCol As Object
   Dim vbc As Object
#End If
   
   Set VbcCol = CurrentVbProject.VBComponents
   For Each vbc In VbcCol
      If vbc.Name = CodeModulName Then
         CodeModuleExists = True
         Exit For
      End If
   Next
   
End Function
