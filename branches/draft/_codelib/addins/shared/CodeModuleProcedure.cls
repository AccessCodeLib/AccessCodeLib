VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeModuleProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: CodeModuleProcedure (Josef Pötzl)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' CodeModuleProcedure
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/shared/CodeModuleProcedure.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/shared/CodeModuleReader.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Public Name As String

#If EARLYBINDING Then
   Public ProcVbComponent As VBIDE.VBComponent
   Public ProcKind As vbext_ProcKind
#Else
   Public ProcVbComponent As Object
   Public ProcKind As Long
#End If


Public Function SearchStrings(Optional ByVal OnlyWithModulePrefix As Boolean = False) As String()

   Dim SearchStringArray() As String
   Dim ModulName As String
   Dim i As Long
   Dim MaxCnt As Long
   
   ModulName = ProcVbComponent.Name
   If ProcVbComponent.Type = vbext_ct_Document Then
      'eigentlich nicht notwendig, da Klassen nicht nach Procs durchsucht werden müssen
      If ModulName Like "Form_" Then
         ModulName = Mid(ModulName, 6)
      ElseIf ModulName Like "Report_" Then
         ModulName = Mid(ModulName, 8)
      End If
   End If
   
#If USEREGEX = 1 Then
   ReDim SearchStringArray(0)
   If OnlyWithModulePrefix Then
      SearchStringArray(0) = "(\n| )" & ModulName & "\." & Me.Name & "( |\.|,|:|\(|\n)"
   Else
      SearchStringArray(0) = "(\n| )(" & ModulName & "\.|)" & Me.Name & "( |\.|,|:|\(|\n)"
   End If
#Else

   If OnlyWithModulePrefix Then
      ReDim SearchStringArray(5)
   Else
      ReDim SearchStringArray(11)
   End If

   SearchStringArray(0) = " " & Me.Name & " "
   SearchStringArray(1) = " " & Me.Name & "[.,:(]" 'für like-Suche ist das so in Ordnung
   SearchStringArray(2) = " " & Me.Name & vbNewLine
   MaxCnt = 2

   For i = 0 To MaxCnt
      SearchStringArray(i + (MaxCnt + 1) * (1 + OnlyWithModulePrefix)) = " " & ModulName & "." & Mid(SearchStringArray(i), 2)
   Next
   MaxCnt = MaxCnt + (MaxCnt + 1) * (1 + OnlyWithModulePrefix)

   For i = 0 To MaxCnt
      SearchStringArray(i + MaxCnt + 1) = vbNewLine & Mid(SearchStringArray(i), 2)
   Next

#End If

   SearchStrings = SearchStringArray

End Function
