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
   Public ProcVBComponent As VBIDE.VBComponent
   Public procKind As vbext_ProcKind
#Else
   Public ProcVBComponent As Object
   Public procKind As Long
#End If


Public Function SearchStrings(Optional ByVal onlyWithModulePrefix As Boolean = False) As String()

   Dim searchStringArray() As String
   Dim modulName As String
   Dim i As Long
   Dim maxCnt As Long
   
   modulName = ProcVBComponent.Name
   If ProcVBComponent.Type = vbext_ct_Document Then
      'eigentlich nicht notwendig, da Klassen nicht nach Procs durchsucht werden müssen
      If modulName Like "Form_" Then
         modulName = Mid(modulName, 6)
      ElseIf modulName Like "Report_" Then
         modulName = Mid(modulName, 8)
      End If
   End If
   
#If USEREGEX = 1 Then
   ReDim searchStringArray(0)
   If onlyWithModulePrefix Then
      searchStringArray(0) = "(\n| )" & modulName & "\." & Me.Name & "( |\.|,|:|\(|\n)"
   Else
      searchStringArray(0) = "(\n| )(" & modulName & "\.|)" & Me.Name & "( |\.|,|:|\(|\n)"
   End If
#Else

   If onlyWithModulePrefix Then
      ReDim searchStringArray(5)
   Else
      ReDim searchStringArray(11)
   End If

   searchStringArray(0) = " " & Me.Name & " "
   searchStringArray(1) = " " & Me.Name & "[.,:(]" 'für like-Suche ist das so in Ordnung
   searchStringArray(2) = " " & Me.Name & vbNewLine
   maxCnt = 2

   For i = 0 To maxCnt
      searchStringArray(i + (maxCnt + 1) * (1 + onlyWithModulePrefix)) = " " & modulName & "." & Mid(searchStringArray(i), 2)
   Next
   maxCnt = maxCnt + (maxCnt + 1) * (1 + onlyWithModulePrefix)

   For i = 0 To maxCnt
      searchStringArray(i + maxCnt + 1) = vbNewLine & Mid(searchStringArray(i), 2)
   Next

#End If

   SearchStrings = searchStringArray

End Function
