VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeModuleHeaderItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: CodeModuleHeaderItem (Josef Pötzl, 2010-02-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Konstanten, Variablen, Types und Enums
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/shared/CodeModuleHeaderItem.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/shared/CodeModuleReader.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Public Enum CodeModuleHeaderItemType
   CodeModuleHeaderItemType_Var = 0
   CodeModuleHeaderItemType_Const = 1
   CodeModuleHeaderItemType_Enum = 2
   CodeModuleHeaderItemType_Type = 3
   CodeModuleHeaderItemType_Declare = 4
End Enum

Public Name As String
Public ItemType As CodeModuleHeaderItemType

#If EARLYBINDING Then
Public HeaderItemVBComponent As VBIDE.VBComponent
#Else
Public HeaderItemVBComponent As Object
#End If

Public Property Get SearchStrings() As String()

#If USEREGEX = 1 Then
   Dim searchStringArray(0) As String
#Else
   Dim searchStringArray(11) As String
#End If
   Dim modulName As String
   Dim i As Long
   Dim maxCnt As Long
   
   modulName = HeaderItemVBComponent.Name
   If HeaderItemVBComponent.Type = vbext_ct_Document Then
      'eigentlich nicht notwendig, da Klassen nicht nach Procs durchsucht werden müssen
      If modulName Like "Form_" Then
         modulName = Mid(modulName, 6)
      ElseIf modulName Like "Report_" Then
         modulName = Mid(modulName, 8)
      End If
   End If
   
#If USEREGEX = 1 Then
   searchStringArray(0) = "(\n| )(" & modulName & "\.|)" & Me.Name & "( |\.|,|:|\(|\n)"
#Else

   searchStringArray(0) = " " & Me.Name & " "
   searchStringArray(1) = " " & Me.Name & "[.,:(]"
   searchStringArray(2) = " " & Me.Name & vbNewLine
   maxCnt = 2

   For i = 0 To maxCnt
      searchStringArray(i + maxCnt + 1) = " " & modulName & "." & Mid(searchStringArray(i), 2)
   Next
   maxCnt = 2 * maxCnt + 1
   

   For i = 0 To maxCnt
      searchStringArray(i + maxCnt + 1) = vbNewLine & Mid(searchStringArray(i), 2)
   Next
   
#End If
   
   SearchStrings = searchStringArray

End Property
