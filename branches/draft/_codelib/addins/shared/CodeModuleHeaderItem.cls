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
   cmItemType_Var = 0
   cmItemType_Const = 1
   cmItemType_Enum = 2
   cmItemType_Type = 3
   cmItemType_Declare = 4
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
   Dim SearchStringArray(0) As String
#Else
   Dim SearchStringArray(11) As String
#End If
   Dim ModulName As String
   Dim i As Long
   Dim MaxCnt As Long
   
   ModulName = HeaderItemVBComponent.Name
   If HeaderItemVBComponent.Type = vbext_ct_Document Then
      'eigentlich nicht notwendig, da Klassen nicht nach Procs durchsucht werden müssen
      If ModulName Like "Form_" Then
         ModulName = Mid(ModulName, 6)
      ElseIf ModulName Like "Report_" Then
         ModulName = Mid(ModulName, 8)
      End If
   End If
   
#If USEREGEX = 1 Then
   SearchStringArray(0) = "(\n| )(" & ModulName & "\.|)" & Me.Name & "( |\.|,|:|\(|\n)"
#Else

   SearchStringArray(0) = " " & Me.Name & " "
   SearchStringArray(1) = " " & Me.Name & "[.,:(]"
   SearchStringArray(2) = " " & Me.Name & vbNewLine
   MaxCnt = 2

   For i = 0 To MaxCnt
      SearchStringArray(i + MaxCnt + 1) = " " & ModulName & "." & Mid(SearchStringArray(i), 2)
   Next
   MaxCnt = 2 * MaxCnt + 1
   

   For i = 0 To MaxCnt
      SearchStringArray(i + MaxCnt + 1) = vbNewLine & Mid(SearchStringArray(i), 2)
   Next
   
#End If
   
   SearchStrings = SearchStringArray

End Property
