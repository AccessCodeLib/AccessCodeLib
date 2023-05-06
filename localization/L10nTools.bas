Attribute VB_Name = "L10nTools"
'---------------------------------------------------------------------------------------
' Modul: L10nTools
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Pötzl
' <summary>
' Localization (L10n) Functions
' </summary>
' <remarks></remarks>
'
' \ingroup localization
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>localization/L10nTools.bas</file>
'  <use>localization/L10nDict.cls</use>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

Public Property Get L10n() As L10nDict
   Set L10n = L10nDict
End Property

Public Function L10nMsgBox(ByVal Prompt As Variant, _
              Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
              Optional ByVal Title As Variant, _
              Optional ByVal HelpFile As Variant, _
              Optional ByVal Context As Variant) As VbMsgBoxResult
   
   If Not IsMissing(Prompt) Then
      Prompt = L10n.Text(Prompt)
   End If
   
   If Not IsMissing(Title) Then
      Title = L10n.Text(Title)
   End If
   
   L10nMsgBox = VBA.MsgBox(Prompt, Buttons, Title, HelpFile, Context)
   
End Function
