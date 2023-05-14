Attribute VB_Name = "L10nTools"
'---------------------------------------------------------------------------------------
' Modul: L10nTools
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Pötzl
' <summary>
' Localization (L10n) Functions
' </summary>
' <remarks>
' Use compiler constant L10nMsgBoxReplacement to overwrite MsgBox and InPutBox functions.
' </remarks>
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

#If L10nMsgBoxReplacement = 1 Then

Public Function MsgBox(ByVal Prompt As Variant, _
              Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
              Optional ByVal Title As Variant, _
              Optional ByVal HelpFile As Variant, _
              Optional ByVal Context As Variant) As VbMsgBoxResult

   MsgBox = L10n.MsgBox(Prompt, Buttons, Title, HelpFile, Context)

End Function

Public Function InputBox(ByVal Prompt As Variant, _
              Optional ByVal Title As Variant, _
              Optional ByVal Default As Variant, _
              Optional ByVal XPos As Variant, Optional ByVal YPos As Variant, _
              Optional ByVal HelpFile As Variant, _
              Optional ByVal Context As Variant) As String

   InputBox = L10n.InputBox(Prompt, Title, Default, XPos, YPos, HelpFile, Context)

End Function

#End If
