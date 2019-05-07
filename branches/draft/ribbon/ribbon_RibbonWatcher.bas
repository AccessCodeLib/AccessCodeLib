Attribute VB_Name = "ribbon_RibbonWatcher"
'---------------------------------------------------------------------------------------
' Module: mod_RibbonWatcher_CallBacks (Josef Pötzl, 2010-04-10)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Globale Prozeduren für RibbonWatcher (inkl. CallBack-Prozeduren)
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>ribbon/ribbon_RibbonWatcher.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>ribbon/RibbonWatcher.cls</use>
'  <ref><name>Office</name><major>2</major><minor>4</minor><guid>{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private m_CurrentRibbonWatcher As RibbonWatcher
Public AppLoadedForRibbonWatcher As Boolean

'RibbonWatcher-Instanz
Public Property Get CurrentRibbonWatcher() As RibbonWatcher
On Error Resume Next
   If m_CurrentRibbonWatcher Is Nothing Then
      Set m_CurrentRibbonWatcher = New RibbonWatcher
   End If
   Set CurrentRibbonWatcher = m_CurrentRibbonWatcher
End Property


'####################################################################################################
'
' Callback-Prozeduren
'

Public Sub RibbonWatcherCallBack_OnLoad(ByRef ribbon As IRibbonUI)
On Error Resume Next
   Set CurrentRibbonWatcher.RibbonUI = ribbon

End Sub

Public Sub RibbonWatcherCallBack_OnAction(ByRef rc As IRibbonControl)
On Error Resume Next
   CurrentRibbonWatcher.CallRibbonControlOnAction rc

End Sub

Public Sub RibbonWatcherCallBack_GetLabel(ByRef rc As IRibbonControl, _
                                          ByRef Label As Variant)
On Error Resume Next
   Label = CurrentRibbonWatcher.GetRibbonControlLabel(rc)
End Sub

Public Sub RibbonWatcherCallBack_GetImages(ByRef rc As IRibbonControl, _
                                           ByRef Image As Variant)
On Error Resume Next
   Dim strImage As String
   strImage = CurrentRibbonWatcher.GetRibbonControlImage(rc)
   If Len(strImage) > 0 Then
      Image = strImage
   Else
      Image = "CreateReport"
   End If
   
End Sub

Public Sub RibbonWatcherCallBack_GetVisible(ByRef rc As IRibbonControl, _
                                            ByRef Visible As Variant)
On Error Resume Next
   Visible = CurrentRibbonWatcher.GetRibbonControlVisible(rc)

End Sub

Public Sub RibbonWatcherCallBack_GetGroupVisible(ByRef rc As IRibbonControl, _
                                                 ByRef Visible As Variant)

   Visible = CurrentRibbonWatcher.GetRibbonGroupVisible(rc)

End Sub

Public Sub RibbonWatcherCallBack_EditBox_getText(ByRef rc As IRibbonControl, ByRef Text As Variant)
   CurrentRibbonWatcher.RaiseEditBoxGetText rc, Text
End Sub

Public Sub RibbonWatcherCallBack_EditBox_onChange(ByRef rc As IRibbonControl, ByRef Text As Variant)
    CurrentRibbonWatcher.RaiseEditBoxOnChange rc, Text
End Sub


Sub RibbonWatcherCallBack_DropDown_getItemCount(ByRef rc As IRibbonControl, _
                                                ByRef Count)
    CurrentRibbonWatcher.RaiseDropDownGetItemCount rc, Count
End Sub

Sub RibbonWatcherCallBack_DropDown_getItemID(ByRef rc As IRibbonControl, _
                                             ByRef Index As Integer, _
                                             ByRef ItemID)
    CurrentRibbonWatcher.RaiseDropDownGetItemID rc, Index, ItemID
End Sub

Sub RibbonWatcherCallBack_DropDown_getItemLabel(rc As IRibbonControl, _
                           Index As Integer, _
                           ByRef Label)
    CurrentRibbonWatcher.RaiseDropDownGetItemLabel rc, Index, Label
End Sub

Sub RibbonWatcherCallBack_DropDown_onAction(ByRef rc As IRibbonControl, _
                                            ByRef selectedId As String, _
                                            ByRef selectedIndex As Integer)
    CurrentRibbonWatcher.RaiseDropDownOnAction rc, selectedId, selectedIndex
End Sub

Public Sub RibbonWatcherCallBack_DropDown_getSelectedItemID(ByRef rc As IRibbonControl, ByRef ItemID)
   CurrentRibbonWatcher.RaiseDropDownGetSelectedItemID rc, ItemID
End Sub

Public Sub RibbonWatcherCallBack_DropDown_getSelectedItemIndex(ByRef rc As IRibbonControl, ByRef Index)
   Index = 0
   CurrentRibbonWatcher.RaiseDropDownGetSelectedItemIndex rc, Index
End Sub
