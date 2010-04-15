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
Option Compare Database
Option Explicit

Private m_CurrentRibbonWatcher As RibbonWatcher

'RibbonWatcher-Instanz
Public Property Get CurrentRibbonWatcher() As RibbonWatcher
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

   Set CurrentRibbonWatcher.RibbonUI = ribbon

End Sub

Public Sub RibbonWatcherCallBack_OnAction(ByRef rc As IRibbonControl)
   
   CurrentRibbonWatcher.CallRibbonControlOnAction rc

End Sub

Public Sub RibbonWatcherCallBack_GetLabel(ByRef rc As IRibbonControl, _
                                          ByRef Label As Variant)

   Label = CurrentRibbonWatcher.GetRibbonControlLabel(rc)

End Sub

Public Sub RibbonWatcherCallBack_GetImages(ByRef rc As IRibbonControl, _
                                           ByRef image As Variant)

   image = CurrentRibbonWatcher.GetRibbonControlImage(rc)

End Sub

Public Sub RibbonWatcherCallBack_GetVisible(ByRef rc As IRibbonControl, _
                                            ByRef visible As Variant)

   visible = CurrentRibbonWatcher.GetRibbonControlVisible(rc)

End Sub

Public Sub RibbonWatcherCallBack_GetGroupVisible(ByRef rc As IRibbonControl, _
                                                 ByRef visible As Variant)

   visible = True

End Sub

'Public Sub RibbonWatcherCallBack_DropDown_OnAction( _
'                      ByRef rc As IRibbonControl, _
'                      ByRef selectedId As String, _
'                      ByRef selectedIndex As Integer)
'
'
'   Select Case rc.Id
'      Case "xxx"
'
'      Case Else
'         MsgBox rc.Id
'
'   End Select
'
'End Sub
'
'Public Sub RibbonWatcherCallBack_DropDown_GetSelectedItemIndex( _
'                           ByRef rc As IRibbonControl, _
'                           ByRef Index As Variant)
'
'   ' Callback getSelectedItemIndex
'
'   Select Case rc.Id
'      Case "xxx"
'         'Index = xxx + 1
'      Case Else
'
'   End Select
'
'
'End Sub
'
'Public Sub RibbonWatcherCallBack_DropDown_GetItemCount( _
'                     ByRef rc As IRibbonControl, _
'                     ByRef Count As Variant)
'
'
'   Select Case rc.Id
'      Case "xxx"
'         'Count = xxx + 1
'      Case Else
'         MsgBox "DropDown_GetItemCount: " & rc.Id
'   End Select
'
'
'End Sub
'
'Public Sub RibbonWatcherCallBack_DropDown_GetItemLabel(rc As IRibbonControl, _
'                     Index As Integer, _
'                     ByRef Label As Variant)
'
'
'   Select Case rc.Id
'      Case "xxx"
'         If Index < 1 Then
'            Label = "xxx"
'         Else
'            'label = xxx(Index - 1).xxx
'         End If
'      Case Else
'         MsgBox "DropDown_GetItemLabel: " & rc.Id
'   End Select
'
'
'End Sub
'
'Public Sub RibbonWatcherCallBack_DropDown_GetItemScreentip(rc As IRibbonControl, _
'                        Index As Integer, _
'                        ByRef screentip As Variant)
'
'   Select Case rc.Id
'      Case "xxx"
'         If Index < 0 Then
'            screentip = "N/A"
'         Else
'            'screentip = xxx(Index - 1).xxx
'         End If
'      Case Else
'         MsgBox "DropDown_GetItemScreentip: " & rc.Id
'   End Select
'
'
'End Sub
'
'Public Sub RibbonWatcherCallBack_DropDown_GetItemSupertip(rc As IRibbonControl, _
'                       Index As Integer, _
'                       ByRef supertip As Variant)
'
'   Select Case rc.Id
'      Case "xxx"
'         'supertip = xxx(Index - 1).xxx
'      Case Else
'         MsgBox "DropDown_GetItemSupertip: " & rc.Id
'   End Select
'
'
'End Sub
'
'Public Sub RibbonWatcherCallBack_EditBox_getText(rc As IRibbonControl, ByRef text)
'
'   Select Case rc.Id
'      Case "TEST123"
'
'   End Select
'
'End Sub
'
'Public Sub RibbonWatcherCallBack_EditBox_onChange(rc As IRibbonControl, _
'                                    strText As String)
'
'    Select Case rc.Id
'        Case "TEST123"
'
'    End Select
'
'End Sub
