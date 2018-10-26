VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: RibbonWatcher (Josef Pötzl)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Zeigt die zuletzt verwendeten Ribbon-Controls in einer Ribbon-Gruppe bzw. Tab an.
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>ribbon/RibbonWatcher.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>ribbon/RibbonControlInfo.cls</use>
'  <use>ribbon/ribbon_RibbonWatcher.bas</use>
'  <ref><name>Office</name><major>2</major><minor>4</minor><guid>{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}</guid></ref>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private m_RibbonUI As IRibbonUI

'Benutzer
Private m_User As Variant

' Anzahl der Ribbon-Controls:
Private m_LargeButtonCount As Long
Private m_NormalButtonCount As Long
Private m_ComboboxItemCount As Long
Private m_RibbonControlInfoArraySize As Long 'Damit nicht immer neu berechnet werden muss

'Standardwerte, falls die Werte nich von außen eingestellt werden
Private Const m_DEFAULT_LargeButtonCount As Long = 6
Private Const m_DEFAULT_NormalButtonCount As Long = 9
Private Const m_DEFAULT_ComboboxItemCount As Long = 0

Private m_DaoDatabase As DAO.Database
Private ribbonControlInfoArray() As RibbonControlInfo
Private m_ribbonControlInfoArrayFilled As Boolean

Private m_EventGroupCollection As Collection

Public Event RibbonTabChanged(ByVal RibbonTabID As String)
Public Event SaveRibbonControlAction(ByVal RibbonControlId As String, ByRef Cancel As Boolean)

Public Event GetVisible(ByVal rc As IRibbonControl, ByRef Visible As Boolean)
Public Event OnAction(ByVal rc As IRibbonControl)

Public Event EditBoxOnChange(ByVal rc As IRibbonControl, ByRef strText As Variant)
Public Event EditBoxGetText(ByVal rc As IRibbonControl, ByRef Text)

Public Event DropDownGetItemCount(ByVal rc As IRibbonControl, ByRef Count)
Public Event DropDownGetItemId(ByVal Control As IRibbonControl, ByRef Index As Integer, ByRef ItemID)
Public Event DropDownGetItemLabel(ByVal Control As IRibbonControl, ByRef Index As Integer, ByRef Label)
Public Event DropDownGetSelectedItemID(ByVal rc As IRibbonControl, ByRef selectedId)
Public Event DropDownGetSelectedItemIndex(ByVal rc As IRibbonControl, ByRef selectedIndex As Variant)
Public Event DropDownOnAction(ByVal rc As IRibbonControl, ByRef selectedId As String, ByRef selectedIndex As Variant)

#If VBA7 Then
   Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal nCount As Long)
#Else
   Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal nCount As Long)
#End If

Private Sub Class_Initialize()

   Set m_EventGroupCollection = New Collection

   'Standardwerte einstellen:
   InitRibbonGroupLayout
   
End Sub

Public Sub InitRibbonGroupLayout(Optional ByVal lLargeButtonCount As Long = m_DEFAULT_LargeButtonCount, _
                                 Optional ByVal lNormalButtonCount As Long = m_DEFAULT_NormalButtonCount, _
                                 Optional ByVal lComboboxItemCount As Long = m_DEFAULT_ComboboxItemCount)
   
   m_LargeButtonCount = lLargeButtonCount
   m_NormalButtonCount = lNormalButtonCount
   m_ComboboxItemCount = lComboboxItemCount
   m_RibbonControlInfoArraySize = m_LargeButtonCount + m_NormalButtonCount + m_ComboboxItemCount
   
   'Array initialisieren
   InitRibbonControlInfoArray m_RibbonControlInfoArraySize
   
End Sub

Private Sub InitRibbonControlInfoArray(ByVal lArraySize As Long)
   Dim i As Long
   ReDim ribbonControlInfoArray(1 To lArraySize)
   For i = 1 To lArraySize
      Set ribbonControlInfoArray(i) = New RibbonControlInfo
   Next
   m_ribbonControlInfoArrayFilled = False
End Sub

Private Sub FillRibbonInfoArray()

   Dim db As DAO.Database
   Dim qdf As DAO.QueryDef
   Dim rst As DAO.Recordset
   Dim strSQL As String
   Dim currentArrayCnt As Long
   Dim strID As String
   
On Error GoTo HandleErr

   If Not AppLoadedForRibbonWatcher Then
      Exit Sub
   End If

   strSQL = "select TOP " & m_RibbonControlInfoArraySize & _
            " RibbonControlId, rcLabel, rcOnAction, rcImageMso" & _
            " from usysUserRibbonClicks" & _
            " where UserId=[UserParam]" & _
            " order by clicks desc, RibbonControlId asc"

   Set db = Me.DaoDatabase
   Set qdf = db.CreateQueryDef("", strSQL)
   qdf.Parameters("UserParam").Value = Me.User

   Set rst = qdf.OpenRecordset()
   currentArrayCnt = 0
   Do While Not rst.EOF
      currentArrayCnt = currentArrayCnt + 1
      If currentArrayCnt > m_RibbonControlInfoArraySize Then
         Exit Do
      End If
      
      With ribbonControlInfoArray(currentArrayCnt)
         strID = rst.Fields("RibbonControlId")
         .Label = Nz(rst.Fields("rcLabel"), "---")
         .OnAction = Nz(rst.Fields("rcOnAction"), vbNullString)
         .ImageMso = Nz(rst.Fields("rcImageMso"), vbNullString)
         .Visible = True
         
         If .ID <> strID Then
            .ID = strID
            RefreshRibbonWatcherControl currentArrayCnt
         End If
   
      End With
      
      rst.MoveNext
   Loop
   rst.Close
   
   m_ribbonControlInfoArrayFilled = True
   
ExitHere:
   Exit Sub

HandleErr:
   If Err.Number = 3078 Then 'Tabelle nicht gefunden
      If CreateUserRibbonClickLogTable() Then Resume
   End If
   
   Err.Raise Err.Number, "fillRibbonInfoArray->" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
   Resume ExitHere
   
End Sub

Private Function CreateUserRibbonClickLogTable() As Boolean

   Dim strSQL As String
   
   strSQL = "create table usysUserRibbonClicks (UserId varchar(50) not null, RibbonControlId varchar(150) not null, clicks int not null , rcLabel varchar(255), rcOnAction varchar(255), rcImageMso varchar(255)" & _
            ", CONSTRAINT PK_usysUserRibbonClicks PRIMARY KEY (UserId, RibbonControlId))"
   DaoDatabase.Execute strSQL, dbFailOnError
   
   strSQL = "CREATE INDEX IX_Sort ON usysUserRibbonClicks (UserId asc, clicks asc)"
   DaoDatabase.Execute strSQL, dbFailOnError
   
   CreateUserRibbonClickLogTable = True

End Function

Public Sub RefreshRibbonWatcherControl(rwcNr As Long)
   If Not (m_RibbonUI Is Nothing) Then
      m_RibbonUI.InvalidateControl "RWC" & rwcNr
   End If
End Sub

Public Property Get User() As Variant 'Variant, damit Text oder Zahl möglich wird
   '/**
   ' * @todo Windows-User (Environ("username")) bzw. DB-User (CurrentUser) einsetzen - eventuell zuvor per Ereignis nachfragen
   '**/
   If IsEmpty(m_User) Then
      m_User = Environ("username")
      If Len(m_User) = 0 Then m_User = CurrentUser
   End If
   User = m_User
End Property

Public Property Let User(ByVal NewValue As Variant)
   m_User = NewValue
End Property

Public Property Get DaoDatabase() As DAO.Database
'/**
' * @todo bei Nothing eventuell noch zusätzlich mit Ereignis abfragen
'**/
   If m_DaoDatabase Is Nothing Then
      Set m_DaoDatabase = CurrentDb
   End If
   Set DaoDatabase = m_DaoDatabase
End Property

Public Property Set DaoDatabase(NewRef As DAO.Database)
   Set m_DaoDatabase = NewRef
End Property

Public Sub Refresh()
   
   FillRibbonInfoArray
   
   If Not (m_RibbonUI Is Nothing) Then
      m_RibbonUI.Invalidate
   End If
   
End Sub

Public Property Let UsedRibbonItemXml(ByVal XmlString As String)
   
   Dim strID As String
   Dim strImageMso As String
   Dim strOnAction As String
   Dim strLabel As String
      
   Dim lngPos As Long
   
   If Len(XmlString) = 0 Then
      Exit Property
   End If
   
   Do While InStr(1, XmlString, "  ") > 0
      XmlString = Replace(XmlString, "  ", " ")
   Loop
   
   XmlString = Replace(XmlString, "= """, "=""")
   XmlString = Replace(XmlString, """ =", """=")
   
   strID = GetXmlTagValue(XmlString, "id")
   If StrPtr(strID) = 0 Then
      Exit Property
   End If
 
   strLabel = GetXmlTagValue(XmlString, "label")
   strOnAction = GetXmlTagValue(XmlString, "onAction")
   strImageMso = GetXmlTagValue(XmlString, "imageMso")
   
   SaveRibbonClick strID, strLabel, strOnAction, strImageMso
   
End Property

Private Function GetXmlTagValue(ByVal SourceString As String, TagName As String) As String

   Dim lngStartPos As Long
   
   lngStartPos = InStr(1, SourceString, TagName & "=""")
   If lngStartPos = 0 Then
      GetXmlTagValue = vbNullString
      Exit Function
   End If
   
   lngStartPos = lngStartPos + Len(TagName) + 2
   GetXmlTagValue = Mid(SourceString, lngStartPos, InStr(lngStartPos, SourceString, """") - lngStartPos)
   
End Function

Private Sub SaveRibbonClick(ByRef rcID As String, rcLabel As String, rcOnAction As String, rcImageMso As String)
   
   Dim db As DAO.Database
   Dim qdf As DAO.QueryDef
   Dim strSQL As String
   Dim bolCancel As Boolean
   
On Error GoTo HandleErr

   RaiseEvent SaveRibbonControlAction(rcID, bolCancel)
   If bolCancel Then
      Exit Sub
   End If

   strSQL = "update usysUserRibbonClicks set clicks=clicks+1"
   If Len(rcLabel) > 0 Then
      strSQL = strSQL & ", rcLabel=[rcLabelParam]"
   End If
   If Len(rcOnAction) > 0 Then
      strSQL = strSQL & ", rcOnAction=[rcOnActionParam]"
   End If
   If Len(rcImageMso) > 0 Then
      strSQL = strSQL & ", rcImageMso=[rcImageMsoParam]"
   End If
   strSQL = strSQL & " where UserId = [UserParam] AND RibbonControlId=[RibbonControlIdParam]"
   
   Set db = DaoDatabase
   Set qdf = db.CreateQueryDef("", strSQL)
   
   qdf.Parameters("UserParam").Value = Me.User 'Damit ist Datentyp vom Feld [UserId] egal (nützlich falls Userkennung durch Zahl dargestellt wird)
   qdf.Parameters("RibbonControlIdParam").Value = rcID
   If Len(rcLabel) > 0 Then
      qdf.Parameters("rcLabelParam").Value = rcLabel
   End If
   If Len(rcOnAction) > 0 Then
      qdf.Parameters("rcOnActionParam").Value = rcOnAction
   End If
   If Len(rcImageMso) > 0 Then
      qdf.Parameters("rcImageMsoParam").Value = rcImageMso
   End If
   
   qdf.Execute dbFailOnError
   
   If qdf.RecordsAffected = 0 Then
      qdf.SQL = "insert into usysUserRibbonClicks (UserId, RibbonControlId, clicks, rcLabel, rcOnAction, rcImageMso)" & _
                " Values ([UserParam], [RibbonControlIdParam], 1, [rcLabelParam], [rcOnActionParam], [rcImageMsoParam])"
      qdf.Parameters("UserParam").Value = Me.User
      qdf.Parameters("RibbonControlIdParam").Value = rcID
      qdf.Parameters("rcLabelParam").Value = rcLabel
      qdf.Parameters("rcOnActionParam").Value = rcOnAction
      qdf.Parameters("rcImageMsoParam").Value = rcImageMso
      qdf.Execute dbFailOnError
   End If
   
   FillRibbonInfoArray
   
ExitHere:
   Exit Sub

HandleErr:
   If Err.Number = 3078 Then 'Tabelle nicht gefunden
      If CreateUserRibbonClickLogTable() Then Resume
   End If
   
   Err.Raise Err.Number, "saveRibbonClick->" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
   Resume ExitHere
   
End Sub


Public Property Get RibbonUI() As IRibbonUI
   If m_RibbonUI Is Nothing Then
      If Application.TempVars("RibbonUI").Value > 0 Then
         GetRibbonFromPointer Application.TempVars("RibbonUI").Value
      End If
   End If
   Set RibbonUI = m_RibbonUI
End Property

Private Sub GetRibbonFromPointer(ByVal RibbonPtr As Long)

   Dim tmpObj As IRibbonUI

   CopyMemory tmpObj, RibbonPtr, 4
   Set m_RibbonUI = tmpObj
   CopyMemory tmpObj, 0&, 4

End Sub

Public Property Set RibbonUI(ByVal NewRef As IRibbonUI)
   Set m_RibbonUI = NewRef
   Application.TempVars("RibbonUI").Value = ObjPtr(m_RibbonUI) ' Adresse speichern, falls Laufzeitfehler m_RibbonUI leert
End Property

Public Function RibbonXml(Optional ByVal TabXmlBefore As String, Optional TabXmlAfter As String, _
                              Optional ByVal ActivateTabChangeEvents As Boolean = False, _
                              Optional ByVal StartFromScratch As Boolean = False, _
                              Optional ByVal RibbonWatcherTabLabel As String = "Zuletzt verwendet") As String

   Dim strRibbonXml As String
   Dim arrRibbonTabs As Variant
   Dim i As Long
   Dim Pos As Long
   Dim tempTabStringStart As String
   Dim tempTabStringEnd As String
   Dim TempString As String
   Dim tempTabID As String
   
   Dim StartFromScratchString As String
   If StartFromScratch Then
      StartFromScratchString = "true"
   Else
      StartFromScratchString = "false"
   End If

   strRibbonXml = "<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" onLoad=""RibbonWatcherCallBack_OnLoad"">" & vbNewLine & _
               "<ribbon startFromScratch=""" & StartFromScratchString & """><tabs>" & vbNewLine & _
               TabXmlBefore & vbNewLine & _
               RibbonTabXml(RibbonWatcherTabLabel) & vbNewLine & _
               TabXmlAfter & vbNewLine & _
               "</tabs></ribbon></customUI>"
               
   If ActivateTabChangeEvents Then
      arrRibbonTabs = Split(strRibbonXml, "</tab>", , vbTextCompare)
      For i = 0 To (UBound(arrRibbonTabs) - 1)
         tempTabID = Mid(arrRibbonTabs(i), InStr(1, arrRibbonTabs(i), "<tab ", vbTextCompare) + 4)
         tempTabID = Mid(tempTabID, InStr(1, tempTabID, "id=""", vbTextCompare) + 4)
         tempTabID = Left(tempTabID, InStr(1, tempTabID, """", vbTextCompare) - 1)

         Pos = InStr(1, arrRibbonTabs(i), "<group ", vbTextCompare)
         
         If Pos > 0 Then
            tempTabStringStart = Left(arrRibbonTabs(i), Pos + 6)
            TempString = Mid(arrRibbonTabs(i), Pos + 7)
            
            Pos = InStr(1, TempString, "</group>", vbTextCompare)
            tempTabStringEnd = Mid(TempString, Pos)
            TempString = Left(TempString, Pos - 1) '=> alles innerhalb der ersten Group
   
            TempString = TempString & "<button id=""RWEC" & tempTabID & """ tag=""" & tempTabID & """ label=""xxx"" onAction=""RibbonWatcherCallBack_OnAction"" getVisible=""RibbonWatcherCallBack_GetVisible"" />"
            
            TempString = tempTabStringStart & TempString & tempTabStringEnd
            
            arrRibbonTabs(i) = TempString & "</tab>"
         End If
      
      Next
      strRibbonXml = Join(arrRibbonTabs, "")
   End If

   RibbonXml = strRibbonXml
               
End Function

Public Function RibbonTabXml(Optional ByVal TabLabel As String = "Zuletzt verwendet") As String

   RibbonTabXml = "<tab id=""RibbonWatcherTab"" label=""" & TabLabel & """>" & vbNewLine & _
                  RibbonGroupXml & vbNewLine & _
                  "</tab>"
   
End Function

Public Function RibbonGroupXml(Optional ByVal GroupLabel As String = "Zuletzt verwendet") As String
   
   Dim i As Long
   Dim strGroupXml As String
   
   strGroupXml = "<group id=""RibbonWatcherGroup"" label=""" & GroupLabel & """ getVisible=""RibbonWatcherCallBack_GetGroupVisible"">" & vbNewLine
   
   For i = 1 To m_LargeButtonCount
      strGroupXml = strGroupXml & GetRibbonControlXml("RWC" & i, "large") & vbNewLine
   Next
   
   If m_LargeButtonCount > 0 And m_NormalButtonCount > 0 Then
      strGroupXml = strGroupXml & "<separator id=""RWSEPLN"" />" & vbNewLine
   End If
   
   For i = 1 To m_NormalButtonCount
      strGroupXml = strGroupXml & GetRibbonControlXml("RWC" & m_LargeButtonCount + i, "normal") & vbNewLine
   Next

   RibbonGroupXml = strGroupXml & "</group>"
   
End Function

Private Function GetRibbonControlXml(ByVal sID As String, ByVal sSize As String) As String
   GetRibbonControlXml = "<button id=""" & sID & """ size=""" & sSize & """ getLabel=""RibbonWatcherCallBack_GetLabel"" onAction=""RibbonWatcherCallBack_OnAction"" getImage=""RibbonWatcherCallBack_GetImages"" getVisible=""RibbonWatcherCallBack_GetVisible"" />"
End Function

Public Function GetRibbonControlLabel(ByVal rc As IRibbonControl) As Variant
   GetRibbonControlLabel = GetLabel(Val(Mid(rc.ID, 4)))
End Function

Public Function GetLabel(ByVal rwcNr As Long) As String
   If rwcNr > 0 Then
      If Not m_ribbonControlInfoArrayFilled Then FillRibbonInfoArray
      GetLabel = ribbonControlInfoArray(rwcNr).Label
   End If
End Function

Public Function GetRibbonControlImage(ByVal rc As IRibbonControl) As Variant
   GetRibbonControlImage = GetImage(Val(Mid(rc.ID, 4)))
End Function

Public Function GetImage(ByVal rwcNr As Long) As Variant
   If rwcNr > 0 Then
      If Not m_ribbonControlInfoArrayFilled Then FillRibbonInfoArray
      GetImage = ribbonControlInfoArray(rwcNr).ImageMso
   End If
End Function

Public Function GetRibbonGroupVisible(ByVal rc As IRibbonControl) As Boolean
' @TODO: Add property

   GetRibbonGroupVisible = True

End Function

Public Function GetRibbonControlVisible(ByVal rc As IRibbonControl) As Boolean

   Dim bVisible As Boolean

   If Left(rc.ID, 4) = "RWEC" Then
      SetActivatedEventControl rc
      bVisible = False
   ElseIf Left(rc.ID, 3) = "RWC" Then
      bVisible = GetRibbonWatcherControlVisible(Val(Mid(rc.ID, 4)))
   Else
      RaiseGetVisible rc, bVisible
   End If
   
   GetRibbonControlVisible = bVisible

End Function

Public Sub RaiseGetVisible(ByVal rc As IRibbonControl, ByRef Visible As Boolean)
   RaiseEvent GetVisible(rc, Visible)
End Sub

Public Function GetRibbonWatcherControlVisible(ByVal rwcNr As Long) As Boolean
   If rwcNr > 0 Then
      If Not m_ribbonControlInfoArrayFilled Then FillRibbonInfoArray
      GetRibbonWatcherControlVisible = ribbonControlInfoArray(rwcNr).Visible
   End If
End Function

Public Sub CallRibbonControlOnAction(ByVal rc As IRibbonControl)

   If Left(rc.ID, 3) = "RWC" Then
      CallOnAction Val(Mid(rc.ID, 4))
   Else
      RaiseOnAction rc
   End If

End Sub

Public Sub RaiseOnAction(ByVal rc As IRibbonControl)
   RaiseEvent OnAction(rc)
End Sub

Public Sub CallOnAction(rwcNr As Long)

   Dim strOnAction As String
   Dim strProc As String
   Dim strOnActionParams As String
   Dim onActionParams As Variant
   Dim Params(29) As Variant
   Dim bolWithParams As Boolean
   Dim lngPos As Long
   Dim i As Long
   
   If Not m_ribbonControlInfoArrayFilled Then FillRibbonInfoArray
   
   For i = 0 To 29
      Params(i) = GetIsMissing()
   Next
   
   strOnAction = ribbonControlInfoArray(rwcNr).OnAction
   If Left(strOnAction, 1) = "=" Then 'Kein Callback, sondern Funktionsaufruf
      strProc = Mid(strOnAction, 2)
      If Right(strProc, 1) = ")" Then strProc = Left(strProc, Len(strProc) - 1)
      lngPos = InStr(1, strProc, "(")
      If lngPos > 0 Then
         strOnActionParams = Mid(strProc, lngPos + 1, Len(strProc) - lngPos)
         bolWithParams = (Len(strOnActionParams) > 0)
         If bolWithParams Then
            onActionParams = Split(strOnActionParams, ",")
            For i = 0 To UBound(onActionParams)
               If Left(onActionParams(i), 1) Like "['""]" And Right(onActionParams(i), 1) Like "['""]" Then '@todo: genauer prüfen!
                  onActionParams(i) = Mid(onActionParams(i), 2, Len(onActionParams(i)) - 2)
               End If
               Params(i) = onActionParams(i)
            Next
         End If
         strProc = Left(strProc, lngPos - 1)
      End If
   Else 'Callback
      strProc = strOnAction
      bolWithParams = True
      Set Params(0) = ribbonControlInfoArray(rwcNr)
'/**
' * @todo Was ist mit Callback-Aufrufen, die noch weitere Parameter benötigen?
'**/

   End If
  

   Application.Run strProc, Params(0), Params(1), Params(2), Params(3), Params(4), Params(5), Params(6), Params(7), Params(8), Params(9), _
                            Params(10), Params(11), Params(12), Params(13), Params(14), Params(15), Params(16), Params(17), Params(18), Params(19), _
                            Params(20), Params(21), Params(22), Params(23), Params(24), Params(25), Params(26), Params(27), Params(28), Params(29)
   
End Sub

Private Function GetIsMissing(Optional x As Variant) As Variant
   GetIsMissing = x
End Function

Private Sub SetActivatedEventControl(ByVal rc As IRibbonControl)
   
   Dim Item As IRibbonControl
   Dim bolItemExits As Boolean
   
   For Each Item In m_EventGroupCollection
      If Item.ID = rc.ID Then
         bolItemExits = True
      Else
         Me.RibbonUI.InvalidateControl Item.ID
      End If
   Next
   
   If Not bolItemExits Then
      m_EventGroupCollection.Add rc, rc.ID
   End If
   
   RaiseEvent RibbonTabChanged(rc.Tag)
   
End Sub

' EditBox
Public Sub RaiseEditBoxOnChange(ByVal rc As IRibbonControl, ByRef Text As Variant)
   RaiseEvent EditBoxOnChange(rc, Text)
End Sub

Public Sub RaiseEditBoxGetText(ByVal rc As IRibbonControl, ByRef Text As Variant)
   RaiseEvent EditBoxGetText(rc, Text)
End Sub

'DropDown
Public Sub RaiseDropDownGetItemCount(ByVal rc As IRibbonControl, _
                                     ByRef Count)
   RaiseEvent DropDownGetItemCount(rc, Count)
End Sub

Public Sub RaiseDropDownGetItemID(ByVal rc As IRibbonControl, _
                                  ByRef Index As Integer, _
                                  ByRef ItemID)
   RaiseEvent DropDownGetItemId(rc, Index, ItemID)
End Sub

Public Sub RaiseDropDownGetItemLabel(ByVal rc As IRibbonControl, _
                                     ByRef Index As Integer, _
                                     ByRef Label)
   RaiseEvent DropDownGetItemLabel(rc, Index, Label)
End Sub

Public Sub RaiseDropDownGetSelectedItemID(ByVal rc As IRibbonControl, _
                                          ByRef selectedId)
   RaiseEvent DropDownGetSelectedItemID(rc, selectedId)
End Sub

Public Sub RaiseDropDownGetSelectedItemIndex(ByVal rc As IRibbonControl, _
                                             ByRef selectedIndex As Variant)
   RaiseEvent DropDownGetSelectedItemIndex(rc, selectedIndex)
End Sub

Public Sub RaiseDropDownOnAction(ByVal rc As IRibbonControl, _
                                 ByRef selectedId As String, _
                                 ByRef selectedIndex As Integer)
   RaiseEvent DropDownOnAction(rc, selectedId, selectedIndex)
End Sub
