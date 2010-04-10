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
' Class Module: RibbonWatcher (Josef Pötzl, 2010-04-10)
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
Option Compare Database
Option Explicit

Private m_RibbonUI As IRibbonUI

Private Const m_LargeButtonCount As Long = 6
Private Const m_NormalButtonCount As Long = 9
Private Const m_ComboboxItemCount As Long = 0
Private Const m_RibbonControlInfoArraySize As Long = m_LargeButtonCount + m_NormalButtonCount + m_ComboboxItemCount
'/**
' * @todo Anzahl der Buttons als Eigenschaft gestalten
'**/

Private m_DaoDatabase As DAO.Database
Private ribbonControlInfoArray() As RibbonControlInfo
Private m_ribbonControlInfoArrayFilled As Boolean

Private Sub Class_Initialize()
   initRibbonControlInfoArray m_RibbonControlInfoArraySize
   'fillRibbonInfoArray ... hier weg, damit Tabelle bzw. DaoDatabase zuvor per Eigenschaft eingestellt werden kann
End Sub

Private Sub initRibbonControlInfoArray(ByVal lArraySize As Long)
   Dim i As Long
   ReDim ribbonControlInfoArray(1 To lArraySize)
   For i = 1 To lArraySize
      Set ribbonControlInfoArray(i) = New RibbonControlInfo
   Next
End Sub

Private Sub fillRibbonInfoArray()

   Dim rst As DAO.Recordset
   Dim strSQL As String
   Dim currentArrayCnt As Long
   Dim strID As String
   
On Error GoTo HandleErr

   strSQL = "select TOP " & m_RibbonControlInfoArraySize & _
            " RibbonControlId, rcLabel, rcOnAction, rcImageMso" & _
            " from usysUserRibbonClicks" & _
            " where UserId='" & Me.User & "'" & _
            " order by clicks desc, RibbonControlId asc"

   Set rst = DaoDatabase.OpenRecordset(strSQL)
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
         .visible = True
         
         If .Id <> strID Then
            .Id = strID
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
      If createUserRibbonClickLogTable() Then Resume
   End If
   
   Err.Raise Err.Number, "fillRibbonInfoArray->" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
   Resume ExitHere
   
End Sub

Private Function createUserRibbonClickLogTable() As Boolean

   Dim strSQL As String
   
   strSQL = "create table usysUserRibbonClicks (UserId varchar(50) not null, RibbonControlId varchar(150) not null, clicks int not null , rcLabel varchar(255), rcOnAction varchar(255), rcImageMso varchar(255)" & _
            ", CONSTRAINT PK_usysUserRibbonClicks PRIMARY KEY (UserId, RibbonControlId))"
   DaoDatabase.Execute strSQL, dbFailOnError
   
   strSQL = "CREATE INDEX IX_Sort ON usysUserRibbonClicks (UserId asc, clicks asc)"
   DaoDatabase.Execute strSQL, dbFailOnError
   
   createUserRibbonClickLogTable = True

End Function

Public Sub RefreshRibbonWatcherControl(rwcNr As Long)
   If Not (m_RibbonUI Is Nothing) Then
      m_RibbonUI.InvalidateControl "RWC" & rwcNr
   End If
End Sub

Public Property Get User() As String
   User = "test"
   '/**
   ' * @todo Windows-User (Environ("username")) bzw. DB-User (CurrentUser) einsetzen - zuvor per Ereignis nachfragen
   '**/
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

Public Property Set DaoDatabase(newRef As DAO.Database)
   Set m_DaoDatabase = newRef
End Property

Public Sub Refresh()
   
   fillRibbonInfoArray
   
   If Not (m_RibbonUI Is Nothing) Then
      m_RibbonUI.Invalidate
   End If
   
End Sub

Public Property Let UsedRibbonItemXml(ByVal xmlString As String)
   
   Dim strID As String
   Dim strImageMso As String
   Dim strOnAction As String
   Dim strLabel As String
   
   Dim lngPos As Long
   
   If Len(xmlString) = 0 Then
      Exit Property
   End If
   
   Do While InStr(1, xmlString, "  ") > 0
      xmlString = Replace(xmlString, "  ", " ")
   Loop
   
   xmlString = Replace(xmlString, "= """, "=""")
   xmlString = Replace(xmlString, """ =", """=")
   
   strID = getXmlTagValue(xmlString, "id")
   If StrPtr(strID) = 0 Then
      Exit Property
   End If
 
   strLabel = getXmlTagValue(xmlString, "label")
   strOnAction = getXmlTagValue(xmlString, "onAction")
   strImageMso = getXmlTagValue(xmlString, "imageMso")
   
   saveRibbonClick strID, strLabel, strOnAction, strImageMso
   
End Property

Private Function getXmlTagValue(ByVal sourceString As String, TagName As String) As String

   Dim lngStartPos As Long
   
   lngStartPos = InStr(1, sourceString, TagName & "=""")
   If lngStartPos = 0 Then
      getXmlTagValue = vbNullString
   End If
   
   lngStartPos = lngStartPos + Len(TagName) + 2
   getXmlTagValue = Mid(sourceString, lngStartPos, InStr(lngStartPos, sourceString, """") - lngStartPos)
   
End Function

Private Sub saveRibbonClick(ByRef rcID As String, rcLabel As String, rcOnAction As String, rcImageMso As String)
   
   Dim db As DAO.Database
   Dim strSQL As String
   
On Error GoTo HandleErr

   strSQL = "update usysUserRibbonClicks set clicks=clicks+1"
   If Len(rcLabel) > 0 Then
      strSQL = strSQL & ", rcLabel='" & Replace(rcLabel, "'", "''") & "'"
   End If
   If Len(rcLabel) > 0 Then
      strSQL = strSQL & ", rcOnAction='" & Replace(rcOnAction, "'", "''") & "'"
   End If
   If Len(rcLabel) > 0 Then
      strSQL = strSQL & ", rcImageMso='" & Replace(rcImageMso, "'", "''") & "'"
   End If
   strSQL = strSQL & " where UserId = '" & Me.User & "' AND RibbonControlId='" & Replace(rcID, "'", "''") & "'"
   
   Set db = DaoDatabase
   db.Execute strSQL, dbFailOnError
   If db.RecordsAffected = 0 Then
      strSQL = "insert into usysUserRibbonClicks (UserId, RibbonControlId, clicks, rcLabel, rcOnAction, rcImageMso)" & _
               " Values ('" & Me.User & "', '" & Replace(rcID, "'", "''") & "', 1, '" & Replace(rcLabel, "'", "''") & "', '" & Replace(rcOnAction, "'", "''") & "', '" & Replace(rcImageMso, "'", "''") & "')"
      db.Execute strSQL, dbFailOnError
   End If
   
   fillRibbonInfoArray
   
ExitHere:
   Exit Sub

HandleErr:
   If Err.Number = 3078 Then 'Tabelle nicht gefunden
      If createUserRibbonClickLogTable() Then Resume
   End If
   
   Err.Raise Err.Number, "saveRibbonClick->" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
   Resume ExitHere
   
End Sub

Public Property Get RibbonUI() As IRibbonUI
   Set RibbonUI = m_RibbonUI
End Property

Public Property Set RibbonUI(newRef As IRibbonUI)
   Set m_RibbonUI = newRef
End Property

Public Property Get RibbonXml(Optional ByVal TabXmlBefore As String, Optional TabXmlAfter As String) As String

   RibbonXml = "<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" onLoad=""RibbonWatcherCallBack_OnLoad"">" & vbNewLine & _
               "<ribbon startFromScratch=""false""><tabs>" & vbNewLine & _
               TabXmlBefore & vbNewLine & _
               RibbonTabXml & vbNewLine & _
               TabXmlAfter & vbNewLine & _
               "</tabs></ribbon></customUI>"
               
End Property

Public Property Get RibbonTabXml(Optional ByVal sTabName As String = "Zuletzt verwendet") As String

   RibbonTabXml = "<tab id=""RibbonWatcherTab"" label=""" & sTabName & """>" & vbNewLine & _
                  RibbonGroupXml & vbNewLine & _
                  "</tab>"
   
End Property

Public Property Get RibbonGroupXml(Optional ByVal sGroupName As String = "Zuletzt verwendet") As String
   
   Dim i As Long
   Dim strGroupXml As String
   
   strGroupXml = "<group id=""RibbonWatcherGroup"" label=""" & sGroupName & """>" & vbNewLine
   
   For i = 1 To m_LargeButtonCount
      strGroupXml = strGroupXml & getRibbonControlXml("RWC" & i, "large") & vbNewLine
   Next
   
   If m_LargeButtonCount > 0 And m_NormalButtonCount > 0 Then
      strGroupXml = strGroupXml & "<separator id=""RWSEPLN"" />" & vbNewLine
   End If
   
   For i = 1 To m_NormalButtonCount
      strGroupXml = strGroupXml & getRibbonControlXml("RWC" & m_LargeButtonCount + i, "normal") & vbNewLine
   Next
   
   '/**
   ' * @todo Combobox anfügen
   '**/
   
   RibbonGroupXml = strGroupXml & "</group>"
   
End Property

Private Function getRibbonControlXml(sID As String, sSize As String) As String
   getRibbonControlXml = "<button id=""" & sID & """ size=""" & sSize & """ getLabel=""RibbonWatcherCallBack_GetLabel"" onAction=""RibbonWatcherCallBack_OnAction"" getImage=""RibbonWatcherCallBack_GetImages"" getVisible=""RibbonWatcherCallBack_GetVisible"" />"
End Function

Public Function GetRibbonControlLabel(rc As IRibbonControl) As Variant
   GetRibbonControlLabel = GetLabel(Val(Mid(rc.Id, 4)))
End Function

Public Function GetLabel(rwcNr As Long) As String
   If rwcNr > 0 Then
      If Not m_ribbonControlInfoArrayFilled Then fillRibbonInfoArray
      GetLabel = ribbonControlInfoArray(rwcNr).Label
   End If
End Function

Public Function GetRibbonControlImage(rc As IRibbonControl) As Variant
   GetRibbonControlImage = GetImage(Val(Mid(rc.Id, 4)))
End Function

Public Function GetImage(rwcNr As Long) As Variant
   If rwcNr > 0 Then
      If Not m_ribbonControlInfoArrayFilled Then fillRibbonInfoArray
      GetImage = ribbonControlInfoArray(rwcNr).ImageMso
   End If
End Function

Public Function GetRibbonControlVisible(rc As IRibbonControl) As Variant
   GetRibbonControlVisible = GetVisible(Val(Mid(rc.Id, 4)))
End Function

Public Function GetVisible(rwcNr As Long) As Boolean
   If rwcNr > 0 Then
      If Not m_ribbonControlInfoArrayFilled Then fillRibbonInfoArray
      GetVisible = ribbonControlInfoArray(rwcNr).visible
   End If
End Function

Public Sub CallRibbonControlOnAction(rc As IRibbonControl)
   CallOnAction Val(Mid(rc.Id, 4))
End Sub

Public Sub CallOnAction(rwcNr As Long)

   Dim strOnAction As String
   Dim strProc As String
   Dim strOnActionParams As String
   Dim onActionParams As Variant
   Dim params(29) As Variant
   Dim bolWithParams As Boolean
   Dim lngPos As Long
   Dim i As Long
   
   If Not m_ribbonControlInfoArrayFilled Then fillRibbonInfoArray
   
   For i = 0 To 29
      params(i) = getIsMissing()
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
               params(i) = onActionParams(i)
            Next
         End If
         strProc = Left(strProc, lngPos - 1)
      End If
   Else 'Callback
      strProc = strOnAction
      bolWithParams = True
      Set params(0) = ribbonControlInfoArray(rwcNr)
'/**
' * @todo Was ist mit Callback-Aufrufen, die noch weitere Parameter benötigen?
'**/

   End If
  

   Application.Run strProc, params(0), params(1), params(2), params(3), params(4), params(5), params(6), params(7), params(8), params(9), _
                            params(10), params(11), params(12), params(13), params(14), params(15), params(16), params(17), params(18), params(19), _
                            params(20), params(21), params(22), params(23), params(24), params(25), params(26), params(27), params(28), params(29)
   
End Sub

Private Function getIsMissing(Optional x As Variant) As Variant
   getIsMissing = x
End Function
