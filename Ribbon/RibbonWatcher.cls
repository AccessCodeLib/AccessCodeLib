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
' RibbonWatcher
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>Ribbon/RibbonWatcher.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>Ribbon/RibbonControlInfo.cls</use>
'  <use>Ribbon/ribbon_RibbonWatcher.bas</use>
'  <ref><name>Office</name><major>2</major><minor>4</minor><guid>{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}</guid></ref>
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

Private RibbonControlInfoArray() As RibbonControlInfo
'
Private Sub Class_Initialize()
   Dim i As Long
   ReDim RibbonControlInfoArray(1 To m_RibbonControlInfoArraySize)
   For i = 1 To m_RibbonControlInfoArraySize
      Set RibbonControlInfoArray(i) = New RibbonControlInfo
   Next
   fillRibbonInfoArray
End Sub

Public Property Get User() As String
   User = "test"
   '/**
   ' * @todo Windows-User bzw. DB-User einsetzen
   '**/
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

Public Property Set RibbonUI(newRef As IRibbonUI)
   Set m_RibbonUI = newRef
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
   
   Dim strSQL As String

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
   strSQL = strSQL & " where User = '" & Me.User & "' AND RibbonControl='" & Replace(rcID, "'", "''") & "'"
   
   Dim db As DAO.Database
   Set db = CurrentDb
   db.Execute strSQL, dbFailOnError
   If db.RecordsAffected = 0 Then
      strSQL = "insert into usysUserRibbonClicks (user, RibbonControl, clicks, rcLabel, rcOnAction, rcImageMso)" & _
               " Values ('" & Me.User & "', '" & Replace(rcID, "'", "''") & "', 1, '" & Replace(rcLabel, "'", "''") & "', '" & Replace(rcOnAction, "'", "''") & "', '" & Replace(rcImageMso, "'", "''") & "')"
      db.Execute strSQL, dbFailOnError
   End If
   
   fillRibbonInfoArray
   
End Sub

Private Sub fillRibbonInfoArray()
   
   Dim rst As DAO.Recordset
   Dim strSQL As String
   Dim currentArrayCnt As Long
   Dim strID As String
   
   strSQL = "select TOP " & m_RibbonControlInfoArraySize & _
            " RibbonControl, rcLabel, rcOnAction, rcImageMso" & _
            " from usysUserRibbonClicks" & _
            " where User='" & Me.User & "'" & _
            " order by clicks desc, RibbonControl asc"
'/**
' * @todo Tabelle per Code erstellen, falls sie fehlt
'**/
   Set rst = CurrentDb.OpenRecordset(strSQL)
   currentArrayCnt = 0
   Do While Not rst.EOF
      currentArrayCnt = currentArrayCnt + 1
      If currentArrayCnt > m_RibbonControlInfoArraySize Then
         Exit Do
      End If
      
      With RibbonControlInfoArray(currentArrayCnt)
         strID = rst.Fields("RibbonControl")
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
   
End Sub

Public Sub RefreshRibbonWatcherControl(rwcNr As Long)
   If Not (m_RibbonUI Is Nothing) Then
      m_RibbonUI.InvalidateControl "RWC" & rwcNr
   End If
End Sub

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

Public Function GetRibbonControlLabel(control As IRibbonControl) As Variant
   GetRibbonControlLabel = GetLabel(Val(Mid(control.Id, 4)))
End Function

Public Function GetLabel(rwcNr As Long) As String
   If rwcNr > 0 Then
      GetLabel = RibbonControlInfoArray(rwcNr).Label
   End If
End Function

Public Function GetRibbonControlImage(control As IRibbonControl) As Variant
   GetRibbonControlImage = GetImage(Val(Mid(control.Id, 4)))
End Function

Public Function GetImage(rwcNr As Long) As Variant
   If rwcNr > 0 Then
      GetImage = RibbonControlInfoArray(rwcNr).ImageMso
   End If
End Function

Public Function GetRibbonControlVisible(control As IRibbonControl) As Variant
   GetRibbonControlVisible = GetVisible(Val(Mid(control.Id, 4)))
End Function

Public Function GetVisible(rwcNr As Long) As Boolean
   If rwcNr > 0 Then
      GetVisible = RibbonControlInfoArray(rwcNr).visible
   End If
End Function

Public Sub CallRibbonControlOnAction(control As IRibbonControl)
   CallOnAction Val(Mid(control.Id, 4))
End Sub

Public Sub CallOnAction(rwcNr As Long)

   Dim strOnAction As String
   Dim strProc As String
   Dim strOnActionParams As String
   Dim params As Variant
   Dim bolWithParams As Boolean
   Dim lngPos As Long
   Dim i As Long
   
   strOnAction = RibbonControlInfoArray(rwcNr).OnAction
   If Left(strOnAction, 1) = "=" Then 'Kein Callback, sondern Funktionsaufruf
      strProc = Mid(strOnAction, 2)
      If Right(strProc, 1) = ")" Then strProc = Left(strProc, Len(strProc) - 1)
      lngPos = InStr(1, strProc, "(")
      If lngPos > 0 Then
         strOnActionParams = Mid(strProc, lngPos + 1, Len(strProc) - lngPos)
         bolWithParams = (Len(strOnActionParams) > 0)
         If bolWithParams Then
            params = Split(strOnActionParams, ",")
            For i = 0 To UBound(params)
               If Left(params(i), 1) Like "['""]" And Right(params(i), 1) Like "['""]" Then '@todo: genauer prüfen!
                  params(i) = Mid(params(i), 2, Len(params(i)) - 2)
               End If
            Next
         End If
         strProc = Left(strProc, lngPos - 1)
      End If
   Else 'Callback
      strProc = strOnAction
      bolWithParams = True
      ReDim params(0)
      Set params(0) = RibbonControlInfoArray(rwcNr)
      '/**
      ' * @todo Wie ruft man eine CallBack-Funktion auf, wenn man keine control-Referenz hat?
      '**/
   End If
  
   If bolWithParams Then
      Select Case UBound(params)
         Case 0
            Application.Run strProc, params(0)
         Case 1
            Application.Run strProc, params(0), params(1)
         Case 2
            Application.Run strProc, params(0), params(1), params(2)
         Case 3
            Application.Run strProc, params(0), params(1), params(2), params(3)
         Case Else
            MsgBox "hoppla, die Sub 'CallOnAction' ist noch nicht fertig", vbCritical
   '/**
   ' * @todo Case für Application.Run erweitern
   '**/
      End Select
   
   Else
      Application.Run strProc
   End If
   
End Sub
