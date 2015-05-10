Option Explicit

' ================================================
' Code by Andreas Vogt Email: info@accessblog.de
' ================================================

Private WithEvents m_frm As Form

Private m_Identifier As Long
Private m_IdFieldName As String
Private m_UserAction As String

Private rstAuditLog As ADODB.Recordset
Private dtmCurrentDateTime As Date
Private strUserName As String

Public Property Set FormObj(ByRef FRM_ As Access.Form)
    Set m_frm = FRM_
    m_IdFieldName = getIDField(FRM_)
    m_frm.BeforeUpdate = "[Event Procedure]"
    m_frm.AfterUpdate = "[Event Procedure]"
    m_frm.OnDelete = "[Event Procedure]"
    m_frm.AfterDelConfirm = "[Event Procedure]"
End Property

Private Property Get UserAction() As String
    UserAction = m_UserAction
End Property

Private Property Let UserAction(ByVal UserAction_ As String)
    m_UserAction = UserAction_
End Property

Private Property Get lastIdentifier() As Long
    lastIdentifier = m_Identifier
End Property

Private Property Let lastIdentifier(ByVal lastident As Long)
    m_Identifier = lastident
End Property

Private Sub Class_Terminate()
    Set m_frm = Nothing
End Sub

Private Sub m_frm_Delete(Cancel As Integer)
    UserAction = "DELETE"
    DataChanges
End Sub

Private Sub m_frm_AfterDelConfirm(Status As Integer)
    If Status <> acDeleteOK Then Call DataRedoDelete
End Sub

Private Sub m_frm_BeforeUpdate(Cancel As Integer)
    If m_frm.NewRecord Then
        UserAction = "NEW"
        DataChanges
    Else
        UserAction = "EDIT"
        DataChanges
    End If
End Sub

Private Sub DataChanges()
    Dim CTL As Control
    Dim lngInsertedID As Long

    Set rstAuditLog = New ADODB.Recordset
    With rstAuditLog
        .Open "SELECT * FROM tblAuditTrailLog", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
        dtmCurrentDateTime = Now()
        strUserName = Environ("USERNAME")
        Select Case UserAction
        Case "EDIT"
            For Each CTL In m_frm.Controls
                If CTL.Tag = "Audit" Then
                    If Nz(CTL.Value) <> Nz(CTL.OldValue) Then
                        WriteEditLog CTL
                    End If
                End If
            Next CTL
        Case Else
            lngInsertedID = WriteNewDeleteLog(CTL)
            If UserAction = "DELETE" Then lastIdentifier = lngInsertedID
        End Select
    End With
    rstAuditLog.Close
    Set rstAuditLog = Nothing
End Sub

Private Sub WriteEditLog(ByRef CTL_ As Access.Control)
    Dim strSQL As String
    strSQL = "Insert into tblAuditTrailLog (DateTime, UserName, FormName, FieldName, ActionType, RecordID, OldValue, NewValue) Values (" & _
             "'" & dtmCurrentDateTime & "', " & _
             "'" & strUserName & "', " & _
             "'" & m_frm.Name & "', " & _
             "'" & CTL_.ControlSource & "', " & _
             "'" & UserAction & "', " & _
             "'" & m_frm.Controls(m_IdFieldName).Value & "', " & _
             "'" & CTL_.OldValue & "', " & _
             "'" & CTL_.Value & "', " & _
             ")"
    CurrentProject.Connection.Execute strSQL
End Sub

Private Function WriteNewDeleteLog(ByRef CTL_ As Access.Control) As Long
    Dim strSQL As String
    Dim rstTemp As ADODB.Recordset
    
    strSQL = "Insert into tblAuditTrailLog (DateTime, UserName, FormName, ActionType, RecordID) Values (" & _
             "'" & dtmCurrentDateTime & "', " & _
             "'" & strUserName & "', " & _
             "'" & m_frm.Name & "', " & _
             "'" & UserAction & "', " & _
             "'" & m_frm.Recordset.Fields(m_IdFieldName).Value & "')"
    With CurrentProject.Connection
        .Execute strSQL
        If .RecordsAffected > 0 Then
            Set rstTemp = .OpenRecordset("SELECT @@IDENTITY")
            WriteNewDeleteLog = rstTemp(0)
        End If
    End With
    
    rstTemp.Close
    Set rstTemp = Nothing
End Function

Private Sub DataRedoDelete()
    Dim rst As ADODB.Recordset

    Set rstAuditLog = New ADODB.Recordset
    With rstAuditLog
        .Open "SELECT * FROM tblAuditTrailLog", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
        .Find "ID = " & lastIdentifier
        If Not .EOF Then
            .Delete
        End If
    End With

    rstAuditLog.Close
    Set rstAuditLog = Nothing
End Sub

Private Function getIDField(FRM_ As Access.Form) As String
    Dim i As Long
    With FRM_.Recordset
        For i = 0 To .Fields.Count - 1
            If .Fields(i).Type = 4 Then
                getIDField = .Fields(i).Name
                Exit For
            End If
        Next i
    End With
End Function
