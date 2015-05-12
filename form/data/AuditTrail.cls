VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AuditTrail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/data/AuditTrail.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>ADODB</name><major>2</major><minor>1</minor><guid>{00000201-0000-0010-8000-00AA006D2EA4}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------

Option Explicit
Option Compare Text

' ================================================
' Code by Andreas Vogt Email: info@accessblog.de
' ================================================

Private WithEvents m_frm As Form

Private m_Identifier As Long
Private m_IdFieldName As String
Private m_UserAction As String
Private m_AuditTabExists As Boolean

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

Private Property Get AuditTabExists() As Boolean
    If Not m_AuditTabExists Then
        m_AuditTabExists = Not IsNull(DLookup("[Name]", "MSysObjects", "[Name] = 'tblAuditTrailLog' AND (Type = 1 Or Type = 6)"))
    End If
    AuditTabExists = m_AuditTabExists
End Property

Private Property Get UserAction() As String
    UserAction = m_UserAction
End Property

Private Property Let UserAction(ByVal UserAction_ As String)
    m_UserAction = UserAction_
End Property

Private Property Get LastIdentifier() As Long
    LastIdentifier = m_Identifier
End Property

Private Property Let LastIdentifier(ByVal lastident As Long)
    m_Identifier = lastident
End Property

Private Sub Class_Initialize()
    Dim strSQL As String
    strSQL = "Create Table tblAuditTrailLog (" & _
             "ID COUNTER CONSTRAINT ID PRIMARY KEY, " & _
             "AuditTime date, " & _
             "UserName Text, " & _
             "FormName Text, " & _
             "FieldName Text, " & _
             "ActionType Text, " & _
             "RecordID Text, " & _
             "OldValue Text, " & _
             "NewValue Text" & _
             ")"
    If Not AuditTabExists Then
        CurrentProject.Connection.Execute strSQL
    End If
End Sub

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
        lngInsertedID = WriteNewDeleteLog
        If UserAction = "DELETE" Then LastIdentifier = lngInsertedID
    End Select
End Sub

Private Sub WriteEditLog(ByRef CTL_ As Access.Control)
    Dim strSQL As String
    strSQL = "Insert into tblAuditTrailLog (AuditTime, UserName, FormName, FieldName, ActionType, RecordID, OldValue, NewValue) Values (" & _
             "'" & dtmCurrentDateTime & "', " & _
             "'" & strUserName & "', " & _
             "'" & m_frm.Name & "', " & _
             "'" & CTL_.ControlSource & "', " & _
             "'" & UserAction & "', " & _
             "'" & m_frm.Controls(m_IdFieldName).Value & "', " & _
             "'" & CTL_.OldValue & "', " & _
             "'" & CTL_.Value & "' " & _
             ")"
    CurrentProject.Connection.Execute strSQL
    
End Sub

Private Function WriteNewDeleteLog() As Long
    Dim strSQL As String
    Dim RecordsAffected As Long
    Dim cnn As New ADODB.Connection
    Dim rstTemp As ADODB.Recordset

    strSQL = "Insert into tblAuditTrailLog (AuditTime, UserName, FormName, ActionType, RecordID) Values (" & _
             "'" & dtmCurrentDateTime & "', " & _
             "'" & strUserName & "', " & _
             "'" & m_frm.Name & "', " & _
             "'" & UserAction & "', " & _
             "'" & m_frm.Recordset.Fields(m_IdFieldName).Value & "')"
    Set cnn = CurrentProject.Connection
    cnn.Execute strSQL, RecordsAffected
    If RecordsAffected > 0 Then
        Set rstTemp = cnn.Execute("SELECT @@IDENTITY")
        WriteNewDeleteLog = rstTemp(0)
        rstTemp.Close
    End If
    Set rstTemp = Nothing
    Set cnn = Nothing
End Function

Private Sub DataRedoDelete()
    Dim rstTemp As ADODB.Recordset

    Set rstTemp = New ADODB.Recordset
    With rstTemp
        .Open "SELECT * FROM tblAuditTrailLog", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
        .Find "ID = " & LastIdentifier
        If Not .EOF Then
            .Delete
        End If
        .Close
    End With
    Set rstTemp = Nothing
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
