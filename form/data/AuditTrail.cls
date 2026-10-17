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
' Class: form.data.AuditTrail
'---------------------------------------------------------------------------------------
'
' Logs data changes of a bound form (edit, new, delete with old/new values, user, time)
' to the table tblAuditTrailLog (ADODB)
'
' Author:
'     Andreas Vogt, adapted by Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/data/AuditTrail.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>ADODB</name><major>2</major><minor>1</minor><guid>{00000201-0000-0010-8000-00AA006D2EA4}</guid></ref>
'  <description>Logs data changes of a bound form (edit, new, delete with old/new values, user, time) to the table tblAuditTrailLog (ADODB)</description>
'</codelib>
'---------------------------------------------------------------------------------------

Option Explicit
Option Compare Text

' ================================================
' Code by Andreas Vogt Email: info@accessblog.de
' ================================================
'
' Changes to the original code (Josef Poetzl, 2026-10-12):
'   - Insert and delete statements use ADODB command parameters with explicit data types
'     instead of string concatenation (quotes in values, locale-independent date, Null values)
'   - Delete of the log entry via parameterized DELETE instead of opening the table and Find

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
    With NewCommand("Insert into tblAuditTrailLog (AuditTime, UserName, FormName, FieldName, ActionType, RecordID, OldValue, NewValue) Values (?, ?, ?, ?, ?, ?, ?, ?)")
        .Parameters.Append .CreateParameter("AuditTime", adDate, adParamInput, , dtmCurrentDateTime)
        .Parameters.Append .CreateParameter("UserName", adVarWChar, adParamInput, 255, strUserName)
        .Parameters.Append .CreateParameter("FormName", adVarWChar, adParamInput, 255, m_frm.Name)
        .Parameters.Append .CreateParameter("FieldName", adVarWChar, adParamInput, 255, CTL_.ControlSource)
        .Parameters.Append .CreateParameter("ActionType", adVarWChar, adParamInput, 255, UserAction)
        .Parameters.Append .CreateParameter("RecordID", adVarWChar, adParamInput, 255, m_frm.Controls(m_IdFieldName).Value & vbNullString)
        .Parameters.Append .CreateParameter("OldValue", adVarWChar, adParamInput, 255, CTL_.OldValue)
        .Parameters.Append .CreateParameter("NewValue", adVarWChar, adParamInput, 255, CTL_.Value)
        .Execute , , adExecuteNoRecords
    End With
End Sub

Private Function WriteNewDeleteLog() As Long
    Dim RecordsAffected As Long
    Dim rstTemp As ADODB.Recordset

    With NewCommand("Insert into tblAuditTrailLog (AuditTime, UserName, FormName, ActionType, RecordID) Values (?, ?, ?, ?, ?)")
        .Parameters.Append .CreateParameter("AuditTime", adDate, adParamInput, , dtmCurrentDateTime)
        .Parameters.Append .CreateParameter("UserName", adVarWChar, adParamInput, 255, strUserName)
        .Parameters.Append .CreateParameter("FormName", adVarWChar, adParamInput, 255, m_frm.Name)
        .Parameters.Append .CreateParameter("ActionType", adVarWChar, adParamInput, 255, UserAction)
        .Parameters.Append .CreateParameter("RecordID", adVarWChar, adParamInput, 255, m_frm.Recordset.Fields(m_IdFieldName).Value & vbNullString)
        .Execute RecordsAffected, , adExecuteNoRecords
    End With

    If RecordsAffected > 0 Then
        Set rstTemp = CurrentProject.Connection.Execute("SELECT @@IDENTITY")
        WriteNewDeleteLog = rstTemp(0)
        rstTemp.Close
    End If
    Set rstTemp = Nothing
End Function

Private Sub DataRedoDelete()
    With NewCommand("Delete from tblAuditTrailLog where ID = ?")
        .Parameters.Append .CreateParameter("ID", adInteger, adParamInput, , LastIdentifier)
        .Execute , , adExecuteNoRecords
    End With
End Sub

Private Function NewCommand(ByVal CommandText As String) As ADODB.Command
    Set NewCommand = New ADODB.Command
    With NewCommand
        Set .ActiveConnection = CurrentProject.Connection
        .CommandType = adCmdText
        .CommandText = CommandText
    End With
End Function

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
