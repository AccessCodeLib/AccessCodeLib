VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAuditTrail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Ursprüngliches Copyright:
' ================================================
' Code by Martin Green Email: martin@fontstuff.com
' Visit my Office Tips website @ www.fontstuff.com
' YouTube tutorials www.youtube.com/martingreenvba
' ================================================

Private WithEvents m_frm As Form

Private m_Identifier As Long
Private m_IdFieldName As String

Private cnn As ADODB.Connection
Private rst As ADODB.Recordset
Private datTimeCheck As Date
Private strUserID As String

Public Property Set FormObj(ByRef FRM_ As Access.Form)
    Set m_frm = FRM_
    m_IdFieldName = getIDField(FRM_)
    m_frm.BeforeUpdate = "[Event Procedure]"
    m_frm.AfterUpdate = "[Event Procedure]"
    m_frm.OnDelete = "[Event Procedure]"
    m_frm.AfterDelConfirm = "[Event Procedure]"
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
    Call AuditChanges("DELETE")
End Sub

Private Sub m_frm_AfterDelConfirm(Status As Integer)
    If Status <> acDeleteOK Then Call AuditRedoDelete
End Sub

Private Sub m_frm_BeforeUpdate(Cancel As Integer)
    If m_frm.NewRecord Then
        Call AuditChanges("NEW")
    Else
        Call AuditChanges("EDIT")
    End If
End Sub

Private Sub AuditChanges(UserAction As String)
    Dim CTL As Control

    Set cnn = CurrentProject.Connection
    Set rst = New ADODB.Recordset
    rst.Open "SELECT * FROM tblAuditTrail", cnn, adOpenDynamic, adLockOptimistic
    datTimeCheck = Now()
    strUserID = Environ("USERNAME")
    Select Case UserAction
    Case "EDIT"
        For Each CTL In m_frm.Controls
            If CTL.Tag = "Audit" Then
                If Nz(CTL.Value) <> Nz(CTL.OldValue) Then
                    With rst
                        .AddNew
                        ![DateTime] = datTimeCheck
                        ![UserName] = strUserID
                        ![formname] = m_frm.Name
                        ![action] = UserAction
                        ![RecordID] = m_frm.Controls(m_IdFieldName).Value
                        ![FieldName] = CTL.ControlSource
                        ![OldValue] = CTL.OldValue
                        ![NewValue] = CTL.Value
                        .Update
                    End With
                End If
            End If
        Next CTL
    Case Else
        With rst
            .AddNew
            ![DateTime] = datTimeCheck
            ![UserName] = strUserID
            ![formname] = m_frm.Name
            ![action] = UserAction
            ![RecordID] = m_frm.Recordset.Fields(m_IdFieldName).Value
            .Update
            If UserAction = "DELETE" Then lastIdentifier = ![AuditTrailID]
        End With
    End Select
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
End Sub

Private Sub AuditRedoDelete()
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset

    Set cnn = CurrentProject.Connection
    Set rst = New ADODB.Recordset
    rst.Open "SELECT * FROM tblAuditTrail", cnn, adOpenDynamic, adLockOptimistic
    rst.Find "AuditTrailID = " & lastIdentifier
    If Not rst.EOF Then
        rst.Delete
    End If

    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
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
