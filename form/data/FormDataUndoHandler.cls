VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormDataUndoHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: FormDataUndoHandler (Josef Pötzl, 2010-04-15)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erweiterete Undo-Funktionalität für Formulare
' </summary>
' <remarks>
' Mit dieser Klasse kann Undo auch nach dem Speichern des Datensatzes
' (z. B. nach dem Wechsel vom HF in ein UF)durchgeführt werden. \n\n
'
' Anwendung: \n
' Nach der Initialisierung ist der FormDataUndoHandler-Instanz die Referenz des zu überwachenden
' Formulars über die Eigenschaft FormDataUndoHandler.Form zu übergeben. \n
'
' Beispiel: \n
' <code>Set m_MainFormUndoHandler = New FormDataUndoHandler
' Set m_MainFormUndoHandler.Form = Me</code>
'
' Wichtig: Bei Zirkelverweis (.Form = Me) die FormDataUndoHandler-Referenz als Reaktion
'          auf Form.Unload auf Nothing setzen.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/data/FormDataUndoHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit
'/**
' * @todo Nutzung mit Adodb-Rs ermöglichen
'**/

'---------------------------------------------------------------------------------------
' Event: Dirty
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Weitergabe des Dirty-Ereignisses aus dem Formular
' </summary>
' <param name="Cancel">Dirty abbrechen (entspricht Cancel aus Form.Dirty)</param>
' <remarks></remarks>
'**/
Public Event Dirty(Cancel As Integer)

'---------------------------------------------------------------------------------------
' Event: DataSavedOrRestored
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Meldet, dass der Datensatz gespeichert wurde bzw. die Originaldaten erhalten hat.
' </summary>
' <param name="UndoDataExists">True ... Undo-Daten sind verfügbar</param>
' <remarks></remarks>
'**/
Public Event DataSavedOrRestored(ByVal UndoDataExists As Boolean)
   'UndoDataExists ... gibt an, ob Undo-Daten verfügbar sind

Private WithEvents m_Form As Form
Attribute m_Form.VB_VarHelpID = -1

Private Const g_conEventProcedurePropertyTag As String = "[Event Procedure]"

Private m_UndoDataCollection As Collection 'Hier werden die Daten des Formular-Datensatzes zwischengespeichert
Private m_FormDataNewRecordset As Boolean

Private m_Bookmark As Variant

Private Sub Class_Initialize()
   Set m_UndoDataCollection = New Collection
End Sub

Private Sub Class_Terminate()
   Set m_Form = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Property: Form
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Form-Referenz des überwachten Access-Formulars
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get Form() As Form
   Set Form = m_Form
End Property

'---------------------------------------------------------------------------------------
' Property: Form
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Form-Referenz des zu überwachtenden Access-Formulars
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Set Form(newRef As Form)
   Set m_Form = newRef
   initForm
End Property

Private Sub initForm()

   If m_Form Is Nothing Then
      Exit Sub
   End If

   RaiseEvent DataSavedOrRestored(UndoDataExists)

   If Len(m_Form.AfterUpdate) = 0 Then
      m_Form.AfterUpdate = g_conEventProcedurePropertyTag
   End If
   If Len(m_Form.OnCurrent) = 0 Then
      m_Form.OnCurrent = g_conEventProcedurePropertyTag
   End If
   If Len(m_Form.OnDirty) = 0 Then
      m_Form.OnDirty = g_conEventProcedurePropertyTag
   End If

End Sub

Private Sub m_Form_AfterUpdate()
   RaiseEvent DataSavedOrRestored(UndoDataExists)
End Sub

Private Sub m_Form_Current()

   Dim rst As DAO.Recordset
   
   Set m_UndoDataCollection = New Collection
   RaiseEvent DataSavedOrRestored(UndoDataExists)
   
End Sub

' Dirty weitergeben
Private Sub m_Form_Dirty(Cancel As Integer)

   Dim rst As DAO.Recordset
   Dim recordsetDataArray As Variant
   Dim tempArray(0, 0) As Variant
   
   RaiseEvent Dirty(Cancel)
   If Not Cancel Then
      'Daten zwischenspeichern
      If m_Form.NewRecord Then
         m_FormDataNewRecordset = True
         Set m_UndoDataCollection = New Collection 'mit neuer Collection starten
         tempArray(0, 0) = Empty
         m_UndoDataCollection.Add tempArray
      Else
         m_FormDataNewRecordset = False
         
         Set rst = m_Form.RecordsetClone
         rst.Bookmark = m_Form.Bookmark 'sync
         recordsetDataArray = rst.GetRows(1)

         m_UndoDataCollection.Add recordsetDataArray

      End If
   End If
   
End Sub

'---------------------------------------------------------------------------------------
' Property: UndoDataExists
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt an, ob Undo-Daten verfügbar sind.
' </summary>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get UndoDataExists() As Boolean
   UndoDataExists = (m_UndoDataCollection.Count > 0) Or m_FormDataNewRecordset
End Property


'---------------------------------------------------------------------------------------
' Sub: Undo
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Undo-Aktion durchführen
' </summary>
' <remarks>
' Löst am Ende FormDataUndoHandler.DataSavedOrRestored(UndoDataExists) aus.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub Undo()
'/**
' * @todo passt das alles?
'**/

   Dim rst As DAO.Recordset
   Dim fld As DAO.Field
   Dim i As Long
   
   
   Dim bolUpdate As Boolean
   Dim recordsetDataArray As Variant
   
   If m_UndoDataCollection.Count = 0 Then Exit Sub 'es gibt keine Daten zum Zurückschreiben
   
   
   recordsetDataArray = UndoDataArray
   
'   If m_Form.Dirty Then 'Form.Undo ist noch möglich
'      m_Form.Undo 'nicht verwenden, da das sonst Current im UF auslöst
'   Else
   If IsEmpty(recordsetDataArray(0, 0)) Then 'DS löschen, da neu er neu angelegt wurde
      m_Form.Recordset.Delete
      m_Form.Requery
      m_Form.Recordset.Bookmark = m_Form.Bookmark
      m_FormDataNewRecordset = False
   Else ' Daten aus Array zurückschreiben
      Set rst = m_Form.Recordset
      i = 0
      
      For Each fld In rst.Fields
         If fld.Value = recordsetDataArray(i, 0) Then
            'keine Änderung notwendig (NULL-Vergleich schlägt fehlt, will ich aber nicht noch extra prüfen)
         ElseIf fld.DataUpdatable Then
            If Not bolUpdate Then
               rst.Edit
               bolUpdate = True
            End If
            fld.Value = recordsetDataArray(i, 0)
         End If
         i = i + 1
      Next

      If bolUpdate Then
         rst.Update 'reicht da für Formular-Aktualisierung aus?
         m_Form.Refresh
      End If

   End If
   
   If m_UndoDataCollection.Count > 0 Then
      m_UndoDataCollection.Remove m_UndoDataCollection.Count
   End If
   
   RaiseEvent DataSavedOrRestored(UndoDataExists)
   
End Sub

Private Property Get UndoDataArray() As Variant
   UndoDataArray = m_UndoDataCollection.Item(m_UndoDataCollection.Count)
End Property
