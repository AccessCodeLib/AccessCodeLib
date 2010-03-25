VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationHandler_DirTextbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Erweiterung für ApplicationHandler-Klasse: Verzeichnis mittels Textbox durchsuchen"
'---------------------------------------------------------------------------------------
' Klassenmodul: ApplicationHandler_DirTextbox (Josef Pötzl, 2009-12-18)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erweiterung für ApplicationHandler-Klasse: Verzeichnis mittels Textbox durchsuchen
' </summary>
' <remarks>
' </remarks>
'\ingroup usability
' Aktivierung im _config_Application-Modul:\n
' <code>modApplication.AddApplicationHandlerExtension New ApplicationHandler_DirTextbox</code>
'
' Aktivierung für eine Textbox in einem Formular:\n
'<code>
'With CurrentApplication!DirTextbox
'   .Textbox = Me.%TextboxSteuerelement%
'   .BaseFolderPath = %Voreingestellter Pfad%
'End With
'</code>
'
' @todo Filtermöglichkeit ergänzen, dass nur Verzeichnisse ausgewählt werden können.
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>usability/ApplicationHandler_DirTextbox.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'  <use>base/modErrorHandler.bas</use>
'  <use>file/modFiles.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Beispiel
'---------------------------------------------------------------------------------------
'/**
'\class ApplicationHandler_DirTextbox
'<h2>Verwendung</h2>
'\par Deklaration der Erweiterung im codelib-Block von _config_Application
'\verbatim
'  <use>usability/ApplicationHandler_DirTextbox.cls<use>
'\endverbatim
'\par Initialisierungsaufruf
'\code
'Public Sub InitConfig(Optional oCurrentAppHandler As ApplicationHandler = Nothing)
'   ...
'   AddApplicationHandlerExtension New ApplicationHandler_DirTextbox
'   ...
'End Sub
'\endcode
'\par Verwendung in Formular bzw. mit Textbox
'\par
'  Hilfsfunktion (Code könnte auch direkt in Ereignisbehandlung eingefügt werden)
'\verbatim
'Private Sub bindTextbox(ByRef tb As Textbox, Optional ByVal BaseFolderPath As String = vbNullString)
'
'   'Latebindung, damit ApplicationHandler_DirTextbox-Klasse nicht vorhanden sein muss
'   Dim ojbDirTextbox As Object ' ... ApplicationHandler_DirTextbox
'
'   'Standard-Instanz verwenden:
'   Set ojbDirTextbox = CurrentApplication.GetExtension("DirTextbox")
'
'   'extra Instanz für Textbox-Steuerung initialisieren ... ist nützlich,
'   'wenn die Initialisierung bei Form.Load und nicht bei Texbox.GotFocus durchgeführt werden soll
'   'Set ojbDirTextbox = CurrentApplication.GetExtensionInstance("DirTextbox")
'
'   'Textbox binden
'   If Not (ojbDirTextbox Is Nothing) Then
'      Set ojbDirTextbox.Textbox = tb
'      ojbDirTextbox.BaseFolderPath = BaseFolderPath
'   End If
'
'End Sub
'\endverbatim
'\par
'Aufruf in Ereignisbehandlung
'\verbatim
'Private Sub FileSearchTextbox1_GotFocus()
'   bindTextbox Me.FileSearchTextbox1
'End Sub
'\endverbatim
'
'**/
'---------------------------------------------------------------------------------------
'
' Hinweis:
' Klasse kann von sich selbst Instanzen erzeugen und diese "am Leben" erhalten.
'
'
Option Compare Database
Option Explicit

Private Const m_EXTENSION_KEY As String = "DirTextbox"

Private m_InstanceCollection As Collection
Private m_InstanceKey As String

Private Const EventProcedurePropertyTag As String = "[Event Procedure]"
Private WithEvents m_Textbox As Access.Textbox
Attribute m_Textbox.VB_VarHelpID = -1
Private WithEvents m_AccessForm As Access.Form 'für Überwachung, ob m_Textbox entfert wurde
Attribute m_AccessForm.VB_VarHelpID = -1

Private m_TabMode As Boolean
Private m_FolderPath As String
Private m_BaseFolderPath As String

'---------------------------------------------------------------------------------------
' Standard-Initialisierung von Erweiterungen
'---------------------------------------------------------------------------------------

Private WithEvents m_ApplicationHandler As ApplicationHandler
Attribute m_ApplicationHandler.VB_VarHelpID = -1

Public Property Set ApplicationHandlerRef(ByRef objRef As ApplicationHandler)
On Error GoTo HandleErr

   Set m_ApplicationHandler = objRef

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandlerRef", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
End Property

Public Property Get ExtensionKey() As String
On Error Resume Next
   ExtensionKey = m_EXTENSION_KEY
End Property

Private Sub m_AccessForm_Unload(Cancel As Integer)
On Error Resume Next
   'Referenzen entfernen
   Set m_Textbox = Nothing
   Set m_AccessForm = Nothing
   m_InstanceKey = vbNullString 'damit wird beim nächsten Durchlauf der Collection aufgeräumt
End Sub

'---------------------------------------------------------------------------------------
' Standard-Ereignisbehandlung von Erweiterungen
'---------------------------------------------------------------------------------------

' CheckExtension
Private Sub m_ApplicationHandler_CheckExtension(ByVal Key As String, ByRef exists As Boolean)
On Error Resume Next
   If Key = m_EXTENSION_KEY Then exists = True
End Sub

' ExtensionLookup
Private Sub m_ApplicationHandler_ExtensionLookup(ByVal Key As String, ByRef ExtensionReference As Object)

On Error GoTo HandleErr

   If Key = m_EXTENSION_KEY Then
      Set ExtensionReference = Me
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ExtAppFile.m_ApplicationHandler_ExtensionLookup", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

'ExtensionNewInstance
Private Sub m_ApplicationHandler_ExtensionInstance(ByVal sExtensionKey As String, ByVal sInstanceKey As String, ByRef ExtensionReference As Object)
On Error GoTo HandleErr

   If sExtensionKey = m_EXTENSION_KEY Then
      getInstance sInstanceKey, ExtensionReference
   End If

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "m_ApplicationHandler_ExtensionInstance", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
End Sub

Private Sub getInstance(ByVal sInstanceKey As String, ByRef ExtensionReference As Object)
   
   Dim instanceRef As ApplicationHandler_DirTextbox
   Dim tempRef As ApplicationHandler_DirTextbox
   
   Dim col As Collection
   Dim i As Long
   
On Error GoTo HandleErr

   Set col = InstanceCollection
   
   If Len(sInstanceKey) = 0 Then
      sInstanceKey = m_EXTENSION_KEY & "." & ObjPtr(Me)
   End If
   
   'Nach vorhandenem InstanceKey suchen und gleichzeitig aufräumen
   For i = col.Count To 1 Step -1
      Set tempRef = col(i)
      If tempRef.Textbox Is Nothing And Len(tempRef.InstanceKey) = 0 Then ' Col-Eintrag kann entfernt werden
         col.Remove i
      ElseIf tempRef.InstanceKey = sInstanceKey Then
         Set instanceRef = tempRef
         Exit For
      End If
   Next
   
   'Bei Bedarf neu anlegen
   If instanceRef Is Nothing Then
      Set instanceRef = New ApplicationHandler_DirTextbox
      instanceRef.InstanceKey = sInstanceKey
      col.Add instanceRef, sInstanceKey
   End If
   
   Set ExtensionReference = instanceRef

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "getInstance", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Property Get InstanceCollection() As Collection

On Error GoTo HandleErr

   If m_InstanceCollection Is Nothing Then
      Set m_InstanceCollection = New Collection
   End If
   Set InstanceCollection = m_InstanceCollection

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "InstanceCollection", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Friend Property Get InstanceKey() As String
On Error Resume Next
   InstanceKey = m_InstanceKey
End Property

Friend Property Let InstanceKey(ByVal sKey As String)
On Error Resume Next
   m_InstanceKey = sKey
End Property


'ExtensionPropertyLookup
Private Sub m_ApplicationHandler_ExtensionPropertyLookup( _
                     ByVal Key As String, ByVal PropertyName As String, _
                     ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
                     
On Error GoTo HandleErr

   If Key = m_EXTENSION_KEY Then
      getExtensionPropertyLookup PropertyName, ResumeMode, ResumeMessage
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ExtAppFile.m_ApplicationHandler_ExtensionPropertyLookup", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

' AfterDispose
Private Sub m_ApplicationHandler_AfterDispose(ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
'=> Referenz in m_ApplicationHandler auf Nothing setzen
On Error Resume Next
   Set m_ApplicationHandler = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Ergänzungen für Ereiterung: ApplicationHandler_DirTextbox
'---------------------------------------------------------------------------------------
'
Private Sub getExtensionPropertyLookup(ByVal PropertyName As String, ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
   
On Error GoTo HandleErr

   ResumeMode = AppResumeMode_Completed
   
   Select Case PropertyName
      'case ...
          
      Case Else 'Property wurde nicht erkannt
         ResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Error
                 
   End Select

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ExtDirTextbox.getExtensionPropertyLookup", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

'---------------------------------------------------------------------------------------
' Property: BaseFolderPath (Josef Pötzl, 2009-12-18)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ausgangsverzeichnis (wird vor den Text in der TextBox angehängt)
' </summary>
' <remarks>
' String endet immer mit einem "\", auch wenn bei der Einstellung kein abschließender "\"
' eingegeben wurde.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get BaseFolderPath() As String
On Error Resume Next
   BaseFolderPath = m_BaseFolderPath
End Property

'---------------------------------------------------------------------------------------
' Property: BaseFolderPath (Josef Pötzl, 2009-12-18)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ausgangsverzeichnis (wird vor den Text in der TextBox angehängt)
' </summary>
' <remarks>
'
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let BaseFolderPath(ByVal newPath As String)
On Error Resume Next
   m_BaseFolderPath = newPath
   If Len(m_BaseFolderPath) > 0 Then
      If Right$(m_BaseFolderPath, 1) <> "\" Then
         m_BaseFolderPath = m_BaseFolderPath & "\"
      End If
   End If
End Property

Private Sub initNewFileSearch()
'Zwischenspeicher löschen
On Error Resume Next
   m_TabMode = False
   m_FolderPath = vbNullString
End Sub

'---------------------------------------------------------------------------------------
' Property: Textbox (Josef Pötzl, 2009-12-18)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Die Textbox in der der Pfad eingegeben wird
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get Textbox() As Access.Textbox

On Error GoTo HandleErr

   Set Textbox = m_Textbox

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "Textbox", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

'---------------------------------------------------------------------------------------
' Property: Textbox (Josef Pötzl, 2009-12-18)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Die Textbox in der der Pfad eingegeben wird
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Set Textbox(ByVal ref As Access.Textbox)
   
   Dim parentObject As Object
   
On Error GoTo HandleErr

   Set m_Textbox = Nothing
   
   Set m_Textbox = ref
   With m_Textbox
      Set parentObject = m_Textbox.Parent
      If Not (TypeOf parentObject Is Access.Form) Then
         'Falls objParent kein Formular, sondern eine Registerseite zurückgibt
         If TypeOf parentObject Is Access.Page Then
            Set parentObject = parentObject.Parent
         End If
      End If
      If (TypeOf parentObject Is Access.Form) Then
         Set m_AccessForm = parentObject
         If Len(m_AccessForm.OnUnload) = 0 Then
            m_AccessForm.OnUnload = EventProcedurePropertyTag
         End If
      End If
      If Len(.OnKeyDown) = 0 Then
         .OnKeyDown = EventProcedurePropertyTag
      End If
      If Len(.BeforeUpdate) = 0 Then
         .BeforeUpdate = EventProcedurePropertyTag
      End If
   End With
   
   initNewFileSearch

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "Textbox", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Private Sub m_Textbox_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
' Auf Tastendruck in der Textbox reagieren

'/**
' * @todo Code aufräumen
'**/

   Static searchString As String
   Dim matchString As String
   Dim baseFolder As String
   Dim textBoxString As String
   Dim pos As Long

On Error GoTo HandleErr

   baseFolder = BaseFolderPath
   textBoxString = m_Textbox.Text
   
   ' "\" abfangen und prüfen, ob der Einsatz sinnvoll/möglich ist
   If Shift > 0 And KeyCode = 219 Then
      searchString = GetFullPathFromRelativPath(textBoxString, baseFolder)
      If Right$(searchString, 1) = ":" Then
         Exit Sub
      End If
      If InStr(1, searchString, "*") = 0 Then
         If Len(Dir$(searchString)) > 0 Then 'Das kann kein Verzeichnis sein, da bereits Datei erreicht ist
            KeyCode = vbKeyEnd
         ElseIf Len(Dir$(searchString, vbDirectory)) = 0 Then
            KeyCode = vbKeyEnd
         End If
      End If
      Exit Sub
   End If
   
   If KeyCode = vbKeyReturn Then
      searchString = Replace$(GetFullPathFromRelativPath(textBoxString, baseFolder), "/", "\")
      If Right$(searchString, 3) = "\.." Then
         m_FolderPath = GetDirFromFilePath(Left$(searchString, Len(searchString) - 3))
      Else
         'hier Event auslösen, dass Auswahl bestätigt wurde
         Exit Sub
      End If
      
      If Len(m_FolderPath) > 0 Then
         m_FolderPath = m_FolderPath & ".."
      End If
      If Len(baseFolder) > 0 Then
         m_FolderPath = GetRelativPathFromFullPath(m_FolderPath, baseFolder)
      End If
      
      m_Textbox.Text = m_FolderPath
      KeyCode = vbKeyEnd
      
   ElseIf KeyCode = vbKeyTab And Shift = 0 Then
      If m_TabMode Then
      
         matchString = Dir$(, vbDirectory)
         Do While Left$(matchString, 1) = "." ' eins weiter
            matchString = Dir$(, vbDirectory)
         Loop
         If Len(matchString) = 0 Then 'wieder von vorne anfangen
            matchString = Dir$(searchString, vbDirectory)
            Do While matchString = "." ' eins weiter
               matchString = Dir$(, vbDirectory)
            Loop
         End If
         m_Textbox.SetFocus
         m_Textbox.SelStart = Len(m_FolderPath)
         If Right$(m_FolderPath, 2) = ".." Then
            m_FolderPath = Left$(m_FolderPath, Len(m_FolderPath) - 2)
         End If
         m_Textbox.Text = m_FolderPath & matchString
         KeyCode = vbKeyEnd
      Else
         searchString = Replace(GetFullPathFromRelativPath(textBoxString, baseFolder), "/", "\")
         If Len(searchString) = 2 And Right$(searchString, 1) = ":" Then
            searchString = searchString & "\"
         End If
         If Right$(searchString, 1) = "\" Then
            searchString = searchString & "*"
         End If
         
         matchString = Dir$(searchString, vbDirectory)
         If Len(matchString) = 0 Then ' "*" anhängen
            If Right$(searchString, 1) <> "*" Then
               searchString = searchString & "*"
            End If
            matchString = Dir$(searchString, vbDirectory)
         End If
         
         If Len(matchString) > 0 Then
         
            If matchString = "." Then 'Verzeichnis, Suchstring hat mit "\" geendet
               m_FolderPath = textBoxString
               matchString = ""
            Else
            
               m_FolderPath = searchString
               pos = InStr(1, m_FolderPath, "*")
               If pos > 0 Then
                  m_FolderPath = Left$(m_FolderPath, pos - 1)
               End If
               '
               If Len(baseFolder) > 0 Then
                  m_FolderPath = GetRelativPathFromFullPath(m_FolderPath, baseFolder)
               End If
               m_FolderPath = GetDirFromFilePath(m_FolderPath)
            End If
            
            matchString = m_FolderPath & matchString
            
            m_Textbox.Text = matchString
            
         End If
         m_TabMode = True
         KeyCode = vbKeyEnd
      End If
     
   Else
      m_TabMode = False
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "m_Textbox_KeyDown", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      initNewFileSearch
      Resume ExitHere
   End Select
   
End Sub
