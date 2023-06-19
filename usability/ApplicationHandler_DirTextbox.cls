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
' Class: usability.ApplicationHandler_DirTextbox
'---------------------------------------------------------------------------------------
'
' Extension for ApplicationHandler class: Search directory using textbox
'
' Author:
'     Josef Poetzl
'
' Note:
'     Class can create instances of itself and keep them "alive".
'
' Remarks:
' Activation in the _config_Application module:
' <code>
'   Set m_Extensions = New ApplicationHandler_ExtensionCollection
'   ...
'   m_Extensions.Add New ApplicationHandler_DirTextbox
'   ...
' </code>
'
' Activation for a text box in a form:
' <code>
' With CurrentApplication!DirTextbox
'   .Textbox = Me.%TextboxSteuerelement%
'   .BaseFolderPath = %Voreingestellter Pfad%
' End With
' </code>
'
' @todo Add filter option so that only directories can be selected.
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>usability/ApplicationHandler_DirTextbox.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'  <use>file/FileTools.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Example
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
Option Compare Database
Option Explicit

Private Const EXTENSION_KEY As String = "DirTextbox"

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

Public Property Set ApplicationHandlerRef(ByRef ObjRef As ApplicationHandler)
   Set m_ApplicationHandler = ObjRef
End Property

Public Property Get ExtensionKey() As String
   ExtensionKey = EXTENSION_KEY
End Property

Private Sub m_AccessForm_Unload(ByRef Cancel As Integer)
On Error Resume Next
   Set m_Textbox = Nothing
   Set m_AccessForm = Nothing
   m_InstanceKey = vbNullString 'this will be used to clean up the next time the collection is run
End Sub

'---------------------------------------------------------------------------------------
' Standard event handling of extensions
'---------------------------------------------------------------------------------------

' CheckExtension
Private Sub m_ApplicationHandler_CheckExtension(ByVal Key As String, ByRef Exists As Boolean)
   If Key = EXTENSION_KEY Then Exists = True
End Sub

' ExtensionLookup
Private Sub m_ApplicationHandler_ExtensionLookup(ByVal Key As String, ByRef ExtensionReference As Object)
   If Key = EXTENSION_KEY Then
      Set ExtensionReference = Me
   End If
End Sub

'ExtensionNewInstance
Private Sub m_ApplicationHandler_ExtensionInstance(ByVal ExtensionKeyToCheck As String, ByVal InstanceKey As String, ByRef ExtensionReference As Object)
   If ExtensionKeyToCheck = EXTENSION_KEY Then
      GetInstance InstanceKey, ExtensionReference
   End If
End Sub

Private Sub GetInstance(ByVal Key As String, ByRef ExtensionReference As Object)
   
   Dim InstanceRef As ApplicationHandler_DirTextbox
   Dim TempRef As ApplicationHandler_DirTextbox
   Dim col As Collection
   Dim i As Long

   Set col = InstanceCollection
   
   If Len(Key) = 0 Then
      Key = EXTENSION_KEY & "." & ObjPtr(Me)
   End If
   
   'Search for existing InstanceKey and clean up at the same time
   For i = col.Count To 1 Step -1
      Set TempRef = col(i)
      If TempRef.Textbox Is Nothing And Len(TempRef.InstanceKey) = 0 Then ' Col-Eintrag kann entfernt werden
         col.Remove i
      ElseIf TempRef.InstanceKey = Key Then
         Set InstanceRef = TempRef
         Exit For
      End If
   Next
   
   'Create new if required
   If InstanceRef Is Nothing Then
      Set InstanceRef = New ApplicationHandler_DirTextbox
      InstanceRef.InstanceKey = Key
      col.Add InstanceRef, Key
   End If
   
   Set ExtensionReference = InstanceRef

End Sub

Private Property Get InstanceCollection() As Collection
   If m_InstanceCollection Is Nothing Then
      Set m_InstanceCollection = New Collection
   End If
   Set InstanceCollection = m_InstanceCollection
End Property

Friend Property Get InstanceKey() As String
   InstanceKey = m_InstanceKey
End Property

Friend Property Let InstanceKey(ByVal Key As String)
   m_InstanceKey = Key
End Property


'ExtensionPropertyLookup
Private Sub m_ApplicationHandler_ExtensionPropertyLookup( _
                     ByVal Key As String, ByVal PropertyName As String, _
                     ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
   If Key = EXTENSION_KEY Then
      GetExtensionPropertyLookup PropertyName, ResumeMode, ResumeMessage
   End If
End Sub

' AfterDispose
Private Sub m_ApplicationHandler_AfterDispose(ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
   Set m_ApplicationHandler = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Additions for extension: ApplicationHandler_DirTextbox
'---------------------------------------------------------------------------------------
'
Private Sub GetExtensionPropertyLookup(ByVal PropertyName As String, ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

   ResumeMode = AppResumeMode_Completed
   Select Case PropertyName
      'case ...
          
      Case Else 'Property was not recognized
         ResumeMessage = "Property '" & PropertyName & "' is not supported"
   End Select

End Sub

'---------------------------------------------------------------------------------------
' Property: BaseFolderPath
'---------------------------------------------------------------------------------------
'
' Ausgangsverzeichnis (wird vor den Text in der TextBox angehängt)
'
' Remarks:
'     String always ends with a "\", even if no final "\" was entered in the setting.
'
'---------------------------------------------------------------------------------------
Public Property Get BaseFolderPath() As String
   BaseFolderPath = m_BaseFolderPath
End Property

'---------------------------------------------------------------------------------------
' Property: BaseFolderPath
'---------------------------------------------------------------------------------------
'
' Output directory (will be appended in front of the text in the TextBox)
'
'---------------------------------------------------------------------------------------
Public Property Let BaseFolderPath(ByVal NewPath As String)
   m_BaseFolderPath = NewPath
   If Len(m_BaseFolderPath) > 0 Then
      If Right$(m_BaseFolderPath, 1) <> "\" Then
         m_BaseFolderPath = m_BaseFolderPath & "\"
      End If
   End If
End Property

Private Sub InitNewFileSearch()
'Clear cache
   m_TabMode = False
   m_FolderPath = vbNullString
End Sub

'---------------------------------------------------------------------------------------
' Property: Textbox
'---------------------------------------------------------------------------------------
'
' The text box in which the path is entered
'
'---------------------------------------------------------------------------------------
Public Property Get Textbox() As Access.Textbox
   Set Textbox = m_Textbox
End Property

Public Property Set Textbox(ByVal NewRef As Access.Textbox)
   
   Dim parentObject As Object

   Set m_Textbox = Nothing
   
   Set m_Textbox = NewRef
   With m_Textbox
      Set parentObject = m_Textbox.Parent
      If Not (TypeOf parentObject Is Access.Form) Then
         'If objParent does not return a form but a tab page
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
   
   InitNewFileSearch

End Property

Private Sub m_Textbox_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
' Respond to keystrokes in the text box

'/**
' * @todo Clean up code
'**/

   Static SearchString As String
   Dim MatchString As String
   Dim BaseFolder As String
   Dim TextBoxString As String
   Dim Pos As Long

   BaseFolder = BaseFolderPath
   TextBoxString = m_Textbox.Text
   
   ' Intercept "\" and check if the use is reasonable/possible
   If Shift > 0 And KeyCode = 219 Then
      SearchString = GetFullPathFromRelativPath(TextBoxString, BaseFolder)
      If Right$(SearchString, 1) = ":" Then
         Exit Sub
      End If
      If InStr(1, SearchString, "*") = 0 Then
         If Len(Dir$(SearchString)) > 0 Then 'Das kann kein Verzeichnis sein, da bereits Datei erreicht ist
            KeyCode = vbKeyEnd
         ElseIf Len(Dir$(SearchString, vbDirectory)) = 0 Then
            KeyCode = vbKeyEnd
         End If
      End If
      Exit Sub
   End If
   
   If KeyCode = vbKeyReturn Then
      SearchString = Replace$(GetFullPathFromRelativPath(TextBoxString, BaseFolder), "/", "\")
      If Right$(SearchString, 3) = "\.." Then
         m_FolderPath = FileTools.GetDirFromFullFileName(Left$(SearchString, Len(SearchString) - 3))
      Else
         'Trigger event here that selection was confirmed
         Exit Sub
      End If
      
      If Len(m_FolderPath) > 0 Then
         m_FolderPath = m_FolderPath & ".."
      End If
      If Len(BaseFolder) > 0 Then
         m_FolderPath = GetRelativPathFromFullPath(m_FolderPath, BaseFolder)
      End If
      
      m_Textbox.Text = m_FolderPath
      KeyCode = vbKeyEnd
      
   ElseIf KeyCode = vbKeyTab And Shift = 0 Then
      If m_TabMode Then
      
         MatchString = Dir$(, vbDirectory)
         Do While Left$(MatchString, 1) = "." ' one more
            MatchString = Dir$(, vbDirectory)
         Loop
         If Len(MatchString) = 0 Then 'start over again
            MatchString = Dir$(SearchString, vbDirectory)
            Do While MatchString = "." ' one more
               MatchString = Dir$(, vbDirectory)
            Loop
         End If
         m_Textbox.SetFocus
         m_Textbox.SelStart = Len(m_FolderPath)
         If Right$(m_FolderPath, 2) = ".." Then
            m_FolderPath = Left$(m_FolderPath, Len(m_FolderPath) - 2)
         End If
         m_Textbox.Text = m_FolderPath & MatchString
         KeyCode = vbKeyEnd
      Else
         SearchString = Replace(GetFullPathFromRelativPath(TextBoxString, BaseFolder), "/", "\")
         If Len(SearchString) = 2 And Right$(SearchString, 1) = ":" Then
            SearchString = SearchString & "\"
         End If
         If Right$(SearchString, 1) = "\" Then
            SearchString = SearchString & "*"
         End If
         
         MatchString = Dir$(SearchString, vbDirectory)
         If Len(MatchString) = 0 Then ' append "*"
            If Right$(SearchString, 1) <> "*" Then
               SearchString = SearchString & "*"
            End If
            MatchString = Dir$(SearchString, vbDirectory)
         End If
         
         If Len(MatchString) > 0 Then
         
            If MatchString = "." Then 'Directory, search string ended with "\".
               m_FolderPath = TextBoxString
               MatchString = ""
            Else
            
               m_FolderPath = SearchString
               Pos = InStr(1, m_FolderPath, "*")
               If Pos > 0 Then
                  m_FolderPath = Left$(m_FolderPath, Pos - 1)
               End If
               '
               If Len(BaseFolder) > 0 Then
                  m_FolderPath = GetRelativPathFromFullPath(m_FolderPath, BaseFolder)
               End If
               m_FolderPath = FileTools.GetDirFromFullFileName(m_FolderPath)
            End If
            
            MatchString = m_FolderPath & MatchString
            
            m_Textbox.Text = MatchString
            
         End If
         m_TabMode = True
         KeyCode = vbKeyEnd
      End If
     
   Else
      m_TabMode = False
   End If

End Sub
