VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationHandler_DirTextbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Erweiterung f?r ApplicationHandler-Klasse: Verzeichnis mittels Textbox durchsuchen"
'---------------------------------------------------------------------------------------
' Klassenmodul: ApplicationHandler_DirTextbox (Josef P?tzl, 2009-12-18)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erweiterung f?r ApplicationHandler-Klasse: Verzeichnis mittels Textbox durchsuchen
' </summary>
' <remarks>
' </remarks>
'\ingroup usability
' Aktivierung im _config_Application-Modul:\n
' <code>modApplication.AddApplicationHandlerExtension New ApplicationHandler_DirTextbox</code>
'
' Aktivierung f?r eine Textbox in einem Formular:\n
'<code>
'With CurrentApplication!DirTextbox
'   .Textbox = Me.%TextboxSteuerelement%
'   .BaseFolderPath = %Voreingestellter Pfad%
'End With
'</code>
'
' @todo Filterm?glichkeit erg?nzen, dass nur Verzeichnisse ausgew?hlt werden k?nnen.
'**/
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
'  Hilfsfunktion (Code k?nnte auch direkt in Ereignisbehandlung eingef?gt werden)
'\verbatim
'Private Sub bindTextbox(ByRef tb As Textbox, Optional ByVal BaseFolderPath As String = vbNullString)
'
'   'Latebindung, damit ApplicationHandler_DirTextbox-Klasse nicht vorhanden sein muss
'   Dim ojbDirTextbox As Object ' ... ApplicationHandler_DirTextbox
'
'   'Standard-Instanz verwenden:
'   Set ojbDirTextbox = CurrentApplication.GetExtension("DirTextbox")
'
'   'extra Instanz f?r Textbox-Steuerung initialisieren ... ist n?tzlich,
'   'wenn die Initialisierung bei Form.Load und nicht bei Texbox.GotFocus durchgef?hrt werden soll
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

Private Const EXTENSION_KEY As String = "DirTextbox"

Private m_InstanceCollection As Collection
Private m_InstanceKey As String

Private Const EventProcedurePropertyTag As String = "[Event Procedure]"
Private WithEvents m_Textbox As Access.Textbox
Attribute m_Textbox.VB_VarHelpID = -1
Private WithEvents m_AccessForm As Access.Form 'f?r ?berwachung, ob m_Textbox entfert wurde
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
   'Referenzen entfernen
   Set m_Textbox = Nothing
   Set m_AccessForm = Nothing
   m_InstanceKey = vbNullString 'damit wird beim n?chsten Durchlauf der Collection aufger?umt
End Sub

'---------------------------------------------------------------------------------------
' Standard-Ereignisbehandlung von Erweiterungen
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
   
   'Nach vorhandenem InstanceKey suchen und gleichzeitig aufr?umen
   For i = col.Count To 1 Step -1
      Set TempRef = col(i)
      If TempRef.Textbox Is Nothing And Len(TempRef.InstanceKey) = 0 Then ' Col-Eintrag kann entfernt werden
         col.Remove i
      ElseIf TempRef.InstanceKey = Key Then
         Set InstanceRef = TempRef
         Exit For
      End If
   Next
   
   'Bei Bedarf neu anlegen
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
'=> Referenz in m_ApplicationHandler auf Nothing setzen
   Set m_ApplicationHandler = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Erg?nzungen f?r Ereiterung: ApplicationHandler_DirTextbox
'---------------------------------------------------------------------------------------
'
Private Sub GetExtensionPropertyLookup(ByVal PropertyName As String, ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

   ResumeMode = AppResumeMode_Completed
   Select Case PropertyName
      'case ...
          
      Case Else 'Property wurde nicht erkannt
         ResumeMessage = "Property '" & PropertyName & "' is not supported"
   End Select

End Sub

'---------------------------------------------------------------------------------------
' Property: BaseFolderPath (Josef P?tzl, 2009-12-18)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ausgangsverzeichnis (wird vor den Text in der TextBox angeh?ngt)
' </summary>
' <remarks>
' String endet immer mit einem "\", auch wenn bei der Einstellung kein abschlie?ender "\"
' eingegeben wurde.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get BaseFolderPath() As String
   BaseFolderPath = m_BaseFolderPath
End Property

'---------------------------------------------------------------------------------------
' Property: BaseFolderPath (Josef P?tzl, 2009-12-18)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ausgangsverzeichnis (wird vor den Text in der TextBox angeh?ngt)
' </summary>
' <remarks>
'
' </remarks>
'**/
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
'Zwischenspeicher l?schen
   m_TabMode = False
   m_FolderPath = vbNullString
End Sub

'---------------------------------------------------------------------------------------
' Property: Textbox (Josef P?tzl, 2009-12-18)
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
   Set Textbox = m_Textbox
End Property

'---------------------------------------------------------------------------------------
' Property: Textbox (Josef P?tzl, 2009-12-18)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Die Textbox in der der Pfad eingegeben wird
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Set Textbox(ByVal NewRef As Access.Textbox)
   
   Dim parentObject As Object

   Set m_Textbox = Nothing
   
   Set m_Textbox = NewRef
   With m_Textbox
      Set parentObject = m_Textbox.Parent
      If Not (TypeOf parentObject Is Access.Form) Then
         'Falls objParent kein Formular, sondern eine Registerseite zur?ckgibt
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
' Auf Tastendruck in der Textbox reagieren

'/**
' * @todo Code aufr?umen
'**/

   Static SearchString As String
   Dim MatchString As String
   Dim BaseFolder As String
   Dim TextBoxString As String
   Dim Pos As Long

   BaseFolder = BaseFolderPath
   TextBoxString = m_Textbox.Text
   
   ' "\" abfangen und pr?fen, ob der Einsatz sinnvoll/m?glich ist
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
         'hier Event ausl?sen, dass Auswahl best?tigt wurde
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
         Do While Left$(MatchString, 1) = "." ' eins weiter
            MatchString = Dir$(, vbDirectory)
         Loop
         If Len(MatchString) = 0 Then 'wieder von vorne anfangen
            MatchString = Dir$(SearchString, vbDirectory)
            Do While MatchString = "." ' eins weiter
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
         If Len(MatchString) = 0 Then ' "*" anh?ngen
            If Right$(SearchString, 1) <> "*" Then
               SearchString = SearchString & "*"
            End If
            MatchString = Dir$(SearchString, vbDirectory)
         End If
         
         If Len(MatchString) > 0 Then
         
            If MatchString = "." Then 'Verzeichnis, Suchstring hat mit "\" geendet
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
