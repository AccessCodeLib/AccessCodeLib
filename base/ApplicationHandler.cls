VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Hauptsteuerung der Anwendung"
'---------------------------------------------------------------------------------------
' Klassenmodul: ApplicationHandler (Josef P�tzl, 2009-12-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hauptsteuerung der Anwendung
' </summary>
' <remarks>
' </remarks>
'\ingroup base
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/ApplicationHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/modErrorHandler.bas</use>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Public Enum ApplicationHandlerResumeModes
                               ' 0 ... keine Antwort
   AppResumeMode_Completed = 1 ' Anforderung wurde erfolgreich abgearbeitet
   AppResumeMode_Error = 2     ' Event wurde angenommen, aber es traten Fehler auf
   AppResumeMode_Cancel = 4    ' Event wurde angenommen, aber Weiterverabeitung soll gestoppt werden
End Enum

' Integrierte Erweiterungen
Private Const EXTENSION_KEY_AppFile As String = "AppFile"

' Fehlerkennungen
Private Const ERR_CLASS_ID As Long = 1000
Public Enum ApplicationHandlerErrors
   AppErr_EventInterfaceMissing = vbObjectError + ERR_CLASS_ID + 1

End Enum

' API-Funktionen
Private Declare Function FindWindowEx Lib "USER32.DLL" Alias "FindWindowExA" ( _
       ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
       ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
       
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long

Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" ( _
       ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
       
Private Declare Function RedrawWindow Lib "user32" ( _
       ByVal Hwnd As Long, lprcUpdate As Any, _
       ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Const GCL_HBRBACKGROUND As Long = -10
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_ERASE As Long = &H4

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Const SYSCOLOR_COLOR_APPWORKSPACE As Long = 12

' Hilfskonstanten
Private Const m_LOG_FILE As String = "log.txt"

Private Const m_DBPROPNAME_AppTitle As String = "AppTitle"
Private Const m_DBPROPNAME_AppIcon As String = "AppIcon"
Private Const m_DBPROPNAME_UseAppIconForFrmRpt As String = "UseAppIconForFrmRpt"
Private Const m_DBPROPNAME_StartUpForm As String = "StartUpForm"
Private Const m_DBPROPNAME_StartUpMenuBar As String = "StartUpMenuBar"


'Hilfsvariablen
Private m_AppDb As DAO.Database           ' Ersatz f�r CurrentDb bzw. CodeDb, siehe Prop: AppDb

Private m_ApplicationName As String       'Zwischenspeicher f�r Anwendungsnamen (kurz)
Private m_ApplicationFullName As String   'Zwischenspeicher f�r Anwendungsnamen (lang)

Private m_ApplicationVersion As String    'anzuzeigende Versionnummer

Private m_varTransferValue As Variant
Private m_PublicPath As String            ' Standardverzeichnis f�r Dateiexplorer

Private m_ApplicationStartFormName As String ' Formular das in Start-Methode ge�ffnet wird

Private m_CustomMDIBackgroundColor As Long

'Events
Public Event BeforeDispose(ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
Public Event AfterDispose(ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

Public Event CheckExtension(ByVal Key As String, ByRef exists As Boolean)

Public Event ExtensionLookup(ByVal Key As String, ByRef ExtensionReference As Object)

Public Event ExtensionInstance(ByVal ExtensionKey As String, ByVal InstanzKey As String, ByRef ExtensionReference As Object)

Public Event ExtensionPropertyLookup(ByVal Key As String, ByVal PropertyName As String, _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

Public Event ExtensionProcedureCall(ByVal Key As String, ByVal ProcedureName As String, _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant, _
            ByRef Params() As Variant)

Public Event AppFileBeforeCreateFile( _
            ByVal sFileID As String, ByVal sFileName As String, _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
            
Public Event AppFileBeforeSaveFile( _
            ByVal sFileID As String, ByVal sFileName As String, ByVal saveVersion As Boolean, _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
            
Public Event UpdateApplication( _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

Public Event CheckUpdate( _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

Public Event BeforeOpenStartForm( _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
            
Public Event BeforeStartApplication( _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

Public Event AfterStartApplication( _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

'---------------------------------------------------------------------------------------
' Dispose-Ablauf
'---------------------------------------------------------------------------------------
Private m_Disposed As Boolean
Private Sub Class_Terminate()
On Error Resume Next
   If Not m_Disposed Then
      Dispose
   End If
End Sub

Public Sub Dispose(Optional ByRef ResumeMode As ApplicationHandlerResumeModes, Optional ByRef ResumeMessage As Variant)
   
   Dim tempResumeMode As ApplicationHandlerResumeModes
   Dim tempResumeMessage As Variant
   
On Error Resume Next
   
   RaiseEvent BeforeDispose(tempResumeMode, tempResumeMessage)
   If tempResumeMode = AppResumeMode_Cancel Then
      ResumeMode = tempResumeMode
      ResumeMessage = tempResumeMessage
      Exit Sub
   End If
   
   Set m_AppDb = Nothing
   
   m_Disposed = True
   
   RaiseEvent AfterDispose(tempResumeMode, tempResumeMessage)
   
   ResumeMode = AppResumeMode_Completed
   ResumeMessage = tempResumeMessage
   
End Sub

'---------------------------------------------------------------------------------------
' Function: Start (Josef P�tzl, 2009-12-15)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Start der Anwendung
' </summary>
' <returns>Boolean</returns>
' <remarks>
' Nicht Verwechseln mit Initialisierung von ApplicationHandler \n
' Diese Prozedur l�st die Startroutine f�r den Anwender aus.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Start(Optional ByRef ResumeMessage As Variant) As Boolean
   
   Dim currentStartFormName As String
   Dim eventResumeMode As ApplicationHandlerResumeModes
   Dim eventResumeMessage As Variant
   
On Error GoTo HandleErr

   'Evente-Schnittstelle f�r Anwendungsstart
   ' ... erm�glicht z. B. das Durchf�hren eines Login-Vorgangs
   RaiseEvent BeforeStartApplication(eventResumeMode, eventResumeMessage)
   Select Case eventResumeMode
      Case ApplicationHandlerResumeModes.AppResumeMode_Cancel
         'Start abbrechen
         ResumeMessage = eventResumeMessage
         Start = False
         Exit Function
      Case ApplicationHandlerResumeModes.AppResumeMode_Completed
         'Startprozedur wurde von Erweiterung �bernommen
         Start = eventResumeMessage
         Exit Function
      Case ApplicationHandlerResumeModes.AppResumeMode_Error
         'Mit Fehlermeldung abbrechen
         Err.Raise vbObjectError, "BeforeStartApplication", eventResumeMessage
         Exit Function
      Case Else
         'weitermachen
   End Select

   'Update pr�fen
   If checkVersionUpdate Then
      If Me.UpdateApplication Then
         Start = False
         Application.Quit acQuitSaveNone
         Exit Function
      End If
   End If
   
   'Evente-Schnittstelle f�r Startformular
   '... erm�glich das �ndern des Startformulars durch ERweiterungen
   '    z. B. falls ein benutzerspezifisches Formulare ge�ffnet werden soll
   eventResumeMode = 0
   eventResumeMessage = Empty
   RaiseEvent BeforeOpenStartForm(eventResumeMode, eventResumeMessage)
   Select Case eventResumeMode
      Case ApplicationHandlerResumeModes.AppResumeMode_Cancel
         'Startform nicht �ffnen
         currentStartFormName = vbNullString
      Case ApplicationHandlerResumeModes.AppResumeMode_Completed
         'Name des StartFormulars wurde �bertragen
         currentStartFormName = Nz(eventResumeMessage, vbNullString)
      Case Else
         currentStartFormName = Me.ApplicationStartFormName
   End Select
   
   If Len(currentStartFormName) > 0 Then
      DoCmd.OpenForm currentStartFormName
   End If

   'Benachrichtigung �ber erfolgten Anwendungsstart
   eventResumeMode = 0
   eventResumeMessage = Empty
   RaiseEvent AfterStartApplication(eventResumeMode, eventResumeMessage)

   Start = True
   
ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandler.Start", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function


'---------------------------------------------------------------------------------------
' Property: ApplicationName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Name der aktuellen Anwendung
' </summary>
' <returns>String</returns>
' <remarks>
' Reihenfolge der Namensermittlung:
' 1. �ber gesetzten Wert
' 2. aus Titel-Eigenschaft mit CurrentDb.Properties("AppTitle")
' 3. aus Dateinamen mit CurrentDb.Name
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ApplicationName() As String
On Error Resume Next

   If Len(m_ApplicationName) = 0 Then 'Wert aus Titel-Eigenschaft, da Konstante nicht eingestellt wurde
      m_ApplicationName = AppDb.Properties("AppTitle").Value
      If Len(m_ApplicationName) = 0 Then 'Wert aus Dateinamen
         m_ApplicationName = AppDb.Name
         m_ApplicationName = Left$(m_ApplicationName, InStrRev(m_ApplicationName, ".") - 1)
      End If
   End If
   
   ApplicationName = m_ApplicationName
   
End Property

'---------------------------------------------------------------------------------------
' Property: ApplicationName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Name der aktuellen Anwendung einstellen
' </summary>
' <param name="appName">Anwendungsname</param>
'**/
'---------------------------------------------------------------------------------------
Public Property Let ApplicationName(ByVal appName As String)
On Error Resume Next
    m_ApplicationName = appName
End Property


'---------------------------------------------------------------------------------------
' Property: ApplicationFullName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Vollst�ndiger Name der aktuellen Anwendung
' </summary>
' <returns>String</returns>
' <remarks>
' Reihenfolge der Namensermittlung:
' 1. �ber gesetzten Wert
' 2. �ber ApplicationName
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ApplicationFullName() As String
On Error Resume Next

   If Len(m_ApplicationFullName) = 0 Then 'Kurzform verwenden
      m_ApplicationFullName = ApplicationName
   End If
   
   ApplicationFullName = m_ApplicationFullName
   
End Property

'---------------------------------------------------------------------------------------
' Property: ApplicationFullName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Name der aktuellen Anwendung einstellen
' </summary>
' <param name="appName">Anwendungsname</param>
'**/
'---------------------------------------------------------------------------------------
Public Property Let ApplicationFullName(ByVal appName As String)
On Error Resume Next
    m_ApplicationFullName = appName
End Property

'---------------------------------------------------------------------------------------
' Property: ApplicationVersion
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Vollst�ndiger Name der aktuellen Anwendung
' </summary>
' <returns>String</returns>
' <remarks>
' Reihenfolge der Namensermittlung:
' 1. �ber gesetzten Wert
' 2. �ber ApplicationName
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get Version() As String
On Error Resume Next

   If Len(m_ApplicationVersion) = 0 Then ' ... aus Properties lesen?
'/**
' @todo Versionskennung aus DB-Eigenschaften bzw. aus den Dateieigenschaften lesen
'**/
      m_ApplicationVersion = ""
   End If
   
   Version = m_ApplicationVersion
   
End Property

'---------------------------------------------------------------------------------------
' Property: ApplicationFullName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Name der aktuellen Anwendung einstellen
' </summary>
' <param name="appName">Anwendungsname</param>
'**/
'---------------------------------------------------------------------------------------
Public Property Let Version(ByVal appVersion As String)
On Error Resume Next
    m_ApplicationVersion = appVersion
End Property

'---------------------------------------------------------------------------------------
' Property: ApplicationStartForm
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Formular, das in der Methode Start ge�ffnet wird
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ApplicationStartFormName() As String
On Error Resume Next

   If Len(m_ApplicationStartFormName) = 0 Then ' ... aus Properties lesen?
      m_ApplicationStartFormName = ""
   End If
   
   ApplicationStartFormName = m_ApplicationStartFormName
   
End Property

'---------------------------------------------------------------------------------------
' Property: ApplicationStartForm
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Formular, das in der Methode Start ge�ffnet wird
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let ApplicationStartFormName(ByVal FormName As String)
On Error Resume Next
    m_ApplicationStartFormName = FormName
End Property


'---------------------------------------------------------------------------------------
' Property: AppDb (Josef P�tzl, 2009-12-13)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ersatz f�r CurrentDb bzw. CodeDb (Standard: CodeDb)
' </summary>
' <returns>zugewiesene DAO.Database-Instanz</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get AppDb() As DAO.Database

On Error GoTo HandleErr

   If m_AppDb Is Nothing Then
      Set m_AppDb = CodeDb
   End If
   Set AppDb = m_AppDb

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandler.AppDb", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Friend Property Set AppDb(ByRef db As DAO.Database)

On Error GoTo HandleErr

   Set m_AppDb = db

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandler.AppDb", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Public Function GetTransferValue(Optional ByVal vValueIfNull As Variant = Null, Optional ByVal bolNoDel As Boolean = False) As Variant

On Error GoTo HandleErr

   GetTransferValue = Nz(m_varTransferValue, vValueIfNull)
   If Not bolNoDel Then m_varTransferValue = Null

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandler.GetTransferValue", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

Public Sub SetTransferValue(ByVal newValue As Variant)

On Error GoTo HandleErr

   m_varTransferValue = newValue

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandler.SetTransferValue", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Public Property Get PublicPath() As String
On Error Resume Next
   PublicPath = m_PublicPath
End Property

Friend Property Let PublicPath(ByVal sPublicPath As String)
On Error Resume Next
   m_PublicPath = sPublicPath
End Property

Public Property Get ApplicationTitle() As String
On Error Resume Next
   ApplicationTitle = AppDb.Properties("AppTitle").Value
End Property

Friend Property Let ApplicationTitle(ByVal AppTitle As String)

On Error GoTo HandleErr

   If Len(AppTitle) > 0 Then
      setAppDbProperty m_DBPROPNAME_AppTitle, dbText, AppTitle
   Else
      deleteAppDbProperty m_DBPROPNAME_AppTitle
   End If
   Access.Application.RefreshTitleBar

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandler.ApplicationTitle", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Property

Friend Sub SetAppIcon(ByVal AppIconFile As String, Optional ByVal UseAppIconForFrmRpt As Boolean = True)

   Dim bolCheck As Boolean

On Error GoTo HandleErr

   If Len(Dir$(AppIconFile)) > 0 Then
      setAppDbProperty m_DBPROPNAME_AppIcon, dbText, AppIconFile
      setAppDbProperty m_DBPROPNAME_UseAppIconForFrmRpt, dbBoolean, UseAppIconForFrmRpt
   Else
      '1. Pr�fen ob AppFile-Erweiterung vorhanden ist
      RaiseEvent CheckExtension(EXTENSION_KEY_AppFile, bolCheck)
      If bolCheck Then
         bolCheck = CreateAppFile("AppIcon", AppIconFile)
      End If
      '2. Properties einstellen
      If bolCheck Then
         setAppDbProperty m_DBPROPNAME_AppIcon, dbText, AppIconFile
         setAppDbProperty m_DBPROPNAME_UseAppIconForFrmRpt, dbBoolean, UseAppIconForFrmRpt
      Else
         deleteAppDbProperty m_DBPROPNAME_AppIcon
      End If
   End If
   Application.RefreshTitleBar

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandler.SetAppIcon", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Friend Property Let StartUpForm(ByVal StartUpFormName As String)

On Error GoTo HandleErr

   If Len(StartUpFormName) > 0 Then
      setAppDbProperty m_DBPROPNAME_StartUpForm, dbText, "Form." & StartUpFormName
   Else
      deleteAppDbProperty m_DBPROPNAME_StartUpForm
   End If

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "StartUpForm", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Friend Property Let StartUpMenuBar(ByVal MenuBarName As String)

On Error GoTo HandleErr

   If Len(MenuBarName) > 0 Then
      setAppDbProperty m_DBPROPNAME_StartUpMenuBar, dbText, MenuBarName
   Else
      deleteAppDbProperty m_DBPROPNAME_StartUpMenuBar
   End If

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandler.StartUpMenuBar", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Public Sub SetApplicationProperty(ByVal propName As String, ByVal PropValue As Variant, Optional ByVal PropType As Long = dbText)
   setAppDbProperty propName, PropType, PropValue
End Sub

Private Sub setAppDbProperty(ByVal propName As String, ByVal PropType As Long, ByVal PropValue As Variant)
   
   Dim db As DAO.Database
   Dim propCol As DAO.Properties

On Error GoTo HandleErr

   Set db = AppDb
   Set propCol = db.Properties
   If dbPropertyExists(propCol, propName) Then
      db.Properties(propName) = PropValue
   Else
      db.Properties.Append db.CreateProperty(propName, PropType, PropValue)
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandler.setAppDbProperty", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub deleteAppDbProperty(ByVal sPropName As String)
   
   Dim db As DAO.Database
   Dim propCol As DAO.Properties
   
On Error GoTo HandleErr

   Set db = AppDb
   Set propCol = db.Properties
   
   If dbPropertyExists(propCol, sPropName) Then
      propCol.Delete sPropName
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandler.deleteAppDbProperty", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Function dbPropertyExists(ByRef propCol As DAO.Properties, ByVal sPropName As String) As Boolean
   
   Dim prop As DAO.Property
   
On Error GoTo HandleErr

   For Each prop In propCol
      If prop.Name = sPropName Then
         dbPropertyExists = True
         Exit For
      End If
   Next

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandler.dbPropertyExists", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

Public Sub WriteApplicationLogEntry(ByVal msg As String)

   Dim strLogFile As String
   Dim lfileNo As Long
   
On Error GoTo HandleErr

   strLogFile = CurrentProject.Path & "\" & m_LOG_FILE
   
   lfileNo = FreeFile
   Open strLogFile For Append As lfileNo
   Print #lfileNo, msg
   Close #lfileNo

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandler.WriteApplicationLogEntry", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Public Property Get MdiBackColor() As Long
'/**
' * @bug Ermitteln der Hintergrundfarbe des MDI-Bereichs Funktioniert nicht unter Access 2007
'**/
On Error GoTo HandleErr

   If m_CustomMDIBackgroundColor <> 0 Then
      MdiBackColor = m_CustomMDIBackgroundColor
   Else
      MdiBackColor = GetSysColor(SYSCOLOR_COLOR_APPWORKSPACE)
   End If

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "MdiBackColor", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Property

Friend Property Let MdiBackColor(ByVal Color As Long)
  
   Dim NewBrush As Long
   Dim hMDI As Long

On Error GoTo HandleErr

   m_CustomMDIBackgroundColor = Color
   hMDI = FindWindowEx(Application.hWndAccessApp, 0&, "MDIClient", vbNullString)
   
   'Brush erzeugen
   NewBrush = CreateSolidBrush(m_CustomMDIBackgroundColor)
   'Brush zuweisen
   SetClassLong hMDI, GCL_HBRBACKGROUND, NewBrush
   'Fenster neuzeichnen (gesamtes Fenster inkl. Background)
   RedrawWindow hMDI, ByVal 0&, ByVal 0&, RDW_INVALIDATE Or RDW_ERASE

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandler.MdiBackColor", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Property


'---------------------------------------------------------------------------------------
' Function: CreateAppFile
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Datei aus AppFile-Tabelle extrahieren
' </summary>
' <param name="sFileID">Kennung in Tabelle</param>
' <param name="sFileName"></param>
' <returns>True = Datei wurde gespeichert</returns>
' <remarks>
' Erfordert AppFile-Erweiterung (wird �ber Ereignis-Schnittstelle angesprochen)
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function CreateAppFile(ByVal sFileID As String, ByVal sFileName As String) As Boolean

   Dim eventResumeMode As ApplicationHandlerResumeModes ' R�ckgabewert aus Event-Schnittstelle
   Dim eventResumeMessage As Variant                    ' Nachricht, f�r Abbruchgrund (wird derzeit noch nicht genutzt)

On Error GoTo HandleErr

On Error Resume Next

   '---------------------------------------------------
   ' Schnittstelle f�r andere Klassen .. damit wird es m�glich,
   ' dass die Erstellung der Datei eine andere Klasse �bernimmt
   '
   eventResumeMode = 0
   RaiseEvent AppFileBeforeCreateFile(sFileID, sFileName, eventResumeMode, eventResumeMessage)
   If eventResumeMode <> 0 Then
      CreateAppFile = (eventResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Completed)
   Else
      Err.Raise ApplicationHandlerErrors.AppErr_EventInterfaceMissing, _
                "CreateAppFile", "Auf das AppFileBeforeCreateFile-Ereignis wurde nicht reagiert"
   End If
  

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandler.CreateAppFile", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function

'---------------------------------------------------------------------------------------
' Function: SaveAppFile
'---------------------------------------------------------------------------------------
'/
' <summary>
' Datei in AppFile-Tabelle speichern
' </summary>
' <param name="sFileID">Kennung in Tabelle (Feld "id")</param>
' <param name="sFileName">Dateiname</param>
' <param name="saveVersion">Version abspeichern</param>
' <remarks>
' Erfordert ApplicationHandler_File (wird �ber Ereignis-Schnittstelle angesprochen)
' </remarks>
'/
'---------------------------------------------------------------------------------------
Friend Function SaveAppFile(ByVal sFileID As String, ByVal sFileName As String, Optional ByVal saveVersion As Boolean = False) As Boolean

   Dim eventResumeMode As ApplicationHandlerResumeModes ' R�ckgabewert aus Event-Schnittstelle
   Dim eventResumeMessage As Variant                    ' Nachricht, f�r Abbruchgrund

On Error GoTo HandleErr

   '---------------------------------------------------
   ' Event-Schnittstelle AppFileBeforeSaveFile
   '
   eventResumeMode = 0
   RaiseEvent AppFileBeforeSaveFile(sFileID, sFileName, saveVersion, eventResumeMode, eventResumeMessage)
   If eventResumeMode <> 0 Then
      SaveAppFile = (eventResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Completed)
   Else
      Err.Raise ApplicationHandlerErrors.AppErr_EventInterfaceMissing, _
                "CreateAppFile", "Auf das AppFileBeforeSaveFile-Ereignis wurde nicht reagiert"
   End If

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandler.SaveAppFile", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function

Public Property Get Extensions(ByVal ExtensionKey As String) As Object
'Anm.: auf Auflistungs-Klasse verzichtet .. der Code ist zwar nicht so elegant, daf�r ist eine kaum verwendete Klasse weniger

'/**
' * @todo sollte man eine Schnittstelle statt Late binding verwenden?
'**/

   Dim extRef As Object
   
On Error GoTo HandleErr

   RaiseEvent ExtensionLookup(ExtensionKey, extRef)
   Set Extensions = extRef

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandler.Extensions", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Public Function GetExtensionInstance(ByVal ExtensionKey As String, Optional ByVal InstanceKey As String = vbNullString) As Object

   Dim extRef As Object

On Error GoTo HandleErr

   RaiseEvent ExtensionInstance(ExtensionKey, InstanceKey, extRef)
   Set GetExtensionInstance = extRef

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandler.GetExtension", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function

Public Function GetExtensionProperty(ByVal ExtensionKey As String, ByVal PropertyName As String, _
                            Optional ByVal ValueIfMissing As Variant = Empty) As Variant
   
   Dim eventResumeMode As ApplicationHandlerResumeModes
   Dim eventResumeMessage As Variant
   
On Error GoTo HandleErr

   RaiseEvent ExtensionPropertyLookup(ExtensionKey, PropertyName, eventResumeMode, eventResumeMessage)
   
   If eventResumeMode = AppResumeMode_Completed Then
      GetExtensionProperty = eventResumeMessage
   Else
      GetExtensionProperty = ValueIfMissing
   End If

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandler.GetExtensionProperty", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function


Public Function UpdateApplication() As Boolean
' True = Anwendung schlie�en
 
' => Ausgelager an Erweiterung: damit wird die Verwendugn flexibler
'    und jeder kann sich aussuchen nach welcher Update-Methode vorgegangen werden soll

   Dim eventResumeMode As ApplicationHandlerResumeModes
   Dim eventResumeMessage As Variant

On Error GoTo HandleErr

   RaiseEvent UpdateApplication(eventResumeMode, eventResumeMessage)
   If eventResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Completed Then
      UpdateApplication = eventResumeMessage
   Else
      UpdateApplication = False
   End If

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "UpdateApplication", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

Private Function checkVersionUpdate() As Boolean
' True = Update erforderlich
   
   Dim eventResumeMode As ApplicationHandlerResumeModes
   Dim eventResumeMessage As Variant

On Error GoTo HandleErr

   RaiseEvent CheckUpdate(eventResumeMode, eventResumeMessage)
   If eventResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Completed Then
      checkVersionUpdate = eventResumeMessage
   Else
      checkVersionUpdate = False
   End If

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "checkVersionUpdate", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

'---------------------------------------------------------------------------------------
' Sub: CallExtensionProcedure (Josef P�tzl, 2010-03-13)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Prozeduren aus Erweiterungen aufrufen
' </summary>
' <param name="Key">Extension-Kennung</param>
' <param name="ProcedureName">Name der zu startenden Prozedur</param>
' <param name="ResumeMode">R�ckgabekennung</param>
' <param name="ResumeMessage">Antwort</param>
' <param name="Params">Array f�r �bergabe- und R�ckgabeparameter</param>
' <returns></returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub CallExtensionProcedure(ByVal Key As String, ByVal ProcedureName As String, _
            ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant, _
            ByRef Params() As Variant)
            
On Error GoTo HandleErr

   RaiseEvent ExtensionProcedureCall(Key, ProcedureName, ResumeMode, ResumeMessage, Params)

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "CallExtensionProcedure", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
            
End Sub