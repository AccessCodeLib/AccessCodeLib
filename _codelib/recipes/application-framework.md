# Application start with the ApplicationHandler framework

Keywords: application framework, startup, AutoExec, application name, version, start form, extensions, settings,
login, version check, tracing, Anwendungsstart, Anwendungsrahmen, Startformular, Erweiterungen, Einstellungen

## Import

- `base/_initApplication.bas` (adds `base/modApplication.bas`, `base/ApplicationHandler.cls` and the configuration
  template `base/_config_Application.bas`)
- optional extensions, e.g. `base/ApplicationHandler_DbConnection.cls`, `base/ApplicationHandler_Settings.cls`,
  `base/ApplicationHandler_Version.cls`, `base/log/ApplicationHandler_Tracing.cls`,
  `base/login/ApplicationHandler_AppLogin_Win.cls`, `usability/ApplicationHandler_DirTextbox.cls`

## Usage

1. Adapt `_config_Application` (application name, version, start form, extensions):

```vba
Private Const m_ApplicationVersion As String = "1.0.0"
Private Const m_ApplicationName As String = "MyApp"
Private Const m_ApplicationStartFormName As String = "frmMain"

#Const USE_EXTENSIONS = True

Public Sub InitConfig(Optional oCurrentAppHandler As ApplicationHandler = Nothing)
   ' ...
   With oCurrentAppHandler
      .ApplicationName = m_ApplicationName
      .Version = m_ApplicationVersion
      .ApplicationStartFormName = m_ApplicationStartFormName
   End With

#If USE_EXTENSIONS = True Then
   Set m_Extensions = New ApplicationHandler_ExtensionCollection
   With m_Extensions
      Set .ApplicationHandler = oCurrentAppHandler
      .Add New ApplicationHandler_DbConnection
   End With
#End If
End Sub
```

2. Start the application, e.g. in the macro `AutoExec` with `RunCode StartApplication()`.

3. Use it in the application:

```vba
Debug.Print CurrentApplicationName
TraceLog "Import started", FileName          ' with ApplicationHandler_Tracing: message + optional values
```

## Notes

- `_config_Application` is a template: adapt it per application. It is stored in the application, not in the
  repository, and the import does not overwrite an existing module of this name.
- Extensions are activated in `InitConfig` via `ApplicationHandler_ExtensionCollection.Add`.
- `base/frmAppWatcher.frm` (hidden form) disposes the ApplicationHandler when the application is closed.

## See also

- [backend-connection.md](backend-connection.md) - DbConnection extension
- `_codelib/samples/` - sample configurations (e.g. with login or license check)
