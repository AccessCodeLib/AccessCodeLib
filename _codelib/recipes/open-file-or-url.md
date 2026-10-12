# Open a file or URL, send mail, start a program

Keywords: ShellExecute, open file with associated application, open URL, browser, mailto, run as administrator,
start program and wait, Datei öffnen, Standardprogramm, Link öffnen, E-Mail, als Administrator ausführen, Programm starten

## Import

- `api/winapi/WinApiShellTools.cls`

## Usage

```vba
Dim Shell As WinApiShellTools
Set Shell = New WinApiShellTools

' File, folder or URL with the associated application
Shell.Execute "C:\Temp\report.pdf"
Shell.Execute "https://access-codelib.net"
Shell.Execute "C:\Temp\report.pdf", "print"            ' other verb of ShellExecute

' New mail in the default mail program (mailto)
Shell.ShellExecuteSendMail "info@example.com", "Subject", "Text"

' Start a program and wait until it has ended
Shell.LaunchAppSynchronous "notepad.exe", "C:\Temp\readme.txt"

' Run as administrator (UAC dialog)
Shell.ExecuteAsAdmin "cmd"
```

Or via the factory: `WinAPI.Shell.Execute "C:\Temp\report.pdf"` (import `api/winapi/WinApiToolsFactory.bas`).

## Notes

- All methods return `True` on success.
- `ShellExecuteSendMail` builds a `mailto:` link without URL encoding: special characters in subject or body
  (e.g. `&`, `?`, line breaks) must be encoded by the caller.
- `api/winapi/modWinAPI.bas` contains the same functions as standard module (older variant); prefer the class.
- `file/FileTools.bas` has `OpenFile` and `OpenFilePath` for files and folders.

## See also

- [file-dialogs-and-paths.md](file-dialogs-and-paths.md)
- `_examples/api/winapi/WinApiShellTools_Example.bas`
