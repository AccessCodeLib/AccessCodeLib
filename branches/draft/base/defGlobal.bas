Attribute VB_Name = "defGlobal"
Attribute VB_Description = "Allgemeine Konstanten und Eigenschaften"
'---------------------------------------------------------------------------------------
' Modul: defGlobal (2009-07-27)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Allgemeine Konstanten und Eigenschaften
' </summary>
' <remarks></remarks>
' \ingroup base
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/defGlobal.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>base/modApplication.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Explicit
Option Compare Text

'---------------------------------------------------------------------------------------
'
' Konstanten
'

' SQL-Konstanten

   Public Const SQL_DEFAULT_TEXTDELIMITER As String = "'"
   Public Const SQL_DEFAULT_DATEFORMAT As String = "\#yyyy\-mm\-dd\#"

'---------------------------------------------------------------------------------------
'
' Hilfs-Variablen
'


'---------------------------------------------------------------------------------------
'
' Hilfs-Prozeduren
'

'
' Private Hilfsvariablen für die Prozeduren
'
Private m_ApplicationName As String         'Zwischenspeicher für Anwendungsnamen, falls
                                            'CurrentApplication.ApplicationName nicht läuft

'---------------------------------------------------------------------------------------
' Property: CurrentApplicationName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Name der aktuellen Anwendung
' </summary>
' <returns>String</returns>
' <remarks>
' Verwendet CurrentApplication.ApplicationName
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get CurrentApplicationName() As String
' inkl. Notfall-Errorhandler, falls CurrentApplication nicht instanziert ist

On Error GoTo HandleErr

   CurrentApplicationName = CurrentApplication.ApplicationName
      
ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   CurrentApplicationName = getApplicationNameFromDb
   Resume ExitHere

End Property

Private Function getApplicationNameFromDb() As String

On Error Resume Next

   If Len(m_ApplicationName) = 0 Then 'Wert aus Titel-Eigenschaft, da Konstante nicht eingestellt wurde
      m_ApplicationName = CodeDb.Properties("AppTitle").Value
      If Len(m_ApplicationName) = 0 Then 'Wert aus Dateinamen
         m_ApplicationName = CodeDb.Name
         m_ApplicationName = Left$(m_ApplicationName, InStrRev(m_ApplicationName, ".") - 1)
      End If
   End If
   
   getApplicationNameFromDb = m_ApplicationName
   
End Function
