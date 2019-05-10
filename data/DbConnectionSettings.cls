VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbConnectionSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: DbConnectionInfo
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hilfsklasse zum Zwischenspeichern der Verbindungsparameter
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/DbConnectionSettings.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>ADODB</name><major>2</major><minor>1</minor><guid>{00000201-0000-0010-8000-00AA006D2EA4}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'SqlTools-Parameter
Public SqlDateFormat As String
Public SqlWildCardString As String
Public SqlBooleanTrueString As String

'AdodbHandler-Parameter
Public DefaultAdodbCursorLocation As ADODB.CursorLocationEnum

'Datenzugriffssteuerung
Public OpenPermanentBackendRecordset As Boolean

Public Function Clone() As DbConnectionSettings
   Dim NewObject As DbConnectionSettings
   Set NewObject = New DbConnectionSettings
   NewObject.CopyFrom Me
   Set Clone = NewObject
End Function

Public Sub CopyFrom(ByVal NewSettings As DbConnectionSettings)

   With NewSettings
      
      SqlDateFormat = .SqlDateFormat
      SqlWildCardString = .SqlWildCardString
      SqlBooleanTrueString = .SqlBooleanTrueString
      
      DefaultAdodbCursorLocation = .DefaultAdodbCursorLocation
      
      OpenPermanentBackendRecordset = .OpenPermanentBackendRecordset
      
   End With
   
End Sub
