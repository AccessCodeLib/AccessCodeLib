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
' Class: data.DbConnectionSettings
'---------------------------------------------------------------------------------------
'
' Auxiliary class for caching the connection parameters
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

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

'SqlTools parameter

'---------------------------------------------------------------------------------------
' Property: SqlDateFormat
'---------------------------------------------------------------------------------------
'
' Date format string
'
'---------------------------------------------------------------------------------------
Public SqlDateFormat As String

'---------------------------------------------------------------------------------------
' Property: SqlWildCardString
'---------------------------------------------------------------------------------------
'
' SQL wildcard
'
'---------------------------------------------------------------------------------------
Public SqlWildCardString As String

'---------------------------------------------------------------------------------------
' Property: SqlBooleanTrueString
'---------------------------------------------------------------------------------------
'
' SQL string for boolean true
'
'---------------------------------------------------------------------------------------
Public SqlBooleanTrueString As String

'---------------------------------------------------------------------------------------
' Property: DefaultAdodbCursorLocation
'---------------------------------------------------------------------------------------
'
' Default parameter value for ADODB.CursorLocationEnum
'
'---------------------------------------------------------------------------------------
Public DefaultAdodbCursorLocation As ADODB.CursorLocationEnum

'---------------------------------------------------------------------------------------
' Property: OpenPermanentBackendRecordset
'---------------------------------------------------------------------------------------
'
' Open permanent backend recordset if connected
'
'---------------------------------------------------------------------------------------
Public OpenPermanentBackendRecordset As Boolean

'---------------------------------------------------------------------------------------
' Property: Clone
'---------------------------------------------------------------------------------------
'
' Clone of current config
'
'---------------------------------------------------------------------------------------
Public Function Clone() As DbConnectionSettings
   Dim NewObject As DbConnectionSettings
   Set NewObject = New DbConnectionSettings
   NewObject.CopyFrom Me
   Set Clone = NewObject
End Function

'---------------------------------------------------------------------------------------
' Property: CopyFrom
'---------------------------------------------------------------------------------------
'
' Copy DbConnectionSettings
'
' Parameters:
'     NewSettings - DbConnectionSettings to copy
'
'---------------------------------------------------------------------------------------
Public Sub CopyFrom(ByVal NewSettings As DbConnectionSettings)

   With NewSettings
      
      SqlDateFormat = .SqlDateFormat
      SqlWildCardString = .SqlWildCardString
      SqlBooleanTrueString = .SqlBooleanTrueString
      
      DefaultAdodbCursorLocation = .DefaultAdodbCursorLocation
      
      OpenPermanentBackendRecordset = .OpenPermanentBackendRecordset
      
   End With
   
End Sub
