Attribute VB_Name = "SqlTools_DotNetLib"
'---------------------------------------------------------------------------------------
' Module: SqlToolsSetup_DotNetLib
'---------------------------------------------------------------------------------------
'/**
'\author     Josef Poetzl
' <summary>
' Factory-Modul für DotNetLib Version der SqlTools
' </summary>
' <remarks>
' AccessCodeLib.Data.SqlTools.interop
' </remarks>
'\ingroup data
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>/data/SqlTools_DotNetLib.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>_dotnetlib/integration/DotNetLibRepair.frm</use>
'  <test>_test/data/SqlTools_DotNetLibTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_SqlTools As ACLibSqlTools.SqlToolsFactory
Private m_Generator As ACLibSqlTools.SqlGenerator

Public Property Get SqlTools() As ACLibSqlTools.SqlToolsFactory
   If m_SqlTools Is Nothing Then
      Set m_SqlTools = DotNetLibs.SqlTools.CreateObject("SqlToolsFactory")
   End If
   Set SqlTools = m_SqlTools
End Property
