Attribute VB_Name = "SqlTools_DotNetLib"
'---------------------------------------------------------------------------------------
' Package: data.SqlTools_DotNetLib
'---------------------------------------------------------------------------------------
'
' Factory module for the DotNetLib version of SqlTools
'
' Author:
'     Josef Poetzl
'
' Remarks:
'     AccessCodeLib.Data.SqlTools.interop
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/SqlTools_DotNetLib.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>_dotnetlib/integration/DotNetLibRepair.frm</use>
'  <test>_test/data/SqlTools_DotNetLibTests.cls</test>
'  <description>Factory for the .NET version of SqlTools (ACLib DotNetLib)</description>
'  <example>_examples/data/SqlTools_DotNetLib_Examples.bas</example>
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
