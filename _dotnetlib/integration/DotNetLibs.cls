VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DotNetLibs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: DotNetLibs
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' \brief        TODO
' \ingroup DotNetLib
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_dotnetlib/integration/DotNetLibs.cls<file>
'  <use>_dotnetlib/integration/DotNetLibsSetup.bas</use>
'  <use>COM/NetComDomain.cls</use>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private m_SqlTools As NetComDomain

Public Property Get LibPath() As String
   LibPath = LibFiles.LibPath
End Property

Public Property Get SqlTools() As NetComDomain
   If m_SqlTools Is Nothing Then
      Set m_SqlTools = New NetComDomain
      m_SqlTools.LibraryFile = LibPath & "AccessCodeLib.Data.SqlTools.interop.dll"
      m_SqlTools.DefaultNameSpace = "AccessCodeLib.Data.SqlTools.interop"
   End If
   Set SqlTools = m_SqlTools
End Property
