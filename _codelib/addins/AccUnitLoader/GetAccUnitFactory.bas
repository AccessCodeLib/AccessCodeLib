Attribute VB_Name = "AccUnitLoaderFactoryCall"
'---------------------------------------------------------------------------------------
' Package: _codelib.addins.AccUnitLoader.AccUnitLoaderFactoryCall
'---------------------------------------------------------------------------------------
'
' Provides the AccUnitLoader factory (GetAccUnitFactory)
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/AccUnitLoader/GetAccUnitFactory.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/AccUnitLoader/AccUnitLoaderFactory.cls</use>
'</codelib>
'---
Option Compare Database
Option Explicit

Public Function GetAccUnitFactory() As AccUnitLoaderFactory
   CheckAccUnitTypeLibFile
   Set GetAccUnitFactory = New AccUnitLoaderFactory
End Function
