Attribute VB_Name = "modExcel_Tools"
'---------------------------------------------------------------------------------------
' Module: modExcel_Tools (Josef Pötzl)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hilfsfunktionen für Excel
' </summary>
' <remarks>
' </remarks>
'\ingroup api_office
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/office/excel/modExcel_Tools.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>base/modErrorHandler.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

#Const ExcelEarlyBinding = 0

Public Function RecordsetExcelExport(ByRef RecordsetReference As Object, _
                       Optional ByVal TemplateFile As String = vbNullString, _
                       Optional ByVal StartRow As Long = 1, Optional ByVal StartCol As Long = 1, _
                       Optional ByVal WithRecordsetHeaders As Boolean = True, _
                       Optional ByVal ImportTabName As String = vbNullString)
   
On Error GoTo HandleErr

#If ExcelEarlyBinding = 1 Then
   Dim xlApp As Excel.Application
   Dim xlWb As Excel.Workbook
   Dim xlSh As Excel.worksheet
#Else
   Dim xlApp As Object
   Dim xlWb As Object
   Dim xlSh As Object
#End If
  
   Set xlApp = CreateObject("Excel.Application")
   xlApp.Visible = True
   Set xlWb = xlApp.Workbooks.Add(TemplateFile)
   If Len(ImportTabName) > 0 Then
      Set xlSh = xlWb.Sheets(ImportTabName)
   Else
      Set xlSh = xlWb.Sheets(1)
   End If
   
   excelSheetCopyFromRecordset xlSh, RecordsetReference, StartRow, StartCol, WithRecordsetHeaders

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "RecordsetExcelExport", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

Private Sub excelSheetCopyFromRecordset(ByRef xlSheet As Object, _
                                        ByRef rstData As Object, _
                               Optional ByVal StartRowNr As Long = 1, _
                               Optional ByVal StartColNr As Long = 1, _
                               Optional ByVal WithRecordsetHeaders As Boolean = False)
                    
' xlSheet = Excel.worksheet

    Dim lngRstFields As Long
    Dim strRstFieldName() As String
    Dim i As Long
    
    'Überschriften
On Error GoTo HandleErr

    If WithRecordsetHeaders Then
        With rstData
            lngRstFields = .Fields.Count - 1
            ReDim strRstFieldName(lngRstFields)
            For i = 0 To lngRstFields
                strRstFieldName(i) = .Fields(i).Name
            Next i
            xlSheet.Range(xlSheet.Cells(StartRowNr, StartColNr), _
                          xlSheet.Cells(StartRowNr, StartColNr + lngRstFields) _
                         ).Value = strRstFieldName
            StartRowNr = StartRowNr + 1
        End With
    End If
    
    'Daten
    xlSheet.Cells(StartRowNr, StartColNr).CopyFromRecordset rstData

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "excelSheetCopyFromRecordset", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub
