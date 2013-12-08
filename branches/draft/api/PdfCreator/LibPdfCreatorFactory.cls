VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibPdfCreatorFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: LibPdfCreatorFactory
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Factory für die Instance des PDF-Creator Prozesses
' \ingroup PdfCreator
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/PdfCreator/LibPdfCreatorFactory.cls</file>
'  <use>api/PdfCreator/LibPdfCreatorInstance.cls</use>
'  <test>_test/api/PdfCreator/LibPdfCreatorFactoryTests.cls</test>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Public Property Get PdfCreatorInstance() As LibPdfCreatorInstance
On Error GoTo Err_

    Dim CreatorInstance As New LibPdfCreatorInstance
        CreatorInstance.cStart ("/NoProcessingAtStartup")

    If CreatorInstance.cCountOfPrintjobs > 0 Then _
        Err.Raise LibPdfCreatorErrorNumer.WarteSchlangeNichtLeer, "LibPdfCreatorFactory", "Die Steuerung über die PdfCreator-Instanz kann nicht übernommen werden, " & _
                                                                  "es befinden sich noch nicht abgeschlossene Druckaufträge in der Warteschlange."

    'Handle Error Number: 2 -  An instance of the PDFCreator is already running. It is not allowed to have more than one instance running!
    If CreatorInstance.cError.Number = 2 Then
    
        'Nur dann die laufende Instanz von PdfCreator schließen, wenn dieser aktuell keine Jobs verarbeitet
          
        'TODO: Test auf Versionen <> 1.5 ob nicht doch auf Close verzichtet
        '      Und statt dessen einfach der Error-Clear ausreichend ist
        '    CreatorInstance.cClose
            
        'Set CreatorInstance = New LibPdfCreatorInstance
        '    CreatorInstance.cStart ("/NoProcessingAtStartup")
        'Statt dessen:
        CreatorInstance.cErrorClear
        
    End If
        
    Set PdfCreatorInstance = CreatorInstance
    Set CreatorInstance = Nothing
Err_:
    Select Case Err.Number
        Case 0:
        Case Else:
            GoSub Cleanup_
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End Select
Cleanup_:
    Set CreatorInstance = Nothing

End Property
